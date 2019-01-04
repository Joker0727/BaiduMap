using MyTools;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SqlLiteHelperDemo;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tool;

namespace BaiduMap
{
    public partial class Form1 : Form
    {
        public SeleniumHelper sel = null;
        public string basePath = AppDomain.CurrentDomain.BaseDirectory;
        public Thread thread = null;
        public Thread getCommentsthread = null;
        public string mainUrl = "https://map.baidu.com/";
        public bool isLogin = false;
        public int page = 0;
        public int current = 0;
        public MyUtils myUtils = null;
        private readonly string workId = "ww-0020";
        private string searchStr = string.Empty, commentStr = string.Empty, imgFolder = string.Empty;
        public bool isRandom = false;
        public List<string> imgPathList = new List<string>();
        public int fNumber = 0;
        public bool noPic = true;

        public string sqlitePath = AppDomain.CurrentDomain.BaseDirectory + @"sqlit3.db";
        public SQLiteHelper sqlLiteHelper = null;


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Control.CheckForIllegalCrossThreadCalls = false;//屏蔽跨线程操作ui控件的异常        
            myUtils = new MyUtils();
            this.MaximizeBox = false;
        }
        private void textBox3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            this.textBox3.Text = path.SelectedPath;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            noPic = this.checkBox2.Checked;
            string btnStr = this.button1.Text;
            searchStr = this.textBox1.Text;
            //commentStr = this.textBox2.Text;
            imgFolder = this.textBox3.Text;
            isRandom = this.checkBox1.Checked;
            fNumber = Convert.ToInt32(this.numericUpDown1.Text);
            if (string.IsNullOrEmpty(searchStr))
            {
                MessageBox.Show("搜索关键词词不能为空！", "BaiduMap");
                return;
            }
            //if (string.IsNullOrEmpty(imgFolder))
            //{
            //    MessageBox.Show("图片路径不能为空！", "BaiduMap");
            //    return;
            //}
            //if (!Directory.Exists(imgFolder))
            //{
            //    MessageBox.Show("图片路径不存在！", "BaiduMap");
            //    return;
            //}
            //if (commentStr.Length < 15)
            //{
            //    MessageBox.Show("评论内容不能小于15个字！", "BaiduMap");
            //    return;
            //}
            if (!IsAuthorised())
            {
                MessageBox.Show("尚未授权！", "BaiduMap");
                return;
            }
            if (btnStr == "开始")
            {
                this.button1.Text = "暂停";
                if (thread == null)
                {
                    sel = new SeleniumHelper();
                    sel.driver.Navigate().GoToUrl(mainUrl);
                    if (sel.driver.WindowHandles.Count() > 1)
                    {
                        sel.driver.SwitchTo().Window(sel.driver.WindowHandles[1]);
                        sel.driver.Close();
                        sel.driver.SwitchTo().Window(sel.driver.WindowHandles[0]);
                    }
                    MessageBoxButtons message = MessageBoxButtons.OKCancel;
                    DialogResult dr = MessageBox.Show("请先登录成功后，再点击确定！", "BaiduMap", message);
                    if (dr == DialogResult.OK)
                    {
                        imgPathList = myUtils.GetImgs(imgFolder, 2);
                        thread = new Thread(StartWork);
                        thread.IsBackground = true;
                        thread.Start();
                    }
                    else
                    {
                        MessageBox.Show("请退出程序重新你登陆！", "BaiduMap");
                        return;
                    }
                }
                else
                {
                    thread.Resume();
                }
            }
            else
            {
                this.button1.Text = "开始";
                thread?.Suspend();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            getCommentsthread = new Thread(GetComments);
            getCommentsthread.IsBackground = true;
            getCommentsthread.Start();
        }

        public void StartWork()
        {
            page = 1;
            current = 0;
            bool isRefresh = false;
            string shopName = string.Empty;
            ListViewItem lt = null;
            GETSHOPLIST:
            {
                if (current >= 10)
                {
                    page++;
                    current = 0;
                }
                IWebElement soleInput = sel.FindElementById("sole-input");
                soleInput.Clear();
                soleInput.SendKeys(searchStr);
                Thread.Sleep(1000 * 1);

                IWebElement searchButton = sel.FindElementById("search-button");
                //移动光标到指定的元素上perform
                Actions action = new Actions(sel.driver);
                action.MoveToElement(searchButton).Perform();
                searchButton.Click();
                Thread.Sleep(1000 * 1);
                bool isOk = false;
                while (!isOk)
                {
                    IWebElement curPageNode = sel.FindElementByClassName("curPage");
                    if (Convert.ToInt32(curPageNode?.Text) == page) break;
                    var spanNodeList = sel.FindElementsByXPath("//div[@id='poi_page']/p[@class='page']/span");
                    foreach (var nextNode in spanNodeList)
                    {
                        if (Regex.IsMatch(nextNode?.Text, @"^[+-]?\d*[.]?\d*$"))
                        {
                            int pageNumber = Convert.ToInt32(nextNode?.Text);
                            if (pageNumber == page)
                            {
                                nextNode?.Click();
                                isOk = true;
                                break;
                            }
                        }
                    }
                    if (!isOk && spanNodeList.Count > 0)
                    {
                        spanNodeList[spanNodeList.Count - 1]?.Click();
                    }
                }
            }

            Thread.Sleep(1000 * 1);
            var shopNodeList = sel.FindElementsByCss(".search-item.base-item");
            try
            {
                IWebElement shopNameNode = shopNodeList[current].FindElement(By.ClassName("n-blue"));
                shopName = shopNameNode?.Text;
                shopNodeList[current].Click();
                Thread.Sleep(1000 * 2);
                IWebElement writeTextNode = sel.FindElementByClassName("write-btn-text");
                writeTextNode.Click();
                Thread.Sleep(1000 * 1);
                string newWindows = sel.driver.WindowHandles[1];
                sel.driver.SwitchTo().Window(newWindows);
                Thread.Sleep(1000 * 2);

                REFRESH:
                {
                    if (isRefresh)
                    {
                        ((IJavaScriptExecutor)sel.driver).ExecuteScript("location.reload()");//刷新
                        Thread.Sleep(1000 * 2);
                        isRefresh = false;
                    }
                }
                var startNodeList = sel.FindElementsByXPath("//ul[@id='data-star']/li");
                if (startNodeList != null)
                {
                    int nodeCount = startNodeList.Count();
                    for (int i = 0; i < nodeCount; i++)
                    {
                        Thread.Sleep(30);
                        Actions actions = new Actions(sel.driver);
                        actions.MoveToElement(startNodeList[i]).Perform();
                        Thread.Sleep(30);
                        if (i == nodeCount - 1)
                            startNodeList[nodeCount - 1].Click();
                    }
                }
                else
                {
                    isRefresh = true;
                    goto REFRESH;
                }
                var facilitiesNodeList = sel.FindElementsByXPath("//ul[@id='7']/li");
                if (facilitiesNodeList != null)
                {
                    int nodeCount = facilitiesNodeList.Count();
                    for (int i = 0; i < nodeCount; i++)
                    {
                        Thread.Sleep(30);
                        Actions actions = new Actions(sel.driver);
                        actions.MoveToElement(facilitiesNodeList[i]).Perform();
                        Thread.Sleep(30);
                        if (i == nodeCount - 1)
                            facilitiesNodeList[nodeCount - 1].Click();
                    }
                }
                var environmentalNodeList = sel.FindElementsByXPath("//ul[@id='8']/li");
                if (environmentalNodeList != null)
                {
                    int nodeCount = environmentalNodeList.Count();
                    for (int i = 0; i < nodeCount; i++)
                    {
                        Thread.Sleep(30);
                        Actions actions = new Actions(sel.driver);
                        actions.MoveToElement(environmentalNodeList[i]).Perform();
                        Thread.Sleep(30);
                        if (i == nodeCount - 1)
                            environmentalNodeList[nodeCount - 1].Click();
                    }
                }
                var servicesNodeList = sel.FindElementsByXPath("//ul[@id='9']/li");
                if (servicesNodeList != null)
                {
                    int nodeCount = servicesNodeList.Count();
                    for (int i = 0; i < nodeCount; i++)
                    {
                        Thread.Sleep(30);
                        Actions actions = new Actions(sel.driver);
                        actions.MoveToElement(servicesNodeList[i]).Perform();
                        Thread.Sleep(30);
                        if (i == nodeCount - 1)
                            servicesNodeList[nodeCount - 1].Click();
                    }
                }
                IWebElement commentInputNode = sel.FindElementById("remark-text");
                if (commentInputNode != null)
                {
                    commentInputNode.SendKeys(commentStr);
                    Thread.Sleep(1000 * 1);
                    //上传图片                   
                    if (!noPic)
                        UpLoadPic();
                    IWebElement submitBtnNode = sel.FindElementByCss(".submit-btn.state-finish");
                    if (submitBtnNode == null)
                        submitBtnNode = sel.FindElementByCss(".submit-btn.disabled.state-pedding");
                    submitBtnNode.Click();
                }
                Thread.Sleep(1000 * 2);
                sel.driver.Close();
                string oldWindows = sel.driver.WindowHandles[0];
                sel.driver.SwitchTo().Window(oldWindows);
                current++;
                lt = new ListViewItem();
                lt.Text = shopName;
                lt.SubItems.Add("发布成功");
                this.listView1.Items.Add(lt);
                Thread.Sleep(1000 * 2);
                goto GETSHOPLIST;
            }
            catch (Exception ex)
            {
                current++;
                lt = new ListViewItem();
                lt.Text = shopName;
                lt.SubItems.Add("发布失败");
                this.listView1.Items.Add(lt);
                myUtils.WriteLog(ex);
                goto GETSHOPLIST;
            }
        }

        public void UpLoadPic()
        {
            int count = 0;
            if (isRandom)
                count = GetRandom();
            else
                count = fNumber > 0 ? fNumber : 9;
            string tempPath = string.Empty;
            for (int i = 0; i < count; i++)
            {
                try
                {
                    tempPath = imgPathList[i];
                    if (!File.Exists(tempPath))
                        continue;
                    // 上传图片
                    IWebElement addPicBtnNode = sel.FindElementByName("file");
                    addPicBtnNode.SendKeys(tempPath);
                    Thread.Sleep(1000 * 2);
                }
                catch (Exception ex)
                {
                    myUtils.WriteLog("图片上传失败：" + ex);
                }
            }
        }

        public int GetRandom()
        {
            int imgCount = imgPathList.Count();
            Random randObj = new Random();
            int start = 1;//随机数可取该下界值
            int end = imgCount + 1;//随机数不能取该上界值
            int random = randObj.Next(start, end);
            return random;
        }

        /// <summary>
        /// 获取评论
        /// </summary>
        public void GetComments()
        {
            sqlLiteHelper = new SQLiteHelper(sqlitePath);
            string meituanUrl = "https://bj.meituan.com/s/ktv/";
            searchStr = this.textBox1.Text;
            if (string.IsNullOrEmpty(searchStr))
            {
                MessageBox.Show("搜索关键词词不能为空！", "BaiduMap");
                return;
            }
            sel = new SeleniumHelper(1);
            sel.driver.Navigate().GoToUrl(meituanUrl);

            IWebElement searchInput = sel.FindElementByClassName("header-search-input");
            searchInput.Clear();
            searchInput.SendKeys(searchStr);
            Thread.Sleep(500);
            IWebElement searchBtn = sel.FindElementByClassName("header-search-btn");
            searchBtn.Click();
            //var shopNodeList = sel.FindElementsByXPath("//div[@class='common-list-main']/div");
            var shopNodeList = sel.FindElementsByCss(".link.list-item-pic.backup-color");
            string shopUrl = string.Empty, userCommentStr = string.Empty, sqlStr = string.Empty;
            foreach (var shopNode in shopNodeList)
            {
                try
                {
                    shopUrl = shopNode.GetAttribute("href");
                    sel.driver.Navigate().GoToUrl(shopUrl);
                    Thread.Sleep(1000 * 3);
                    while (true)
                    {
                        IWebElement currentPageNode = sel.FindElementByCss(".pagination-item.select.num-item");
                        var pageliaNodeList = sel.FindElementsByXPath("//ul[@class='clearfix']/li/a");
                        int pageNodeCount = pageliaNodeList.Count();


                        var userCommentNodeList = sel.FindElementsByClassName("user-comment-inner");
                        foreach (var userCommentNode in userCommentNodeList)
                        {
                            try
                            {
                                userCommentStr = userCommentNode.Text;
                                if (userCommentStr.Length > 15 & userCommentStr.Length < 2000)
                                {
                                    sqlStr = $"insert into CommentsTable (Comments,IsPublish)values('{userCommentStr}',0)";
                                    sqlLiteHelper.RunSql(sqlStr);
                                }
                            }
                            catch (Exception ex)
                            {
                                myUtils.WriteLog("保存评论失败：" + ex);
                            }

                        }

                        if (currentPageNode.Text == pageliaNodeList[pageNodeCount - 2].Text)
                            break;
                        else
                        {
                            pageliaNodeList[pageNodeCount - 1].Click();
                            Thread.Sleep(1000 * 2);
                        }
                    }
                }
                catch (Exception e)
                {
                    myUtils.WriteLog("进入店铺获取用户评论失败：" + e);
                }
            }
        }

        /// <summary>
        /// 授权
        /// </summary>
        /// <param name="workId"></param>
        /// <returns></returns>
        public bool IsAuthorised()
        {
            string conStr = "Server=111.230.149.80;DataBase=MyDB;uid=sa;pwd=1add1&one";
            bool bo = false;
            try
            {
                using (SqlConnection con = new SqlConnection(conStr))
                {
                    string sql = string.Format("select count(*) from MyWork Where IsAuth = 1 and WorkId ='{0}'", workId);
                    using (SqlCommand cmd = new SqlCommand(sql, con))
                    {
                        con.Open();
                        int count = int.Parse(cmd.ExecuteScalar().ToString());
                        if (count > 0)
                            bo = true;
                    }
                }
            }
            catch (Exception)
            {
                bo = false;
            }

            return bo;
        }
    }
}
