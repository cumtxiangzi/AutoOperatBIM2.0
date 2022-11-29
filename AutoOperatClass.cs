using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace AutoOperateBIM2._0
{
    public class AutoOperatClass
    {
        /// <summary>
        /// 是否正在运行
        /// </summary>
        private bool running = false;

        /// <summary>
        /// 驱动
        /// </summary>
        private IWebDriver driver = null;

        /// <summary>
        /// 构造函数
        /// </summary>
        public AutoOperatClass()
        {

        }
        /// <summary>
        /// 开始运行
        /// </summary>
        public void startRun(int index)
        {
            //"""运行起来"""
            try
            {
                running = true;
                string url = "http://bim.sinoma-tianjin.com:8080/Login";
                string url2 = "https://bim.sinoma-tianjin.com/";
                string url3 = "http://39.106.162.133:9005/Login?sourceurl=http://bim.sinoma-tianjin.com:7788/#/";
                string username = "1739";
                string password = "110122227";
                string project_id = "1007";
                string workshop_id = "919";
                string checkPeople = "魏祺祥";
                string filePath = @"D:\项目设计 模型\1092-都匀上封\801-总降\1092-811-成品\DWG\1092-811-WD-带签名.zip";

                if (driver == null)
                {
                    //谷歌浏览器
                    //ChromeOptions opt = new ChromeOptions();
                    //opt.AddExcludedArgument("enable-automation"); //关闭安全提示
                    //opt.AddAdditionalOption("useAutomationExtension", false);
                    //opt.AddArgument("--start-maximized"); //启动即最大化
                    //opt.AddArgument("--disable-popup-blocking"); //禁用弹出拦截
                    //opt.AddArgument("no-sandbox"); //关闭沙盘
                    //opt.AddArgument("disable-extensions"); //扩展插件检测
                    //opt.AddArgument("no-default-browser-check"); //默认浏览器检测
                    //var chromeDriverService = ChromeDriverService.CreateDefaultService();
                    //chromeDriverService.HideCommandPromptWindow = true;//关闭黑窗口
                    //driver = new ChromeDriver(chromeDriverService, opt);

                    //微软浏览器
                    EdgeOptions opt = new EdgeOptions();
                    opt.AddExcludedArgument("enable-automation"); //关闭安全提示
                    opt.AddAdditionalOption("useAutomationExtension", false);
                    opt.AddArgument("--start-maximized"); //启动即最大化
                    opt.AddArgument("--disable-popup-blocking"); //禁用弹出拦截
                    opt.AddArgument("no-sandbox"); //关闭沙盘
                    opt.AddArgument("disable-extensions"); //扩展插件检测
                    opt.AddArgument("no-default-browser-check"); //默认浏览器检测
                    EdgeDriverService service = EdgeDriverService.CreateDefaultService(@".\", "msedgedriver.exe");
                    service.HideCommandPromptWindow = true;//关闭黑窗口
                    driver = new EdgeDriver(service, opt);

                }
                if (index == 0) //开始运行
                {
                    run0(url, username, password);
                }
                else if (index == 1)
                {
                    run1(url, username, password, project_id, workshop_id);
                }
                else if (index == 20)
                {
                    run20(url2, username, password);
                }
                else if (index == 30)
                {
                    run30(url3, username, password, checkPeople,filePath);
                }

            }
            catch (Exception)
            {

            }
        }

        /// <summary>
        /// 运行
        /// </summary>
        /// <param name="username"></param>
        /// <param name="password"></param>
        private void run0(string url, string accountName, string passWord)
        {
            //# 访问网站
            driver.Navigate().GoToUrl(url);
            //#开始进行操作
            try
            {
                driver.Navigate().GoToUrl(url);

                var account = driver.FindElement(By.Name("Account"));
                account.Clear();
                account.SendKeys(accountName);
                Thread.Sleep(500);

                var code = driver.FindElement(By.Name("Password"));
                code.Clear();
                code.SendKeys(passWord);
                Thread.Sleep(500);

                var okButton = driver.FindElement(By.Id("loginbutton"));
                okButton.Click();
                Thread.Sleep(500);

            }
            catch (Exception)
            {

            }
        }
        private void run1(string url, string accountName, string passWord, string projectCode, string workShopCode)
        {
            //# 访问网站
            driver.Navigate().GoToUrl(url);
            //#开始进行操作
            try
            {
                driver.Navigate().GoToUrl(url);

                var account = driver.FindElement(By.Name("Account"));
                account.Clear();
                account.SendKeys(accountName);
                Thread.Sleep(500);

                var code = driver.FindElement(By.Name("Password"));
                code.Clear();
                code.SendKeys(passWord);
                Thread.Sleep(500);

                var okButton = driver.FindElement(By.Id("loginbutton"));
                okButton.Click();
                Thread.Sleep(500);

                var designManage = driver.FindElement(By.PartialLinkText("设计管理"));
                designManage.Click();
                Thread.Sleep(500);

                var designApprove = driver.FindElements(By.ClassName("menulistdiv"))[2]; //点击设计审批信息，准备发起设计校审
                designApprove.Click();
                Thread.Sleep(500);

                var projectReview = driver.FindElement(By.XPath("/html/body/div[2]/table/tbody/tr/td[1]/table/tbody/tr[2]/td/div/div/div/div[5]")); //点击互提资料单
                projectReview.Click();
                Thread.Sleep(500);

                var iframe = driver.FindElement(By.XPath("/html/body/div[2]/table/tbody/tr/td[3]/div[2]/div[3]/iframe")); //转到iframe
                var projectReviewButton = driver.SwitchTo().Frame(iframe).FindElement(By.CssSelector("#form1 > div.toolbar > a")); //点击互提资料单按钮弹出对话框
                projectReviewButton.Click();
                Thread.Sleep(500);

                var num = driver.WindowHandles;
                var newWindow = driver.SwitchTo().Window(num[1]); //进入到新打开的页面

                var projectList = driver.FindElement(By.XPath("/html/body/form/div[3]/table/tbody/tr[1]/td[2]/select[1]")); //点击项目号下拉列表
                projectList.Click();
                var selectList = new SelectElement(projectList).Options; //选择项目编号
                foreach (var item in selectList)
                {
                    if (item.Text == projectCode)
                    {
                        item.Click();
                        projectList.Click();
                        break;
                    }
                }
                Thread.Sleep(500);

                var workShopList = driver.FindElement(By.XPath("/html/body/form/div[3]/table/tbody/tr[1]/td[2]/select[2]")); //点击车间编号下拉列表
                workShopList.Click();
                var selectList2 = new SelectElement(workShopList).Options; //选择子项车间代号
                foreach (var item in selectList2)
                {
                    if (item.Text.Contains(workShopCode))
                    {
                        item.Click();
                        workShopList.Click();
                        break;
                    }
                }
                Thread.Sleep(500);

                var sheetTitle = driver.FindElement(By.XPath("/html/body/form/div[3]/table/tbody/tr[2]/td[2]/input")); //填写项目校审单信息
                sheetTitle.Click();
                sheetTitle.SendKeys("资料校审" + "-" + DateTime.Now.Year.ToString() + "." + DateTime.Now.Month.ToString() + "." + DateTime.Now.Day.ToString());
                Thread.Sleep(500);

                var designType = driver.FindElement(By.XPath("/html/body/form/div[3]/table/tbody/tr[6]/td[2]/input[2]")); //选择设计阶段为中间阶段
                designType.Click();
                Thread.Sleep(500);

                if (workShopCode.Contains("91") || workShopCode.Contains("711")) //选择校审类型
                {
                    var checkType = driver.FindElement(By.XPath("/html/body/form/div[3]/table/tbody/tr[7]/td[2]/input[1]"));
                    checkType.Click();
                    Thread.Sleep(500);
                }
                else
                {
                    var checkType = driver.FindElement(By.XPath("/html/body/form/div[3]/table/tbody/tr[7]/td[2]/input[2]"));
                    checkType.Click();
                    Thread.Sleep(500);
                }

                var contentIframe = driver.FindElement(By.XPath("/html/body/form/div[3]/table/tbody/tr[9]/td[2]/div/div/div[2]/iframe")); //转到iframe
                var content = driver.SwitchTo().Frame(contentIframe).FindElement(By.XPath("/html/body/p")); //填写校审内容
                content.Click();
                content.SendKeys("给排水资料校审,设备表详见Vault");
                Thread.Sleep(500);

                //巨坑 进入iframe后记得及时退出
                var sendSheet = driver.SwitchTo().ParentFrame().FindElement(By.XPath("/html/body/form/div[1]/div/a[6]")); //发送校审单
                sendSheet.Click();
                Thread.Sleep(500);

                //var wa = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                var checkPeopleIframe = driver.SwitchTo().ActiveElement(); //转到iframe
                var checkPeopleSelect = driver.SwitchTo().Frame(checkPeopleIframe).FindElement(By.XPath("/html/body/table/tbody/tr/td/fieldset/table/tbody/tr[2]/td/input[3]")); //选择校审人员      
                checkPeopleSelect.Click();
                Thread.Sleep(500);

                driver.SwitchTo().ParentFrame(); //退出iframe
                var checkPeopleIframe1 = driver.SwitchTo().ActiveElement(); //转到iframe
                var checkPeopleText = driver.SwitchTo().Frame(checkPeopleIframe1).FindElement(By.XPath("/html/body/table/tbody/tr/td[1]/div[1]/input[1]")); //选择组织机构成员      
                checkPeopleText.Click();
                checkPeopleText.SendKeys("魏祺祥");
                var checkButton = driver.FindElement(By.XPath("/html/body/table/tbody/tr/td[1]/div[1]/input[2]")); //单击查询
                checkButton.Click();
                //特别注意此处要先停留一段时间让网页加载，否则无法定位到元素
                Thread.Sleep(500);
                var peopleName = driver.FindElement(By.XPath("/html/body/table/tbody/tr/td[1]/div[2]/div/div/ul/li[3]/span")); //选择人员
                peopleName.Click();
                var selectButton = driver.FindElement(By.XPath("/html/body/table/tbody/tr/td[2]/div[3]/button"));
                selectButton.Click();
                var confirmButton = driver.FindElement(By.XPath("/html/body/table/tbody/tr/td[2]/div[5]/button"));
                confirmButton.Click();
                Thread.Sleep(500);

                driver.SwitchTo().ParentFrame(); //退出iframe
                var checkPeopleIframe2 = driver.SwitchTo().Frame(checkPeopleIframe); //转到iframe
                var confirmButton1 = driver.FindElement(By.XPath("/html/body/div/input[2]")); //点击确定发送联系单 记得改为确定，目前是取消按钮
                confirmButton1.Click();
                Thread.Sleep(500);
            }
            catch (Exception)
            {

            }
        }
        private void run20(string url, string accountName, string passWord)
        {
            //# 访问网站
            driver.Navigate().GoToUrl(url);
            //#开始进行操作
            try
            {
                driver.Navigate().GoToUrl(url);

                var account = driver.FindElement(By.Name("Account"));
                account.Clear();
                account.SendKeys(accountName);
                Thread.Sleep(500);

                var code = driver.FindElement(By.Name("Password"));
                code.Clear();
                code.SendKeys(passWord);
                Thread.Sleep(500);

                var okButton = driver.FindElement(By.Id("loginbutton"));
                okButton.Click();
                Thread.Sleep(500);

            }
            catch (Exception)
            {

            }
        }
        private void run30(string url, string accountName, string passWord, string checkPeople,string filePath)
        {
            //# 访问网站
            driver.Navigate().GoToUrl(url);
            //#开始进行操作
            try
            {
                driver.Navigate().GoToUrl(url);
                Thread.Sleep(1000);

                var account = driver.FindElement(By.XPath("/html/body/div/div/div/div/form/div[1]/div[1]/div/div[1]/div/input")); //通过浏览器调试程序，选择某个元素后，通过右键复制完整XPATH路径获取
                account.Clear();
                account.SendKeys(accountName);
                Thread.Sleep(500);

                var code = driver.FindElement(By.XPath("/html/body/div/div/div/div/form/div[2]/div[1]/div/div[1]/div/input"));
                code.Clear();
                code.SendKeys(passWord);
                Thread.Sleep(1000);

                var okButton = driver.FindElement(By.XPath("/html/body/div/div/div/div/form/div[3]/div[1]/button"));
                okButton.Click();
                Thread.Sleep(1000);

                var alert = driver.SwitchTo().Alert(); //登录成功弹窗处理
                if (alert != null)
                {
                    alert.Accept();
                }
                Thread.Sleep(1000);

                //var selectList = new SelectElement(projectClass).Options; //选择项目编号
                //foreach (var item in selectList)
                //{
                //    if (item.Text == "市场项目")
                //    {
                //        item.Click();
                //        projectClass.Click();
                //        break;
                //    }
                //}
                //Thread.Sleep(500);

                var selectProject = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/div[2]/div/button")); //切换项目
                selectProject.SendKeys(Keys.Enter);
                Thread.Sleep(1000);

                var projectList = driver.FindElements(By.ClassName("el-tree-node__label")); //获取全部项目
                foreach (var item in projectList)//选择指定项目
                {
                    if (item.Text.Contains("1092"))
                    {
                        item.Click();
                        break;
                    }
                }
                Thread.Sleep(1000);

                var changeProject = driver.FindElement(By.XPath("/html/body/div[2]/div/div[3]/button[2]")); //切换项目确定
                changeProject.SendKeys(Keys.Enter);
                Thread.Sleep(1000);

                var qualityManage = driver.FindElement(By.XPath("/html/body/div/div/div/div[1]/div[1]/div/ul/li[4]/div")); //点击质量管理
                qualityManage.Click();
                Thread.Sleep(500);

                var majorCheck = driver.FindElement(By.XPath("/html/body/div/div/div/div[1]/div[1]/div/ul/li[4]/ul/li[2]")); //点击专业校审
                majorCheck.Click();
                Thread.Sleep(1500);

                var workshopList = driver.FindElements(By.ClassName("el-tree-node__content")); //点击需要校审的车间
                var startFlow = driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div[2]/div/section/main/form/div/div[4]/div/button[3]")); //点击流程发起
                foreach (var item in workshopList)
                {
                    if (item.Text.Contains("833") && item.Text.Contains("施工图"))
                    {
                        item.Click();
                        Thread.Sleep(500);
                        startFlow.Click();
                        break;
                    }
                }
                Thread.Sleep(1000);

                var materialName = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div[2]/div/div/section/div[1]/main[1]/div[2]/div/form/div[4]/div[1]/div/div/div/input")); //点击资料名称
                materialName.SendKeys("给排水校审资料");
                Thread.Sleep(500);

                var designState = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div[2]/div/div/section/div[1]/main[1]/div[2]/div/form/div[4]/div[2]/div/div/div/div/input"));//点击设计阶段
                designState.Click();
                Thread.Sleep(500);

                var designStataList = driver.FindElements(By.ClassName("el-select-dropdown__item"));
                var saveButton = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div[2]/div/div/section/header/div/div[2]/button[2]"));
                foreach (var item in designStataList)
                {
                    if (item.Text.Contains("施工图"))
                    {
                        item.Click();//点击施工图
                        Thread.Sleep(500);
                        saveButton.Click();//点击保存
                        break;
                    }
                }
                Thread.Sleep(500);

                //var drawingSelect = driver.FindElement(By.XPath("/html/body/div[1]/div/div[1]/div[2]/div[2]/div/div/section/div[1]/main[1]/div[2]/div/form/div[7]/div[2]/div[3]/table/tbody/tr[2]/td[6]/div/div/div[1]/div/div/button"));
                //drawingSelect.Click();
                //Thread.Sleep(500);

                var fileSelect = driver.FindElement(By.ClassName("el-upload__input")); //图纸上传
                fileSelect.SendKeys(filePath);
                Thread.Sleep(500);

                var fileUpload = driver.FindElement(By.XPath("/html/body/div[1]/div/div[1]/div[2]/div[2]/div/div/section/div[1]/main[1]/div[2]/div/form/div[7]/div[2]/div[3]/table/tbody/tr[2]/td/button[1]"));
                fileUpload.SendKeys(Keys.Enter);
                Thread.Sleep(500);

                var sendFlow = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div[2]/div/div/section/header/div/div[2]/button[3]"));
                sendFlow.Click();
                Thread.Sleep(500);

                //var selectRadioButton = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div[2]/div/div/section/div[2]/div/div/div[2]/div[1]/div/label/span[1]/span"));
                //if (selectRadioButton.Selected == false)
                //{
                //    selectRadioButton.Click();
                //}
                //Thread.Sleep(500);

                var checkPeopleName = driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[2]/div[2]/div/div/section/div[2]/div/div/div[2]/div[1]/div/label/span[2]/div/div/div/div[1]"));
                if (!checkPeopleName.Text.Contains(checkPeople))
                {
                    checkPeopleName.Click();
                    Thread.Sleep(500);

                    var searchWindow = driver.FindElement(By.CssSelector("body > div.el-select-dropdown.el-popper.is-multiple > div.el-scrollbar > div.el-select-dropdown__wrap.el-scrollbar__wrap >" +
                        " ul > li > div.el-input.el-input--small.el-input-group.el-input-group--append.el-input--suffix > input"));//此处巨坑只能使用CSS选择，其他不行
                    searchWindow.Click();
                    searchWindow.SendKeys(checkPeople);
                    Thread.Sleep(500);

                    var searchButton = driver.FindElement(By.CssSelector("body > div.el-select-dropdown.el-popper.is-multiple > div.el-scrollbar > div.el-select-dropdown__wrap.el-scrollbar__wrap > " +
                        "ul > li > div.el-input.el-input--small.el-input-group.el-input-group--append.el-input--suffix > div > button"));//此处巨坑只能使用CSS选择，其他不行
                    searchButton.Click();
                    Thread.Sleep(500);

                    var selectPeople = driver.FindElement(By.ClassName("el-tree-node__content")); //此处巨坑要先获得树节点对象，然后再树节点对象中查找人名前的选择框
                    if (selectPeople.Text.Contains(checkPeople))
                    {
                        var selectCheckBox = selectPeople.FindElement(By.ClassName("el-checkbox__input")); //人名前的选择框选取
                        selectCheckBox.Click();
                        Thread.Sleep(500);
                    }

                    var selectOkButton = driver.FindElement(By.CssSelector("body > div.el-select-dropdown.el-popper.is-multiple > div.el-scrollbar > div.el-select-dropdown__wrap.el-scrollbar__wrap >" +
                        " ul > div > button.el-button.el-button--primary.el-button--small.is-plain")) ;
                    selectOkButton.Click();
                    Thread.Sleep(500);
                }
            }
            catch (Exception)
            {

            }
        }

        /// <summary>
        /// 停止运行
        /// </summary>
        public void stopRun()
        {
            //"""停止"""
            try
            {
                running = false;
                if (driver != null)
                {
                    driver.Quit();
                    //# 关闭后要为None，否则启动报错
                    driver = null;
                }
            }
            catch (Exception)
            {

            }
            finally
            {
                driver = null;
            }
        }
    }
}
