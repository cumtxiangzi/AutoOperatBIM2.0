using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

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
                string username = "1739";
                string password = "110122227";
                string project_id = "1007";
                string workshop_id = "919";

                if (driver == null)
                {
                    //谷歌浏览器
                    ChromeOptions opt = new ChromeOptions();
                    opt.AddExcludedArgument("enable-automation"); //关闭安全提示
                    opt.AddAdditionalOption("useAutomationExtension", false);
                    opt.AddArgument("--start-maximized"); //启动即最大化
                    opt.AddArgument("--disable-popup-blocking"); //禁用弹出拦截
                    opt.AddArgument("no-sandbox"); //关闭沙盘
                    opt.AddArgument("disable-extensions"); //扩展插件检测
                    opt.AddArgument("no-default-browser-check"); //默认浏览器检测
                    var chromeDriverService = ChromeDriverService.CreateDefaultService();
                    chromeDriverService.HideCommandPromptWindow = true;//关闭黑窗口
                    driver = new ChromeDriver(chromeDriverService, opt);
                }
                if (index == 0) //开始运行
                {
                    run0(url, username, password);
                }
                else if (index == 1)
                {
                    run1(url, username, password, project_id, workshop_id);
                }
                else if (index==20)
                {
                    run20(url2, username, password);
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
