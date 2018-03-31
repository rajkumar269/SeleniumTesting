using System;
using System.Data;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
//Added by David For mouse hover action
using OpenQA.Selenium.Interactions;
using System.Windows.Forms;

using System.Threading;

namespace Setup
{
    [TestFixture(typeof(FirefoxDriver))]
    //[TestFixture(typeof(InternetExplorerDriver))]
    //[TestFixture(typeof(ChromeDriver))]
    //[TestFixture(typeof(SafariDriver))]
    public class WebAccessTestCasesWithMultipleBrowsers<TWebDriver> where TWebDriver : IWebDriver, new()
    {

        private IWebDriver driver;
        public static string sInternetExplorerPath = AppDomain.CurrentDomain.BaseDirectory + @"IEDriver\";
        public static string sChromePath = AppDomain.CurrentDomain.BaseDirectory + @"ChromeDriver\";        
        string CurrentBrowser;        
        int ColumnNo;
        string ImageURL;
        string ErrorText;
        string TestCaseId;
        string ParentWindow;

        [Test]
        public void TestCase()
        {            
            try
            {
                this.driver = new TWebDriver();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Initializing Driver.....");
                if (ex.ToString().Contains("InternetExplorer"))
                {
                    InternetExplorerOptions options = new InternetExplorerOptions();
                    //options.AddAdditionalCapability("version", "11");
                    options.IgnoreZoomLevel = true;
                    this.driver = new InternetExplorerDriver(sInternetExplorerPath, options);
                }
                if (ex.ToString().Contains("Chrome"))
                {
                    ChromeOptions options = new ChromeOptions();
                    options.AddArgument("test-type");
                    options.AddArguments("start-maximized");
                    this.driver = new ChromeDriver(sChromePath, options);
                }
            }
            finally
            {                
                    driver.Manage().Window.Maximize();             
            }
            ICapabilities capabilities = ((RemoteWebDriver)driver).Capabilities;
            SetupClass.BrowserName = capabilities.BrowserName;
            TimeSpan timespan = TimeSpan.FromSeconds(60);
            int timesp = 50;
            foreach (DataTable dt in SetupClass.DS.Tables)
            {
                if (dt.TableName.Contains("Sheet1"))
                {
                    continue;
                }
                int rowCount = dt.Rows.Count;
                int collCount = dt.Columns.Count;

                for (int i = 0; i < rowCount; i++)
                {
                    Boolean TestCase = false;
                    Boolean Developer = false;
                    DateTime StartTestCase;
                    DateTime EndTestCase;
                    try
                    {
                        StartTestCase = DateTime.Now;
                        string childURL = dt.Rows[i][0].ToString().Trim();
                        if (dt.Rows[i][1].ToString().Trim() != "")
                        {
                            TestCaseId = dt.Rows[i][1].ToString();

                            if (childURL != "")
                            {
                                WebDriverWait wait = new WebDriverWait(driver, timespan);
                                string FinalUrl = SetupClass.masterPath + childURL;
                                Console.WriteLine(FinalUrl);
                                driver.Navigate().GoToUrl(FinalUrl);
                                Thread.Sleep(timesp);                                 
                            }
                            string ElementParameter;
                            string ElementAction;
                            string ElementAssert;

                            ElementParameter = dt.Rows[i].ItemArray[3].ToString().Trim();

                            Thread.Sleep(timesp);                            
                            ColumnNo = 3;
                            if (ElementParameter != "")
                            {
                                IWebElement WebElement = null;

                                //Get the Web element
                                Thread.Sleep(timesp);

                                //Created by David to fetch the element 
                                WebElement = ReturnWebElement(ElementParameter);

                                Thread.Sleep(timesp);

                                //Perform Action on web element
                                ElementAction = dt.Rows[i].ItemArray[4].ToString().Trim();

                                Thread.Sleep(timesp);

                                //Created by David to prform the action in the element
                                Developer = PerformActionOnElement(WebElement, ElementAction);
                                SetupClass.Outputds.Tables[dt.TableName].Rows[i][collCount] = (Developer ? "Pass" : "Fail");

                                Thread.Sleep(timesp);
                            }

                            ElementAssert = dt.Rows[i].ItemArray[5].ToString().Trim();

                            if (ElementAssert.Contains("Body:")) //Verify the action yields the expected result
                            {
                                Thread.Sleep(timesp);

                                //Check the assert Value
                                TestCase = CheckAssertValue(ElementAssert);
                                SetupClass.Outputds.Tables[dt.TableName].Rows[i][collCount + 1] = (TestCase ? "Pass" : "Fail");                                
                            }

                            EndTestCase = DateTime.Now;
                            String LoadTimeInSec = "Time Consumed - " + EndTestCase.Subtract(StartTestCase).TotalSeconds.ToString("#.##") + "Sec";
                            if (SetupClass.BrowserName == "internet explorer")
                            {
                                SetupClass.Outputds.Tables[dt.TableName].Rows[i][collCount + 2] = LoadTimeInSec;
                            }
                            else if (SetupClass.BrowserName == "firefox")
                            {
                                SetupClass.Outputds.Tables[dt.TableName].Rows[i][collCount + 3] = LoadTimeInSec;
                            }
                            else if (SetupClass.BrowserName == "chrome")
                            {
                                SetupClass.Outputds.Tables[dt.TableName].Rows[i][collCount + 4] = LoadTimeInSec;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            Console.WriteLine(TestCaseId + " - " + dt.Rows[i].ItemArray[ColumnNo].ToString().Trim());

                            CurrentBrowser = SetupClass.BrowserName;

                            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
                            DateTime today = DateTime.Today;
                            string sToday = today.ToString("MM-dd-yyyy");

                            ImageURL = TestCaseId + "_" + sToday + "_" + SetupClass.BrowserName + ".jpg";
                            ErrorText = ex.Message;

                            if (SetupClass.BrowserName == "internet explorer")
                            {
                                //ErrorValueIE = 1;
                                SetupClass.Outputds.Tables[dt.TableName].Rows[i][collCount + 2] = ErrorText + " - " + ImageURL;

                            }
                            else if (SetupClass.BrowserName == "firefox")
                            {
                                //ErrorValueMozilla = 1;
                                SetupClass.Outputds.Tables[dt.TableName].Rows[i][collCount + 3] = ErrorText + " - " + ImageURL;

                            }
                            else if (SetupClass.BrowserName == "chrome")
                            {
                                //ErrorValueChrome = 1;
                                SetupClass.Outputds.Tables[dt.TableName].Rows[i][collCount + 4] = ErrorText + " - " + ImageURL;

                            }

                            SetupClass.Outputds.Tables[dt.TableName].Rows[i][collCount + 1] = "Fail";

                            ss.SaveAsFile(SetupClass.sScreenShotPath + TestCaseId + "_" + sToday + "_" + SetupClass.BrowserName + ".jpg", System.Drawing.Imaging.ImageFormat.Gif);
                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine("Something went wrong with the diver capabilities...\r\n" + Ex);
                        }

                    }
                }
            }
            driver.Quit();
        }
        public IWebElement ReturnWebElement(string ElementParameter)
        {
            IWebElement WebElement = null;
            TimeSpan timesp = TimeSpan.FromSeconds(15);
            WebDriverWait wait = new WebDriverWait(driver, timesp);

            string[] PageElement = ElementParameter.Split('=');
            if (PageElement[0] == "name")
            {
                wait.Until(ExpectedConditions.ElementExists(By.Name("" + PageElement[1] + "")));
                WebElement = driver.FindElement(By.Name("" + PageElement[1] + ""));
            }
            else if (PageElement[0] == "id")
            {
                wait.Until(ExpectedConditions.ElementExists(By.Id("" + PageElement[1] + "")));
                WebElement = driver.FindElement(By.Id("" + PageElement[1] + ""));
            }
            else if (PageElement[0] == "css")
            {
                wait.Until(ExpectedConditions.ElementExists(By.CssSelector("" + PageElement[1] + "")));
                WebElement = driver.FindElement(By.CssSelector("" + PageElement[1] + ""));
            }
            else if (PageElement[0] == "xpath")
            {
                wait.Until(ExpectedConditions.ElementExists(By.XPath("" + PageElement[1] + "")));
                WebElement = driver.FindElement(By.XPath("" + PageElement[1] + ""));
            }
            else if (PageElement[0] == "link")
            {
                wait.Until(ExpectedConditions.ElementExists(By.LinkText("" + PageElement[1] + "")));
                WebElement = driver.FindElement(By.LinkText("" + PageElement[1] + ""));
            }
            else if (PageElement[0] == "tag")
            {
                wait.Until(ExpectedConditions.ElementExists(By.TagName("" + PageElement[1] + "")));
                WebElement = driver.FindElement(By.TagName("" + PageElement[1] + ""));
            }
            else if (PageElement[0] == "class")
            {
                wait.Until(ExpectedConditions.ElementExists(By.ClassName("" + PageElement[1] + "")));
                WebElement = driver.FindElement(By.ClassName("" + PageElement[1] + ""));
            }
            return WebElement;
        }     


        public Boolean PerformActionOnElement(IWebElement WebElement, string ElementAction)
        {
            //Console.WriteLine("--" + ElementParameter);
            string[] PageElement = ElementAction.Split('=');
            switch (PageElement[0])
            {
                case "SendKeys":
                    if (PageElement[1] == "Tab")
                    {
                        WebElement.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                    else if (PageElement[1] == "Enter")
                    {
                        WebElement.SendKeys(OpenQA.Selenium.Keys.Enter);
                    }
                    else
                    {
                        WebElement.Clear();
                        WebElement.SendKeys("" + PageElement[1] + "");
                    }
                    break;
                case "select":
                    int number;
                    if (Int32.TryParse(PageElement[1], out number))
                    {
                        new SelectElement(WebElement).SelectByIndex(number);
                    }
                    else
                    {
                        new SelectElement(WebElement).SelectByText("" + PageElement[1] + "");
                    }
                    break;
                case "click":
                    WebElement.Click();
                    //((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete");
                    break;
                case "clickPopUp":
                    ParentWindow = driver.CurrentWindowHandle;
                    PopupWindowFinder finder = new PopupWindowFinder(driver);
                    string newHandle = finder.Click(WebElement);
                    driver.SwitchTo().Window(newHandle);
                    break;
                case "closePopup":
                    if (ParentWindow != driver.CurrentWindowHandle)
                    {
                        driver.Close();
                        driver.SwitchTo().Window(ParentWindow);
                    }
                    break;
                case "switch":
                    //if (ParentWindow != driver.CurrentWindowHandle)
                    //{
                    driver.SwitchTo().Window(ParentWindow);
                    //}
                    
                    break;
                case "alert":
                    Thread.Sleep(3000);
                    if (PageElement[1] == "accept")
                    {
                        driver.SwitchTo().Alert().Accept();
                    }
                    else
                    {
                        driver.SwitchTo().Alert().Dismiss();
                    }
                    Thread.Sleep(1000);
                    break;
                case "hover":
                    Actions action = new Actions(driver);
                    action.MoveToElement(WebElement).Perform();
                    break;
                case "scroll":
                    //((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(0, document.body.scrollHeight - 150)");
                    ((IJavaScriptExecutor)driver).ExecuteScript("setInterval(function () { window.scroll(0, document.height); }, 100)");
                    break;
                case "Wait":
                    Thread.Sleep(1000);
                    break;
                case "FileUpload":
                    WebElement.Click();
                    Thread.Sleep(2000);
                    SendKeys.SendWait(PageElement[1]);
                    SendKeys.SendWait(@"{Enter}");
                    break;
                case "type":
                    if (WebElement.GetAttribute("type") != PageElement[1])    
                        return false;
                    break;
                case "width":
                    if (WebElement.GetAttribute("width") != PageElement[1])
                        return false;
                    break;
                case "height":
                    if (WebElement.GetAttribute("height") != PageElement[1])
                        return false;
                    break;
                //Textbox
                case "Required":
                    WebElement.Clear();
                    driver.FindElement(By.CssSelector("input[value='"+ PageElement[1] +"']")).Click();
                    return true;                                   
                case "Numeric":
                    WebElement.SendKeys("1234");
                    driver.FindElement(By.CssSelector("input[value='" + PageElement[1] + "']")).Click();
                    //The page is refreshed and the elements cannot be idenfied after submit    
                    //WebElement.Submit();
                    break;
                case "AlphaNumeric":
                    WebElement.SendKeys("Abc@123");
                    driver.FindElement(By.CssSelector("input[value='" + PageElement[1] + "']")).Click();
                    return true;   
                    //The page is refreshed and the elements cannot be idenfied after submit    
                    //WebElement.Submit();                    
                //Dropdown
                case "RequireSelection":
                    new SelectElement(WebElement).SelectByIndex(0);
                    driver.FindElement(By.CssSelector("input[value='" + PageElement[1] + "']")).Click();
                    return true;    
            }
            return true;
        }
        public Boolean CheckAssertValue(string ElementParameter)
        {

            string[] AssertValue = ElementParameter.Split(':');
            TimeSpan timesp = TimeSpan.FromSeconds(15);
            WebDriverWait wait = new WebDriverWait(driver, timesp);
            wait.Until(ExpectedConditions.ElementExists(By.TagName("" + AssertValue[0] + "")));
            IWebElement ElementName = driver.FindElement(By.TagName("" + AssertValue[0] + ""));
            //Console.WriteLine(ElementName.Text);
            if (!ElementName.Text.Contains(AssertValue[1]))
            {
                //throw new Exception("The Assert value didn't match");
                return false;
            }
            return true;
        }        
    }
}
