using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using MathNet.Numerics.Random;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager.Helpers;
using static Fable.Core.Testing;
using static System.Net.Mime.MediaTypeNames;

namespace SeleniumTest
{
    internal class Program
    {
        private static object workbook;
        private static string t;

        static void Main(string[] args)
        {
            //Console.Write("Test case started");
            ////create the reference for the browser

            ////IWebDriver driver = new EdgeDriver();

            ////new WebDriverManager.DriverManager().SetUpDriver(new ChromeConfig());
            ////ChromeDriver driver = new ChromeDriver();

            ////ChromeDriver driver = new ChromeDriver("webdriver.exe file path");

            ////new WebDriverManager.DriverManager().SetUpDriver(new EdgeConfig());

            //new DriverManager().SetUpDriver(new EdgeConfig(), VersionResolveStrategy.MatchingBrowser);

            //EdgeDriver driver = new EdgeDriver();


            ////navigate to url
            //driver.Navigate().GoToUrl("https://www.google.com/");
            //Thread.Sleep(2000);
            ////Identify the google search text box
            //IWebElement ele = driver.FindElement(By.Name("q"));
            ////enter the value in the google search text box
            //ele.SendKeys("javatpoint tutorials");
            //Thread.Sleep(2000);
            ////Identify the google search button
            //IWebElement ele1 = driver.FindElement(By.Name("btnK"));
            ////click on the google search button
            //ele1.Click();
            //Thread.Sleep(3000);
            ////close the browser
            //driver.Close();
            //Console.Write("test case ended");























            //Console.Write("Test case started");
            //new DriverManager().SetUpDriver(new EdgeConfig(), VersionResolveStrategy.MatchingBrowser);
            //EdgeDriver driver = new EdgeDriver();


            //driver.Navigate().GoToUrl("https://www.google.com/");
            //Thread.Sleep(2000);
            ////maximize the browser
            //driver.Manage().Window.Maximize();

            //IWebElement ele = driver.FindElement(By.Name("q"));

            //ele.SendKeys("w3 schools");
            //Thread.Sleep(2000);

            //IWebElement ele1 = driver.FindElement(By.Name("btnK"));
            //ele1.Click();
            //Thread.Sleep(3000);

            //IWebElement ele2 = driver.FindElement(By.XPath("//*[@id=\"rso\"]/div[1]/div/div/div/div/div/div/div/div[1]/div/span/a/h3"));
            //ele2.Click();
            //Thread.Sleep(3000);

            //IWebElement ele3 = driver.FindElement(By.XPath("//*[@id=\"subtopnav\"]/a[12]"));
            //ele3.Click();
            //Thread.Sleep(3000);
            //string extractText = string.Empty;
            //IWebElement ele4 = driver.FindElement(By.XPath("//*[@id=\"main\"]/h1"));
            //if (ele4.Displayed && ele4.Text != "")
            //{
            //    //Console.WriteLine(ele4.Text);
            //    extractText = ele4.Text;
            //}


            ////IWebElement ele5 = driver.FindElement(By.XPath("//*[@id=\"main\"]/div[3]"));
            //////Console.WriteLine(ele5.Text);
            ////string extractText1 = ele5.Text;


            //IWebElement ele5 = driver.FindElement(By.XPath("//*[@id=\"main\"]/div[3]/h2"));
            ////Console.WriteLine(ele5.Text);
            //string extractText1 = ele5.Text;

            //IWebElement ele6 = driver.FindElement(By.XPath("//*[@id=\"main\"]/div[3]/p[1]"));
            ////Console.WriteLine(ele5.Text);
            //string extractText2 = ele6.Text;

            //IWebElement ele7 = driver.FindElement(By.XPath("//*[@id=\"main\"]/div[3]/p[2]"));
            ////Console.WriteLine(ele5.Text);
            //string extractText3 = ele7.Text;


            //IWebElement ele8 = driver.FindElement(By.XPath("//*[@id=\"main\"]/div[3]/a"));
            ////Console.WriteLine(ele5.Text);
            //string extractText4 = ele8.Text;


            ////List<string> list = new List<string>();
            ////list.Add(extractText);
            ////list.Add(extractText1);
            ////list.Add(extractText2);
            ////list.Add(extractText3);
            ////list.Add(extractText4);



            //using (var workbook = new XLWorkbook())
            //{
            //    var worksheet = workbook.Worksheets.Add("Extracted data");
            //    worksheet.Cell(1, 1).Value = extractText;
            //    worksheet.Cell(2, 1).Value = extractText1;
            //    worksheet.Cell(3, 1).Value = extractText2;
            //    worksheet.Cell(4, 1).Value = extractText3;
            //    worksheet.Cell(5, 1).Value = extractText4;
            //    workbook.SaveAs(@"C:\Users\Nikitha.Thota\source\repos\output.xlsx");

            //}
            ////workbook.SaveAs("extracteddata.xlsx");
            ////workbook.SaveAs("Output.xlsx");

            //driver.Close();
            //driver.Quit();





























            Console.Write("Test case started");
            new DriverManager().SetUpDriver(new EdgeConfig(), VersionResolveStrategy.MatchingBrowser);
            EdgeDriver driver = new EdgeDriver();
            

            driver.Navigate().GoToUrl("https://ui.vision/demo/webtest/frames/");
            //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            Thread.Sleep(2000);
            //maximize the browser
            driver.Manage().Window.Maximize();
            Thread.Sleep(2000);


            driver.SwitchTo().Frame(driver.FindElement(By.XPath("/html/frameset/frame[1]")));
            IWebElement frame1 = driver.FindElement(By.XPath("//*[@id=\"id1\"]/div/input"));
            //frame1.Click();
            frame1.SendKeys("frame1 input");
            Thread.Sleep(2000);


            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(driver.FindElement(By.XPath("/html/frameset/frameset/frame[1]")));
            IWebElement frame2 = driver.FindElement(By.XPath("//*[@id=\"id2\"]/div/input"));
            //frame1.Click();
            frame2.SendKeys("frame2 input");
            Thread.Sleep(2000);




            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(driver.FindElement(By.XPath("/html/frameset/frameset/frame[2]")));
            IWebElement frame3 = driver.FindElement(By.XPath("//*[@id=\"id3\"]/div/input"));
            //frame1.Click();
            frame3.SendKeys("frame3 input");
            Thread.Sleep(2000);

            //driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(driver.FindElement(By.XPath("/html/body/center/iframe")));
            //CustomMethods.Click(driver, By.Id("//*[@id=\"i8\"]/div[3]"));
            IWebElement radioButton = driver.FindElement(By.XPath("//*[@id=\"i8\"]/div[3]"));
            radioButton.Click();
            Thread.Sleep(2000);
            IWebElement checkBox1 = driver.FindElement(By.XPath("//*[@id=\"i19\"]/div[2]"));
            checkBox1.Click();
            Thread.Sleep(2000);
            IWebElement checkBox2 = driver.FindElement(By.XPath("//*[@id=\"i22\"]/div[2]"));
            checkBox2.Click();
            Thread.Sleep(2000);
            IWebElement dropDown = driver.FindElement(By.XPath("//*[@id=\"mG61Hd\"]/div[2]/div[1]/div[2]/div[3]/div/div/div[2]/div/div[1]/div[2]"));
            dropDown.Click();
            Thread.Sleep(2000);
            IWebElement dropDownSelect = driver.FindElement(By.XPath("//*[@id=\"mG61Hd\"]/div[2]/div[1]/div[2]/div[3]/div/div/div[2]/div/div[2]/div[3]/span"));
            dropDownSelect.Click();
            Thread.Sleep(2000);
            IWebElement next = driver.FindElement(By.XPath("//*[@id=\"mG61Hd\"]/div[2]/div[1]/div[3]/div[1]/div[1]/div/span/span"));
            next.Click(); 
            Thread.Sleep(2000);
            IWebElement shortText = driver.FindElement(By.XPath("//*[@id=\"mG61Hd\"]/div[2]/div[1]/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input"));
            shortText.SendKeys("Hy There!");
            Thread.Sleep(2000);
            IWebElement longText = driver.FindElement(By.XPath("//*[@id=\"mG61Hd\"]/div[2]/div[1]/div[2]/div[3]/div/div/div[2]/div/div[1]/div[2]/textarea"));
            longText.SendKeys("welcome to data marshall");
            Thread.Sleep(2000);
            IWebElement submit = driver.FindElement(By.XPath("//*[@id=\"mG61Hd\"]/div[2]/div[1]/div[3]/div[1]/div[1]/div[2]/span/span"));
            submit.Click();
            Thread.Sleep(2000);

            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(driver.FindElement(By.XPath("/html/frameset/frameset/frame[3]")));
            IWebElement frame4 = driver.FindElement(By.Name("mytext4"));
            //frame1.Click();
            frame4.SendKeys("frame4 input");
            Thread.Sleep(2000);




            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame(driver.FindElement(By.XPath("/html/frameset/frame[2]")));
            IWebElement frame5 = driver.FindElement(By.Name("mytext5"));
            //frame1.Click();
            frame5.SendKeys("frame5 input");
            Thread.Sleep(2000);

            //IWebElement link = driver.FindElement(By.LinkText("https://a9t9.com"));
            //link.Click();
            //Thread.Sleep(2000);


            //driver.Close();
            //driver.Quit();


            //Actions actions = new Actions(driver);
            //actions.KeyDown(Keys.Control).SendKeys("t").KeyUp(Keys.Control).Perform();



            //driver.Navigate().GoToUrl(@"");
            // A new window is opened and switches to it
            driver.SwitchTo().NewWindow(WindowType.Window);
            // Loads Sauce Labs open source website in the newly opened window
            driver.Navigate().GoToUrl(@"https://a9t9.com");


            //Actions builder = new Actions(driver);
            //builder.KeyDown(Keys.Control).KeyDown(Keys.Tab).KeyUp(Keys.Tab).KeyUp(Keys.Control);//switch tabs
            //IAction switchTabs = builder.Build();
            //switchTabs.Perform();


            ////right click
            //Actions actions = new Actions(driver);
            //IWebElement elementLocator = (IWebElement)driver.FindElement(By.LinkText("https://a9t9.com"));
            //actions.ContextClick(elementLocator).Perform();



            ////switch with the help of actions class 
            //Actions action = new Actions(driver); 
            //action.KeyDown(Keys.Control).SendKeys(Keys.Tab).Build().Perform(); //opening the URL saved. 
            //driver.GetType("https://a9t9.com");


            //List<IWebElement> checkboxes = driver.FindElements(By.CssSelector(".myCheckboxClass"));
            //IWebElement randomCheckbox = checkboxes.GetType(new Random().NextInt64(List.Size()));







            //IWebElement radioButton = driver.FindElement(By.XPath("//*[@id=\"i8\"]/div[3]/div"));
            //radioButton.Click();

            //IWebElement frame3 = driver.FindElement(By.XPath("//*[@id=\"id3\"]/div/input"));
            //IWebElement frame3 = driver.FindElement(By.XPath("//*[@id=\"id3\"]/div/input"));
            //IWebElement frame3 = driver.FindElement(By.XPath("//*[@id=\"id3\"]/div/input"));




            //IWebElement frame2 = driver.FindElement(By.Name("mytext2"));
            //frame2.SendKeys("frame2 input");
            //Thread.Sleep(2000);

            //IWebElement frame3 = driver.FindElement(By.Name("mytext3"));
            //frame3.SendKeys("frame3 input");
            //Thread.Sleep(2000);

        }


    }
    //// switch frames
    //switchframes(framepath){
    //          driver.SwitchTo().Frame(driver.FindElement(By.XPath(framepath)));
    //}
    //public class CustomMethods
    //{
    //    public static void Click(IWebDriver driver, By locator)
    //    {
    //        //driver.FindElement(locator).Click();
    //        driver.SwitchTo().Frame(driver.FindElement(locator));
    //    }
    //}
    



}
