using NUnit.Framework;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace MSTeamStoreAppsTool
{
    public class ReadApps
    {
        InitializeDriver brow = new InitializeDriver();
        String test_url = "https://teams.microsoft.com/";
        IWebDriver driver;

        [SetUp]
        public void StartBrowser()
        {
            brow.Init_Browser();
        }

        [Test]
        public void Test2()
        {
            brow.Goto(test_url);
            driver = brow.getDriver;
            VerifyTeamsLogin(driver);
            OpenSTore(driver);

            // WriteSample();


        }

        public void WriteSample(List<string> lstAppname, List<string> capability)
        {
            Excel.Application excelApp = new Excel.Application();
            if (excelApp != null)
            {
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets.Add();

                for (int i = 1; i <= lstAppname.Count; i++)
                {
                    excelWorksheet.Cells[i, 1] = lstAppname[i - 1];

                    for (int j = 0; j < capability.Count; j++)
                    {
                        excelWorksheet.Cells[2, 1] = "Value2";
                        excelWorksheet.Cells[3, 1] = "Value3";
                        excelWorksheet.Cells[4, 1] = "Value4";
                    }
                }

                excelApp.ActiveWorkbook.SaveAs(@"C:\abc.xls", Excel.XlFileFormat.xlWorkbookNormal);

                excelWorkbook.Close();
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        [TearDown]
        public void close_Browser()
        {
            brow.Close(driver);
        }

        /// <summary>
        /// Login in to MS Team
        /// </summary>
        /// <param name="driver"></param>
        /// <returns>true or false</returns>
        public bool VerifyTeamsLogin(IWebDriver driver)
        {
            bool isResult = false;
            try
            {
                driver.WaitForElementToLoad(By.Id("i0116"), ConditionType.both, Waiter.Waitfor40SecondsLoad);
                driver.FindElement(By.Id("i0116")).SendKeys("email");
                driver.FindElement(By.Id("idSIButton9")).Click();
                try
                {
                    driver.WaitForElementToLoad(By.Id("i0118"), ConditionType.both, Waiter.Waitfor40SecondsLoad);
                    driver.FindElement(By.Id("i0118")).SendKeys("Password");
                }
                catch
                {
                    Waiter.WaitMilliseconds(Waiter.Waitfor5Seconds);
                    driver.WaitForElementToLoad(By.Id("i0118"), ConditionType.both, Waiter.Waitfor40SecondsLoad);
                    driver.FindElement(By.Id("i0118")).SendKeys("Password");
                }

                driver.WaitForElementToLoad(By.Id("idSIButton9"), ConditionType.both, Waiter.Waitfor40SecondsLoad);
                driver.FindElement(By.Id("idSIButton9")).Click();
                //// YesButtonLogin;
                try
                {
                    driver.WaitForElementToLoad(By.Id("KmsiDescription"), ConditionType.both, Waiter.Waitfor40SecondsLoad);
                    driver.FindElement(By.Id("idSIButton9")).Click();
                }
                catch
                {
                    ////ignore
                }
                try
                {

                    driver.WaitForElementToLoad(By.XPath("//*[@id='download-desktop-page']/div/a"), ConditionType.both, Waiter.Waitfor15SecondsLoad);
                    if (driver.IsPresent(By.XPath("//*[@id='download-desktop-page']/div/a")))
                    {
                        driver.FindElement(By.XPath("//*[@id='download-desktop-page']/div/a")).Click();
                    }

                    while (driver.IsPresent(By.XPath("//span[@class='loadingtext']")))
                    {
                        ////Check for loading page to end.
                    }

                }

                catch (Exception ex)

                {
                    ////bypass catch
                }

                driver.WaitForElementToLoad(By.XPath("//button[@data-tid='app-bar-2a84919f-59d8-4441-a975-2a8c2643b741']"), ConditionType.both, Waiter.Waitfor40SecondsLoad);
                isResult = true;
            }
            catch (Exception ex)
            {

            }

            return isResult;
        }


        public void OpenSTore(IWebDriver driver)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets.Add();
            List<string> capability = null;
            string appname = ""; int rowCount = 1;
            List<string> lstAppname = new List<string>();
            int sno = 1;
            int appCount = 0;
            IList<IWebElement> appDiv = null;

            driver.WaitForElementToLoad(By.XPath("//button[@id='discover-apps-button']"), ConditionType.both, Waiter.Waitfor40SecondsLoad);
            driver.FindElement(By.XPath("//button[@id='discover-apps-button']")).Click();
            driver.WaitForElementToLoad(By.XPath("//div[@class='td-apps-gallery-app']"), ConditionType.both, Waiter.Waitfor40SecondsLoad);
            appDiv = driver.FindElements(By.XPath("//div[@class='td-apps-gallery-app']"));
            appCount = appDiv.Count;
            try
            {
                if (driver.IsPresent(By.XPath("//button[@title='Dismiss']")))
                {
                    driver.WaitForElementToLoad(By.XPath("//button[@title='Dismiss']"), ConditionType.both, Waiter.Waitfor40SecondsLoad);
                    driver.FindElement(By.XPath("//button[@title='Dismiss']")).Click();
                }

                for (int i = 1; i < appCount; i++)
                {

                    capability = new List<string>();
                    if (i == 1)
                    {
                        var appDivPath = driver.FindElements(By.XPath("//div[@class='td-apps-gallery-app'][1]//span[@class='app-name']"));
                        for (int k = 0; k < appDivPath.Count; k++)
                        {
                            capability = new List<string>();
                            appname = appDivPath[k].Text;
                            ///write in excel file
                            if (excelApp != null)
                            {
                                excelWorksheet.Cells[rowCount, 1] = sno;
                                excelWorksheet.Cells[rowCount, 2] = appname;
                                rowCount++;
                                sno++;
                            }                            
                        }
                    }
                    else
                    {
                        try
                        {

                            appname = driver.FindElement(By.XPath("//div[@class='td-apps-gallery-app'][" + i + "]//span[@class='app-name']")).Text;
                        }
                        catch { }

                        if (!lstAppname.Contains(appname) && appname != "")
                        {
                            lstAppname.Add(appname);
                            driver.WaitForElementToLoad(By.XPath("//div[@class='td-apps-gallery-app'][" + i + "]//span[@class='app-name']"), ConditionType.both, Waiter.Waitfor40SecondsLoad);

                            IWebElement control = driver.FindElement(By.XPath("//div[@class='td-apps-gallery-app'][" + i + "]//span[@class='app-name']"));
                            // Create a javascript executor
                            IJavaScriptExecutor js = driver as IJavaScriptExecutor;
                            // Run the javascript command 'scrollintoview on the element
                            js.ExecuteScript(HelperConstants.ArgumentsScrollIntoView, control);
                            //waiter to complete task
                           // Waiter.WaitMilliseconds(Waiter.Waitfor3Seconds);

                            ///write in excel file
                            if (excelApp != null)
                            {
                                excelWorksheet.Cells[rowCount, 1] = sno;
                                excelWorksheet.Cells[rowCount, 2] = appname;
                                rowCount++;
                                sno++;

                            }                           
                        }

                        appDiv = driver.FindElements(By.XPath("//div[@class='td-apps-gallery-app']"));
                        appCount = appDiv.Count;
                    }
                }

                excelApp.ActiveWorkbook.SaveAs(@"D:\abcdef.xls", Excel.XlFileFormat.xlWorkbookNormal);
                excelWorkbook.Close();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {

            }
        }
    }
}
