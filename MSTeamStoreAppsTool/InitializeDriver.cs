using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Text;
//newcomment
namespace MSTeamStoreAppsTool
{
    public class InitializeDriver
    {
      
            IWebDriver webDriver;
            public void Init_Browser()
            {
               // webDriver = new ChromeDriver();
            var options = new ChromeOptions();
            options.AddArgument("no-sandbox");
            options.AddArgument("--incognito");
          webDriver= new ChromeDriver(options);
            webDriver.Manage().Window.Maximize();
            }
            public string Title
            {
                get { return webDriver.Title; }
            }
            public void Goto(string url)
            {
                webDriver.Url = url;
            }
            public void Close(IWebDriver driver)
            {
            if (driver == null)
            {
                return;
            }

            driver.Close();
            driver.Quit();
            driver.Dispose();
        }
            public IWebDriver getDriver
            {
                get { return webDriver; }
            }
        
    }
}
