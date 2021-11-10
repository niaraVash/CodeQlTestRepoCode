

namespace MSTeamStoreAppsTool
{
   
    using OpenQA.Selenium;
    using OpenQA.Selenium.Support.UI;
    using System;
    using System.Diagnostics;
    using System.Threading;

    /// <summary>
    /// Enum for condition.
    /// </summary>
    public enum ConditionType
    {
        /// <summary>
        /// Element exist check.
        /// </summary>
        exists,

        /// <summary>
        /// Element is clickable.
        /// </summary>
        clickable,

        /// <summary>
        /// Chek for both exist and clickable.
        /// </summary>
        both,

        /// <summary>
        /// Chek for visible.
        /// </summary>
        visible,

        /// <summary>
        /// Checkk for Switching.
        /// </summary>
        switchFrame,

        /// <summary>
        /// Checkk for Switching.
        /// </summary>
        elementWait,
    }

    public static class HelperConstants
    {
        public const string Button = "button";
        public const string ElementNotFound = "Element not found";
        public const string ArgumentsScrollIntoView = "arguments[0].scrollIntoView(true);";
    }

    /// <summary>
    /// Its a generic Waiter utility tool class.
    /// Provides Wait methods for an elements, and AJAX elements to load. It uses WebDriverWait (explicit wait) for waiting an element or javaScript.  
    /// It also uses DefaultWait for Waitin for WebElements
    /// </summary>
    public static class Waiter
    {
        #region

        /// <summary>
        /// 5000 milliseconds to wait for an action
        /// </summary>
        public const int Waitfor5Seconds = 5000;

        /// <summary>
        /// 5000 milliseconds to wait for an action
        /// </summary>
        public const int Waitfor3Seconds = 3000;

        /// <summary>
        /// 5 seconds to wait for an action
        /// </summary>
        public const int Waitfor5SecondsLoad = 5;

        /// <summary>
        /// 6 seconds to wait for an action
        /// </summary>
        public const int Waitfor6SecondsLoad = 6;

        /// <summary>
        /// 8 seconds to wait for an action
        /// </summary>
        public const int Waitfor8SecondsLoad = 8;

        /// <summary>
        /// 8 seconds to wait for an action
        /// </summary>
        public const int Waitfor15SecondsLoad = 15;

        /// <summary>
        /// 10 seconds to wait for an action
        /// </summary>
        public const int Waitfor10SecondsLoad = 10;
        
        /// <summary>
        /// 25 seconds to wait for an action
        /// </summary>
        public const int Waitfor40SecondsLoad = 40;

        /// <summary>
        /// Default time in milliseconds to wait for an action
        /// </summary>
        private const int DefaultMilliSeconds = 2000;

        /// <summary>
        /// Timeout Value for a control, set it to default as max timeout is 300, but it should be configurable in the calling method as per the requirement.
        /// </summary>
        private const int DefaultControlWaiterTimeout = 300;


        #endregion





        /// <summary>
        /// To verify is control is persent or not.
        /// </summary>
        /// <param name="driver">web driver</param>
        /// <param name="bylocator">by locator</param>
        /// <returns>true or false</returns>
        public static bool IsPresent(this IWebDriver driver, By bylocator)
        {
            var isPresent = true;
            try
            {
                driver.FindElement(bylocator);
            }
            catch (NoSuchElementException)
            {
                isPresent = false;
            }

            return isPresent;
        }

        /// <summary>
        /// Wait Milliseconds
        /// </summary>
        /// <param name="milliseconds">Wait milliseconds</param>
        public static void WaitMilliseconds(int milliseconds = DefaultMilliSeconds)
        {
            var stopWatch = Stopwatch.StartNew();
            while (stopWatch.ElapsedMilliseconds < milliseconds)
            {
                Thread.Sleep(1);
            }
        }

        

        /// <summary>
        /// Method to wait for element to load  in different scenerios
        /// </summary>
        /// <param name="driver">IWeb Driver</param>
        /// <param name="by">Element</param> 
        /// <param name="condition">Scenerios of element wait</param>
        /// <param name="timeoutSeconds">Time to wait</param>
        public static void WaitForElementToLoad(this IWebDriver driver, By by, ConditionType condition, int timeoutSeconds = DefaultControlWaiterTimeout)
        {
            try
            {
                switch (condition)
                {
                    case ConditionType.exists:
                        WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, timeoutSeconds));
                        var element = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(by));
                        break;
                    case ConditionType.clickable:
                        WebDriverWait clickablewait = new WebDriverWait(driver, new TimeSpan(0, 0, timeoutSeconds));
                        var clickableelement = clickablewait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(by));
                        break;
                    case ConditionType.visible:
                        WebDriverWait visiblewait = new WebDriverWait(driver, new TimeSpan(0, 0, timeoutSeconds));
                        var visibleelement = visiblewait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(by));
                        break;
                    case ConditionType.both:
                        WebDriverWait bothwait = new WebDriverWait(driver, new TimeSpan(0, 0, timeoutSeconds));
                        var bothelement = bothwait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(by));
                        if (bothelement.TagName != HelperConstants.Button)
                        {
                            bothelement = bothwait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(by));
                        }
                        else if (bothelement.TagName == HelperConstants.Button)
                        {
                            bothelement = bothwait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(by));
                        }
                        else
                        {
                           
                            throw new Exception(HelperConstants.ElementNotFound);
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
               
                throw ex;
            }
        }

      
    }
}
