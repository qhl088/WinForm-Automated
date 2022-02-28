using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTestProject1
{
    public class AppSession
    {
        private const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        private const string CalculatorAppId = @"D:\InterviewApp\InterviewApp.exe";
        protected static WindowsDriver<WindowsElement> session;
        
        private WindowsElement _InvildOpe;

        

        /// <summary>
        /// 启动
        /// </summary>
        public static void SetUp()
        {
            if(session == null)
            {
                DesiredCapabilities appCapabilities = new DesiredCapabilities();
                appCapabilities.SetCapability("app",CalculatorAppId);
                appCapabilities.SetCapability("deviceName","WindosPC");
                //"deviceName", "WindowsPC"
                session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl),appCapabilities);
                Assert.IsNotNull(session);
                session.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(1.5));
            }
        }
        
        /// <summary>
        /// 退出
        /// </summary>
        public static void TearDown()
        {
            if(session!=null)
            {
                session.Quit();
                session = null;
            }
        }

        public bool IsWinBoxExist
        {
            get
            {
                _InvildOpe = session.FindElementById("2A.3C074A");
                return _InvildOpe.Displayed;
            }
        }

        //public void IsWarningExist()
        //{
        //    try
        //    {
        //        if(IsWinBoxExist==true)
        //        {
        //            switch()
        //            {
        //                case a:
        //                    break;
        //            }                                          
        //        }
        //    }
        //    catch
        //    {

        //    }
        
    }
}
