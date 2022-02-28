using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1:AppSession
    {
        private static WindowsElement CalculationResult;
        [TestMethod]
        public void Addition()
        {
            session.FindElementByAccessibilityId("num1").SendKeys("9");
            session.FindElementByAccessibilityId("ops").SendKeys("{+}");
            //运算符"{加减乘除}"
            session.FindElementByAccessibilityId("num2").SendKeys("9");
            session.FindElementByAccessibilityId("button1").Click();
            Assert.AreEqual("18", GetCalculatorResultText());
        }

        [TestMethod]
        public void Subtraction()
        {
            session.FindElementByAccessibilityId("num1").SendKeys("9");
            session.FindElementByAccessibilityId("ops").SendKeys("{-}");
            //运算符"{加减乘除}"
            session.FindElementByAccessibilityId("num2").SendKeys("9");
            session.FindElementByAccessibilityId("button1").Click();
            Assert.AreEqual("0", GetCalculatorResultText());
        }

        [TestMethod]
        public void Multiplication()
        {
            session.FindElementByAccessibilityId("num1").SendKeys("9");
            session.FindElementByAccessibilityId("ops").SendKeys("{*}");
            //运算符"{加减乘除}"
            session.FindElementByAccessibilityId("num2").SendKeys("9");
            session.FindElementByAccessibilityId("button1").Click();
            Assert.AreEqual("81", GetCalculatorResultText());
        }

        [TestMethod]
        public void Division()
        {
            session.FindElementByAccessibilityId("num1").SendKeys("9");
            session.FindElementByAccessibilityId("ops").SendKeys("{/}");
            //运算符"{加减乘除}"
            session.FindElementByAccessibilityId("num2").SendKeys("9");
            session.FindElementByAccessibilityId("button1").Click();
            Assert.AreEqual("1", GetCalculatorResultText());
        }

        [TestMethod]
        [DataRow("1", "{+}", "7", "8")]
        [DataRow("9", "{-}", "1", "8")]
        [DataRow("8", "{/}", "8", "1")]
        public void Templatized(string input1, string operation, string input2, string expectedResult)
        {
            // Run sequence of button presses specified above and validate the results
            session.FindElementByAccessibilityId("num1").SendKeys(input1);
            session.FindElementByAccessibilityId("ops").SendKeys(operation);
            session.FindElementByAccessibilityId("num2").SendKeys(input2);
            session.FindElementByAccessibilityId("button1").Click();
            Assert.AreEqual(expectedResult, GetCalculatorResultText());
        }



        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            // Create session to launch a Calculator window
            SetUp();

            // Locate the calculatorResult element
            //calculatorResult= session.FindElementByName("labResult");
            CalculationResult = session.FindElementByAccessibilityId("answer");

            Assert.AreEqual("", CalculationResult.Text);
        }
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TearDown();
        }

        [TestInitialize]
        public void Clear()
        {
            session.FindElementByAccessibilityId("num1").Clear();
            session.FindElementByAccessibilityId("num2").Clear();
            session.FindElementByAccessibilityId("ops").Clear();
            //Assert.IsNull("0", GetCalculatorResultText());
        }

        private string GetCalculatorResultText()
        {
            return CalculationResult.Text.Trim();
        }
    }
}
