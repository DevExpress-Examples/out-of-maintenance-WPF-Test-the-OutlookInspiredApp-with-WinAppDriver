using Microsoft.Win32;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using outlookinspiredapp.uitest;
using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace OutlookInspiredApp.UITest
{
    public class Tests
    {        
        Process pWad;
        const string PathToTheDemo = @"C:\Work\2020.2\Demos.RealLife\DevExpress.OutlookInspiredApp\DevExpress.OutlookInspiredApp.Wpf\bin\Debug\DevExpress.OutlookInspiredApp.Wpf.exe";
        protected DesktopSession desktopSession { get; private set; }
        [OneTimeSetUp]
        public void FixtureSetup()
        {
            StartWAD();
            var options = new AppiumOptions();            
            options.AddAdditionalCapability(capabilityName: "app", capabilityValue: PathToTheDemo);
            options.AddAdditionalCapability(capabilityName: "deviceName", capabilityValue: "WindowsPC");
            options.AddAdditionalCapability(capabilityName: "platformName", capabilityValue: "Windows");
            options.AddAdditionalCapability(capabilityName: "ms:experimental-webdriver", capabilityValue: true);



            var driver = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), options);
            WaitSplashScreen(driver);
            desktopSession = new DesktopSession(driver);
        }

        static void WaitSplashScreen(WindowsDriver<WindowsElement> driver)
        {
            var cwh = driver.CurrentWindowHandle;
            while (driver.WindowHandles.Contains(cwh))
                Thread.Sleep(1000);
            driver.SwitchTo().Window(driver.WindowHandles[0]);
        }

        private void StartWAD()
        {
            var psi = new ProcessStartInfo(@"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe");            
            psi.EnvironmentVariables.Add("DX.UITESTINGENABLED", "1");
            pWad = Process.Start(psi);
        }

        [OneTimeTearDown]
        public void FixtureTearDown()
        {
            desktopSession.CloseApp();
            pWad.Kill();
        }

        [SetUp]
        public void Setup()
        {
        }

        [Test]
        [Order(0)]
        public void CreateEmployee()
        {
            var desktopElement = desktopSession.DesktopSessionElement;            
            var bNewEmployee = desktopElement.FindElementByName("New Employee");
            bNewEmployee.Click();

            WindowsElement newEmployeeWindow = null;
            while (newEmployeeWindow == null)
                newEmployeeWindow = desktopElement.FindElementByName("Employee (New)");

            newEmployeeWindow.FindElementByName("First Name").FindElementByClassName("TextEdit").SendKeys("John");
            newEmployeeWindow.FindElementByName("Last Name").FindElementByClassName("TextEdit").SendKeys("Public");
            newEmployeeWindow.FindElementByName("Title").FindElementByClassName("TextEdit").SendKeys("CTO");
            newEmployeeWindow.FindElementByName("Mobile Phone").FindElementByClassName("ButtonEdit").SendKeys("1111111111");
            newEmployeeWindow.FindElementByName("Email").FindElementByClassName("ButtonEdit").SendKeys("john.public@dx-email.com");


            newEmployeeWindow.FindElementByName("Save & Close").Click();
        }
        [Test]
        [Order(1)]
        public void CreateTask()
        {
            var desktopElement = desktopSession.DesktopSessionElement;
            var bNewEmployee = desktopElement.FindElementByName("Task");
            bNewEmployee.Click();

            WindowsElement newTaskWindow = null;
            while (newTaskWindow == null)
                newTaskWindow = desktopElement.FindElementByName("EmployeeTask (New)");

            newTaskWindow.FindElementByName("Subject").FindElementByClassName("TextEdit").SendKeys("Call Jane");
            var comboBox = newTaskWindow.FindElementByName("Assigned To").FindElementByClassName("ComboBoxEdit");
            comboBox.FindElementByXPath("//Edit[@ClassName=\"ButtonEdit\"]/Button[@ClassName=\"Button\"]").Click();
            comboBox.SendKeys("John Public");            

            newTaskWindow.FindElementByName("Save & Close").Click();
        }

        [Test]
        [Order(2)]
        public void CheckTasks()
        {
            var desktopElement = desktopSession.DesktopSessionElement;
            var dockManager = desktopElement.FindElementByName("DockLayoutManager");
            var grid = dockManager.FindElementByAccessibilityId("tableViewGridControl");

            grid.FindElementByAccessibilityId("SearchComboBox").SendKeys("John Public");
            grid.FindElementByName("John Public").Click();

            var tab = dockManager.FindElementsByClassName("TabItem").First(x => x.GetAttribute("Name") == "Tasks");
            tab.Click();

            var row = tab.FindElementByAccessibilityId("Row0");
            var callJane = row.FindElementsByClassName("TextEdit").FirstOrDefault(x=>x.GetAttribute("Name") == "Call Jane");
            Assert.IsNotNull(callJane);         
        }

        [Test][Explicit]
        public void Test1()
        {
            Assert.Pass();
        }
    }
}