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
            // LeftClick on DataItem "Row0" at (13,17)
            Console.WriteLine("LeftClick on DataItem \"Row0\" at (13,17)");
            string xpath_LeftClickDataItemRow0_13_17 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Custom[@ClassName=\"DevAVDbView\"]/Group[@Name=\"dockLayoutManager\"][@AutomationId=\"dockLayoutManager\"]/Group[@Name=\"LayoutGroup\"][starts-with(@AutomationId,\"LayoutGroup\")]/Group[@Name=\"LayoutGroup\"][starts-with(@AutomationId,\"LayoutGroup\")]/Group[@Name=\"LayoutPanel\"][starts-with(@AutomationId,\"LayoutPanel\")]/Custom[@ClassName=\"EmployeeCollectionView\"]/Group[@Name=\"DockLayoutManager\"][@AutomationId=\"DockLayoutManager\"]/Group[@Name=\"LayoutGroup\"][starts-with(@AutomationId,\"LayoutGroup\")]/Group[@Name=\"LayoutPanel\"][starts-with(@AutomationId,\"LayoutPanel\")]/DataGrid[@AutomationId=\"tableViewGridControl\"]/Pane[@Name=\"DataPanel\"][@AutomationId=\"dataPresenter\"]/DataItem[@Name=\"Row0\"]";
            var winElem_LeftClickDataItemRow0_13_17 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickDataItemRow0_13_17);
            if (winElem_LeftClickDataItemRow0_13_17 != null)
            {
                winElem_LeftClickDataItemRow0_13_17.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickDataItemRow0_13_17}");
            }


            // LeftClick on Text "Tasks" at (12,10)
            Console.WriteLine("LeftClick on Text \"Tasks\" at (12,10)");
            string xpath_LeftClickTextTasks_12_10 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Custom[@ClassName=\"DevAVDbView\"]/Group[@Name=\"dockLayoutManager\"][@AutomationId=\"dockLayoutManager\"]/Group[@Name=\"LayoutGroup\"][starts-with(@AutomationId,\"LayoutGroup\")]/Group[@Name=\"LayoutGroup\"][starts-with(@AutomationId,\"LayoutGroup\")]/Group[@Name=\"LayoutPanel\"][starts-with(@AutomationId,\"LayoutPanel\")]/Custom[@ClassName=\"EmployeeCollectionView\"]/Group[@Name=\"DockLayoutManager\"][@AutomationId=\"DockLayoutManager\"]/Group[@Name=\"LayoutGroup\"][starts-with(@AutomationId,\"LayoutGroup\")]/Group[@Name=\"LayoutPanel\"][starts-with(@AutomationId,\"LayoutPanel\")]/Custom[@ClassName=\"EmployeePanelView\"]/Tab[@ClassName=\"TabControl\"]/TabItem[@ClassName=\"TabItem\"][@Name=\"Tasks\"]/Text[@ClassName=\"TextBlock\"][@Name=\"Tasks\"]";
            var winElem_LeftClickTextTasks_12_10 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickTextTasks_12_10);
            if (winElem_LeftClickTextTasks_12_10 != null)
            {
                winElem_LeftClickTextTasks_12_10.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickTextTasks_12_10}");
            }


            // LeftClick on Edit "Call  Jane" at (19,18)
            Console.WriteLine("LeftClick on Edit \"Call  Jane\" at (19,18)");
            string xpath_LeftClickEditCallJane_19_18 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Custom[@ClassName=\"DevAVDbView\"]/Group[@Name=\"dockLayoutManager\"][@AutomationId=\"dockLayoutManager\"]/Group[@Name=\"LayoutGroup\"][starts-with(@AutomationId,\"LayoutGroup\")]/Group[@Name=\"LayoutGroup\"][starts-with(@AutomationId,\"LayoutGroup\")]/Group[@Name=\"LayoutPanel\"][starts-with(@AutomationId,\"LayoutPanel\")]/Custom[@ClassName=\"EmployeeCollectionView\"]/Group[@Name=\"DockLayoutManager\"][@AutomationId=\"DockLayoutManager\"]/Group[@Name=\"LayoutGroup\"][starts-with(@AutomationId,\"LayoutGroup\")]/Group[@Name=\"LayoutPanel\"][starts-with(@AutomationId,\"LayoutPanel\")]/Custom[@ClassName=\"EmployeePanelView\"]/Tab[@ClassName=\"TabControl\"]/TabItem[@ClassName=\"TabItem\"][@Name=\"Tasks\"]/DataGrid[position()=2]/Pane[@Name=\"DataPanel\"][@AutomationId=\"dataPresenter\"]/DataItem[@Name=\"Row0\"]/Custom[@Name=\"Column1\"][@AutomationId=\"Subject\"]/Edit[@ClassName=\"TextEdit\"][@Name=\"Call  Jane\"]";
            var winElem_LeftClickEditCallJane_19_18 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickEditCallJane_19_18);
            if (winElem_LeftClickEditCallJane_19_18 != null)
            {
                winElem_LeftClickEditCallJane_19_18.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickEditCallJane_19_18}");
            }
        }

        [Test][Explicit]
        public void Test1()
        {
            Assert.Pass();
        }
    }
}