using Microsoft.Win32;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using outlookinspiredapp.uitest;
using System;
using System.Diagnostics;
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
            // LeftClick on Button "New Employee" at (74,97)
            Console.WriteLine("LeftClick on Button \"New Employee\" at (74,97)");
            string xpath_LeftClickButtonNewEmploye_74_97 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Custom[@ClassName=\"DevAVDbView\"]/Group[@Name=\"RibbonControl\"][@AutomationId=\"ribbonControl\"]/Pane[@ClassName=\"RibbonSelectedPageControl\"][@Name=\"Lower Ribbon\"]/Group[@ClassName=\"RibbonPageGroupControl\"][@Name=\"PageGroup0\"]/Button[@Name=\"New Employee\"][@AutomationId=\"xBDBC878BD43BA1825AE8BD81B2A7076A\"]";
            var winElem_LeftClickButtonNewEmploye_74_97 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickButtonNewEmploye_74_97);
            if (winElem_LeftClickButtonNewEmploye_74_97 != null)
            {
                winElem_LeftClickButtonNewEmploye_74_97.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickButtonNewEmploye_74_97}");
            }


            // LeftClick on Edit "" at (170,18)
            Console.WriteLine("LeftClick on Edit \"\" at (170,18)");
            string xpath_LeftClickEdit_170_18 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Window[@ClassName=\"Window\"][@Name=\"Employee (New)\"]/Custom[@ClassName=\"EmployeeView\"]/Group[@Name=\"LayoutControl\"][@AutomationId=\"layoutControl\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"entityGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"imageGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"LayoutGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"LayoutGroup\"]/Pane[@Name=\"First Name\"][@AutomationId=\"First Name\"]/Edit[@ClassName=\"TextEdit\"]/Edit[@AutomationId=\"PART_Editor\"]";
            var winElem_LeftClickEdit_170_18 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickEdit_170_18);
            if (winElem_LeftClickEdit_170_18 != null)
            {
                winElem_LeftClickEdit_170_18.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickEdit_170_18}");
            }


            // KeyboardInput VirtualKeys="Keys.LeftShift + "j" + Keys.LeftShift"ohn"" CapsLock=False NumLock=True ScrollLock=False
            Console.WriteLine("KeyboardInput VirtualKeys=\"Keys.LeftShift + \"j\" + Keys.LeftShift\"ohn\"\" CapsLock=False NumLock=True ScrollLock=False");
            System.Threading.Thread.Sleep(100);
            winElem_LeftClickEdit_170_18.SendKeys(Keys.LeftShift + "j" + Keys.LeftShift);
            winElem_LeftClickEdit_170_18.SendKeys("ohn");


            // LeftClick on Edit "" at (256,2)
            Console.WriteLine("LeftClick on Edit \"\" at (256,2)");
            string xpath_LeftClickEdit_256_2 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Window[@ClassName=\"Window\"][@Name=\"Employee (New)\"]/Custom[@ClassName=\"EmployeeView\"]/Group[@Name=\"LayoutControl\"][@AutomationId=\"layoutControl\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"entityGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"imageGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"LayoutGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"LayoutGroup\"]/Pane[@Name=\"Last Name\"][@AutomationId=\"Last Name\"]/Edit[@ClassName=\"TextEdit\"]/Edit[@AutomationId=\"PART_Editor\"]";
            var winElem_LeftClickEdit_256_2 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickEdit_256_2);
            if (winElem_LeftClickEdit_256_2 != null)
            {
                winElem_LeftClickEdit_256_2.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickEdit_256_2}");
            }


            // KeyboardInput VirtualKeys="Keys.LeftShift + "p" + Keys.LeftShift"ublic"" CapsLock=False NumLock=True ScrollLock=False
            Console.WriteLine("KeyboardInput VirtualKeys=\"Keys.LeftShift + \"p\" + Keys.LeftShift\"ublic\"\" CapsLock=False NumLock=True ScrollLock=False");
            System.Threading.Thread.Sleep(100);
            winElem_LeftClickEdit_256_2.SendKeys(Keys.LeftShift + "P" + Keys.LeftShift);
            winElem_LeftClickEdit_256_2.SendKeys("ublic");


            // LeftClick on Edit "" at (324,16)
            Console.WriteLine("LeftClick on Edit \"\" at (324,16)");
            string xpath_LeftClickEdit_324_16 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Window[@ClassName=\"Window\"][@Name=\"Employee (New)\"]/Custom[@ClassName=\"EmployeeView\"]/Group[@Name=\"LayoutControl\"][@AutomationId=\"layoutControl\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"entityGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"imageGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"LayoutGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"LayoutGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"LayoutGroup\"]/Pane[@Name=\"Title\"][@AutomationId=\"Title\"]/Edit[@ClassName=\"TextEdit\"]/Edit[@AutomationId=\"PART_Editor\"]";
            var winElem_LeftClickEdit_324_16 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickEdit_324_16);
            if (winElem_LeftClickEdit_324_16 != null)
            {
                winElem_LeftClickEdit_324_16.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickEdit_324_16}");
            }


            // KeyboardInput VirtualKeys="Keys.LeftShift + "c" + "t" + "o" + Keys.LeftShift" CapsLock=False NumLock=True ScrollLock=False
            Console.WriteLine("KeyboardInput VirtualKeys=\"Keys.LeftShift + \"c\" + \"t\" + \"o\" + Keys.LeftShift\" CapsLock=False NumLock=True ScrollLock=False");
            System.Threading.Thread.Sleep(100);
            winElem_LeftClickEdit_324_16.SendKeys(Keys.LeftShift + "c" + "t" + "o" + Keys.LeftShift);


            // LeftClick on Edit "" at (392,15)
            Console.WriteLine("LeftClick on Edit \"\" at (392,15)");
            string xpath_LeftClickEdit_392_15 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Window[@ClassName=\"Window\"][@Name=\"Employee (New)\"]/Custom[@ClassName=\"EmployeeView\"]/Group[@Name=\"LayoutControl\"][@AutomationId=\"layoutControl\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"entityGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"imageGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"LayoutGroup\"]/Pane[@Name=\"Mobile Phone\"][@AutomationId=\"Mobile Phone\"]/Edit[@ClassName=\"ButtonEdit\"]/Edit[@AutomationId=\"PART_Editor\"]";
            var winElem_LeftClickEdit_392_15 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickEdit_392_15);
            if (winElem_LeftClickEdit_392_15 != null)
            {
                winElem_LeftClickEdit_392_15.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickEdit_392_15}");
            }


            // KeyboardInput VirtualKeys="Keys.Home + Keys.Home"1111111111"" CapsLock=False NumLock=True ScrollLock=False
            Console.WriteLine("KeyboardInput VirtualKeys=\"Keys.Home + Keys.Home\"1111111111\"\" CapsLock=False NumLock=True ScrollLock=False");
            System.Threading.Thread.Sleep(100);
            winElem_LeftClickEdit_392_15.SendKeys(Keys.Home + Keys.Home);
            winElem_LeftClickEdit_392_15.SendKeys("1111111111");


            // LeftClick on Edit "" at (402,12)
            Console.WriteLine("LeftClick on Edit \"\" at (402,12)");
            string xpath_LeftClickEdit_402_12 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Window[@ClassName=\"Window\"][@Name=\"Employee (New)\"]/Custom[@ClassName=\"EmployeeView\"]/Group[@Name=\"LayoutControl\"][@AutomationId=\"layoutControl\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"entityGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"imageGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"LayoutGroup\"]/Pane[@Name=\"Email\"][@AutomationId=\"Email\"]/Edit[@ClassName=\"ButtonEdit\"]/Edit[@AutomationId=\"PART_Editor\"]";
            var winElem_LeftClickEdit_402_12 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickEdit_402_12);
            if (winElem_LeftClickEdit_402_12 != null)
            {
                winElem_LeftClickEdit_402_12.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickEdit_402_12}");
            }


            // KeyboardInput VirtualKeys=""john.public"Keys.LeftShift + "2" + Keys.LeftShift"dx-email.com"" CapsLock=False NumLock=True ScrollLock=False
            Console.WriteLine("KeyboardInput VirtualKeys=\"\"john.public\"Keys.LeftShift + \"2\" + Keys.LeftShift\"dx-email.com\"\" CapsLock=False NumLock=True ScrollLock=False");
            System.Threading.Thread.Sleep(100);
            winElem_LeftClickEdit_402_12.SendKeys("john.public");
            winElem_LeftClickEdit_402_12.SendKeys(Keys.LeftShift + "2" + Keys.LeftShift);
            winElem_LeftClickEdit_402_12.SendKeys("dx-email.com");          

            // LeftClick on Button "Save" at (20,40)
            Console.WriteLine("LeftClick on Button \"Save\" at (20,40)");
            string xpath_LeftClickButtonSave_20_40 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Window[@ClassName=\"Window\"][@Name=\"Employee (New)\"]/Custom[@ClassName=\"EmployeeView\"]/Group[@Name=\"RibbonControl\"][@AutomationId=\"ribbonControl\"]/Pane[@ClassName=\"RibbonSelectedPageControl\"][@Name=\"Lower Ribbon\"]/Group[@ClassName=\"RibbonPageGroupControl\"][@Name=\"Save\"]/Button[@Name=\"Save\"][@AutomationId=\"xF0A2A8AA912F28E23E49FCAD8DF0C0B9\"]";
            var winElem_LeftClickButtonSave_20_40 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickButtonSave_20_40);
            if (winElem_LeftClickButtonSave_20_40 != null)
            {
                winElem_LeftClickButtonSave_20_40.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickButtonSave_20_40}");
            }


            // LeftClick on Button "Close" at (53,42)
            Console.WriteLine("LeftClick on Button \"Close\" at (53,42)");
            string xpath_LeftClickButtonClose_53_42 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Window[@ClassName=\"Window\"][@Name=\"John Public\"]/Custom[@ClassName=\"EmployeeView\"]/Group[@Name=\"RibbonControl\"][@AutomationId=\"ribbonControl\"]/Pane[@ClassName=\"RibbonSelectedPageControl\"][@Name=\"Lower Ribbon\"]/Group[@ClassName=\"RibbonPageGroupControl\"][@Name=\"Close\"]/Button[@Name=\"Close\"][@AutomationId=\"x5073E1AD7C525B2DB4CD533F9C83448F\"]";
            var winElem_LeftClickButtonClose_53_42 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickButtonClose_53_42);
            if (winElem_LeftClickButtonClose_53_42 != null)
            {
                winElem_LeftClickButtonClose_53_42.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickButtonClose_53_42}");
            }
        }
        [Test]
        [Order(1)]
        public void CreateTask()
        {
            // LeftClick on Button "Task" at (12,46)
            Console.WriteLine("LeftClick on Button \"Task\" at (12,46)");
            string xpath_LeftClickButtonTask_12_46 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Custom[@ClassName=\"DevAVDbView\"]/Group[@Name=\"RibbonControl\"][@AutomationId=\"ribbonControl\"]/Pane[@ClassName=\"RibbonSelectedPageControl\"][@Name=\"Lower Ribbon\"]/Group[@ClassName=\"RibbonPageGroupControl\"][@Name=\"Actions\"]/Button[@Name=\"Task\"][@AutomationId=\"xD4D45A8D530F6B115D09D16A5D8E5D2A\"]";
            var winElem_LeftClickButtonTask_12_46 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickButtonTask_12_46);
            if (winElem_LeftClickButtonTask_12_46 != null)
            {
                winElem_LeftClickButtonTask_12_46.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickButtonTask_12_46}");
            }


            // LeftClick on Edit "" at (221,11)
            Console.WriteLine("LeftClick on Edit \"\" at (221,11)");
            string xpath_LeftClickEdit_221_11 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Window[@ClassName=\"Window\"][@Name=\"EmployeeTask (New)\"]/Custom[@ClassName=\"EmployeeTaskView\"]/Group[@Name=\"LayoutControl\"][@AutomationId=\"LayoutControl\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"LayoutGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"LayoutGroup\"]/Pane[@Name=\"Subject\"][@AutomationId=\"Subject\"]/Edit[@ClassName=\"TextEdit\"]/Edit[@AutomationId=\"PART_Editor\"]";
            var winElem_LeftClickEdit_221_11 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickEdit_221_11);
            if (winElem_LeftClickEdit_221_11 != null)
            {
                winElem_LeftClickEdit_221_11.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickEdit_221_11}");
            }


            // KeyboardInput VirtualKeys="Keys.LeftShift + "c" + Keys.LeftShift"all"Keys.Space + Keys.SpaceKeys.LeftShift + "j" + Keys.LeftShift"ane"" CapsLock=False NumLock=True ScrollLock=False
            Console.WriteLine("KeyboardInput VirtualKeys=\"Keys.LeftShift + \"c\" + Keys.LeftShift\"all\"Keys.Space + Keys.SpaceKeys.LeftShift + \"j\" + Keys.LeftShift\"ane\"\" CapsLock=False NumLock=True ScrollLock=False");
            System.Threading.Thread.Sleep(100);
            winElem_LeftClickEdit_221_11.SendKeys(Keys.LeftShift + "c" + Keys.LeftShift);
            winElem_LeftClickEdit_221_11.SendKeys("all");
            winElem_LeftClickEdit_221_11.SendKeys(Keys.Space + Keys.Space);
            winElem_LeftClickEdit_221_11.SendKeys(Keys.LeftShift + "j" + Keys.LeftShift);
            winElem_LeftClickEdit_221_11.SendKeys("ane");


            // LeftClick on Button "" at (23,20)
            Console.WriteLine("LeftClick on Button \"\" at (23,20)");
            string xpath_LeftClickButton_23_20 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Window[@ClassName=\"Window\"][@Name=\"EmployeeTask (New)\"]/Custom[@ClassName=\"EmployeeTaskView\"]/Group[@Name=\"LayoutControl\"][@AutomationId=\"LayoutControl\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"LayoutGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"LayoutGroup\"]/Pane[@Name=\"Assigned To\"][@AutomationId=\"Assigned To\"]/ComboBox[@ClassName=\"ComboBoxEdit\"][@Name=\"Amelia Harper\"]/Edit[@ClassName=\"ButtonEdit\"][@Name=\"Amelia Harper\"]/Button[@ClassName=\"Button\"]";
            var winElem_LeftClickButton_23_20 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickButton_23_20);
            if (winElem_LeftClickButton_23_20 != null)
            {
                winElem_LeftClickButton_23_20.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickButton_23_20}");
            }


            // LeftClick on Edit "" at (172,15)
            Console.WriteLine("LeftClick on Edit \"\" at (172,15)");
            string xpath_LeftClickEdit_172_15 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Window[@ClassName=\"Window\"][@Name=\"EmployeeTask (New)\"]/Custom[@ClassName=\"EmployeeTaskView\"]/Group[@Name=\"LayoutControl\"][@AutomationId=\"LayoutControl\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"LayoutGroup\"]/Group[@Name=\"LayoutGroup\"][@AutomationId=\"LayoutGroup\"]/Pane[@Name=\"Assigned To\"][@AutomationId=\"Assigned To\"]/ComboBox[@ClassName=\"ComboBoxEdit\"]/Edit[@ClassName=\"ButtonEdit\"]";
            var winElem_LeftClickEdit_172_15 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickEdit_172_15);
            if (winElem_LeftClickEdit_172_15 != null)
            {
                winElem_LeftClickEdit_172_15.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickEdit_172_15}");
            }


            // KeyboardInput VirtualKeys="Keys.LeftShift + "j" + Keys.LeftShift"ohn"Keys.Space + Keys.SpaceKeys.LeftShift + "p" + Keys.LeftShift"ublic"Keys.Return + Keys.Return" CapsLock=False NumLock=True ScrollLock=False
            Console.WriteLine("KeyboardInput VirtualKeys=\"Keys.LeftShift + \"j\" + Keys.LeftShift\"ohn\"Keys.Space + Keys.SpaceKeys.LeftShift + \"p\" + Keys.LeftShift\"ublic\"Keys.Return + Keys.Return\" CapsLock=False NumLock=True ScrollLock=False");
            System.Threading.Thread.Sleep(100);
            winElem_LeftClickEdit_172_15.SendKeys(Keys.LeftShift + "j" + Keys.LeftShift);
            winElem_LeftClickEdit_172_15.SendKeys("ohn");            
            winElem_LeftClickEdit_172_15.SendKeys(Keys.Down + Keys.Down);
            winElem_LeftClickEdit_172_15.SendKeys(Keys.Return + Keys.Return);


            // LeftClick on Button "Save" at (21,61)
            Console.WriteLine("LeftClick on Button \"Save\" at (21,61)");
            string xpath_LeftClickButtonSave_21_61 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Window[@ClassName=\"Window\"][@Name=\"EmployeeTask (New)\"]/Custom[@ClassName=\"EmployeeTaskView\"]/Group[@Name=\"RibbonControl\"][@AutomationId=\"ribbonControl\"]/Pane[@ClassName=\"RibbonSelectedPageControl\"][@Name=\"Lower Ribbon\"]/Group[@ClassName=\"RibbonPageGroupControl\"][@Name=\"Save\"]/Button[@Name=\"Save\"][@AutomationId=\"xF0A2A8AA912F28E23E49FCAD8DF0C0B9\"]";
            var winElem_LeftClickButtonSave_21_61 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickButtonSave_21_61);
            if (winElem_LeftClickButtonSave_21_61 != null)
            {
                winElem_LeftClickButtonSave_21_61.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickButtonSave_21_61}");
            }


            // LeftClick on Button "Close" at (19,69)
            Console.WriteLine("LeftClick on Button \"Close\" at (19,69)");
            string xpath_LeftClickButtonClose_19_69 = "/Pane[@ClassName=\"#32769\"][@Name=\"Desktop 1\"]/Window[@ClassName=\"Window\"][@Name=\"Employees - DevAV\"]/Window[@ClassName=\"Window\"][@Name=\"Call  Jane\"]/Custom[@ClassName=\"EmployeeTaskView\"]/Group[@Name=\"RibbonControl\"][@AutomationId=\"ribbonControl\"]/Pane[@ClassName=\"RibbonSelectedPageControl\"][@Name=\"Lower Ribbon\"]/Group[@ClassName=\"RibbonPageGroupControl\"][@Name=\"Close\"]/Button[@Name=\"Close\"][@AutomationId=\"x5073E1AD7C525B2DB4CD533F9C83448F\"]";
            var winElem_LeftClickButtonClose_19_69 = desktopSession.FindElementByAbsoluteXPath(xpath_LeftClickButtonClose_19_69);
            if (winElem_LeftClickButtonClose_19_69 != null)
            {
                winElem_LeftClickButtonClose_19_69.Click();
            }
            else
            {
                Assert.Fail($"Failed to find element using xpath: {xpath_LeftClickButtonClose_19_69}");
            }




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