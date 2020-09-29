using Microsoft.Win32;
using NUnit.Framework;
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
        public void Test1()
        {
            Assert.Pass();
        }
    }
}