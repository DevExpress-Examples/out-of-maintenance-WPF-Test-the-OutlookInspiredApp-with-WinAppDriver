Imports Microsoft.Win32
Imports NUnit.Framework
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Appium
Imports OpenQA.Selenium.Appium.Windows
Imports outlookinspiredapp.uitest
Imports System
Imports System.Diagnostics
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading

Namespace OutlookInspiredApp.UITest
	Public Class Tests
		Private pWad As Process
		Private Const PathToTheDemo As String = "C:\Users\Public\Documents\DevExpress Demos 20.2\Components\WPF\DevExpress.OutlookInspiredApp.Wpf\Bin\DevExpress.OutlookInspiredApp.Wpf.exe"
		Private privatedesktopSession As DesktopSession
		Protected Property desktopSession() As DesktopSession
			Get
				Return privatedesktopSession
			End Get
			Private Set(ByVal value As DesktopSession)
				privatedesktopSession = value
			End Set
		End Property
		<OneTimeSetUp>
		Public Sub FixtureSetup()
			StartWAD()
			Dim options = New AppiumOptions()
			options.AddAdditionalCapability(capabilityName:= "app", capabilityValue:= PathToTheDemo)
			options.AddAdditionalCapability(capabilityName:= "deviceName", capabilityValue:= "WindowsPC")
			options.AddAdditionalCapability(capabilityName:= "platformName", capabilityValue:= "Windows")
			options.AddAdditionalCapability(capabilityName:= "ms:experimental-webdriver", capabilityValue:= True)



			Dim driver = New WindowsDriver(Of WindowsElement)(New Uri("http://127.0.0.1:4723"), options)
			WaitSplashScreen(driver)
			desktopSession = New DesktopSession(driver)
		End Sub

		Private Shared Sub WaitSplashScreen(ByVal driver As WindowsDriver(Of WindowsElement))
			Dim cwh = driver.CurrentWindowHandle
			Do While driver.WindowHandles.Contains(cwh)
				Thread.Sleep(1000)
			Loop
			driver.SwitchTo().Window(driver.WindowHandles(0))
		End Sub

		Private Sub StartWAD()
			Dim psi = New ProcessStartInfo("C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe")
			psi.EnvironmentVariables.Add("DX.UITESTINGENABLED", "1")
			pWad = Process.Start(psi)
		End Sub

		<OneTimeTearDown>
		Public Sub FixtureTearDown()
			desktopSession.CloseApp()
			pWad.Kill()
		End Sub

		<SetUp>
		Public Sub Setup()
		End Sub

		<Test>
		<Order(0)>
		Public Sub CreateEmployee()
			Dim desktopElement = desktopSession.DesktopSessionElement
			Dim bNewEmployee = desktopElement.FindElementByName("New Employee")
			bNewEmployee.Click()

			Dim newEmployeeWindow As WindowsElement = Nothing
			Do While newEmployeeWindow Is Nothing
				newEmployeeWindow = desktopElement.FindElementByName("Employee (New)")
			Loop

			newEmployeeWindow.FindElementByName("First Name").FindElementByClassName("TextEdit").SendKeys("John")
			newEmployeeWindow.FindElementByName("Last Name").FindElementByClassName("TextEdit").SendKeys("Public")
			newEmployeeWindow.FindElementByName("Title").FindElementByClassName("TextEdit").SendKeys("CTO")
			newEmployeeWindow.FindElementByName("Mobile Phone").FindElementByClassName("ButtonEdit").SendKeys("1111111111")
			newEmployeeWindow.FindElementByName("Email").FindElementByClassName("ButtonEdit").SendKeys("john.public@dx-email.com")


			newEmployeeWindow.FindElementByName("Save & Close").Click()
		End Sub
		<Test>
		<Order(1)>
		Public Sub CreateTask()
			Dim desktopElement = desktopSession.DesktopSessionElement
			Dim bNewEmployee = desktopElement.FindElementByName("Task")
			bNewEmployee.Click()

			Dim newTaskWindow As WindowsElement = Nothing
			Do While newTaskWindow Is Nothing
				newTaskWindow = desktopElement.FindElementByName("EmployeeTask (New)")
			Loop

			newTaskWindow.FindElementByName("Subject").FindElementByClassName("TextEdit").SendKeys("Call Jane")
			Dim comboBox = newTaskWindow.FindElementByName("Assigned To").FindElementByClassName("ComboBoxEdit")
			comboBox.FindElementByXPath("//Edit[@ClassName=""ButtonEdit""]/Button[@ClassName=""Button""]").Click()
			comboBox.SendKeys("John Public")

			newTaskWindow.FindElementByName("Save & Close").Click()
		End Sub

		<Test>
		<Order(2)>
		Public Sub CheckTasks()
			Dim desktopElement = desktopSession.DesktopSessionElement
			Dim dockManager = desktopElement.FindElementByName("DockLayoutManager")
			Dim grid = dockManager.FindElementByAccessibilityId("tableViewGridControl")

			grid.FindElementByAccessibilityId("SearchComboBox").SendKeys("John Public")
			grid.FindElementByName("John Public").Click()

			Dim tab = dockManager.FindElementsByClassName("TabItem").First(Function(x) x.GetAttribute("Name") = "Tasks")
			tab.Click()

			Dim row = tab.FindElementByAccessibilityId("Row0")
			Dim callJane = row.FindElementsByClassName("TextEdit").FirstOrDefault(Function(x) x.GetAttribute("Name") = "Call Jane")
			Assert.IsNotNull(callJane)
		End Sub

		<Test>
		<Explicit>
		Public Sub Test1()
			Assert.Pass()
		End Sub
	End Class
End Namespace
