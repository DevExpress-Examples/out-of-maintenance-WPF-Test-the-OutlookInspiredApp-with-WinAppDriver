'******************************************************************************
'
' Copyright (c) 2019 Microsoft Corporation. All rights reserved.
'
' This code is licensed under the MIT License (MIT).
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.
'
'******************************************************************************

Imports OpenQA.Selenium
Imports OpenQA.Selenium.Appium.Windows
Imports OpenQA.Selenium.Remote
Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Threading

Namespace outlookinspiredapp.uitest
	Public Class DesktopSession
		Private Const WindowsApplicationDriverUrl As String = "http://127.0.0.1:4723/"
'INSTANT VB NOTE: The field desktopSession was renamed since Visual Basic does not allow fields to have the same name as other class members:
		Private desktopSession_Conflict As WindowsDriver(Of WindowsElement)
		Public Sub New(ByVal source As WindowsDriver(Of WindowsElement))
			desktopSession_Conflict = source
		End Sub

		Public ReadOnly Property DesktopSessionElement() As WindowsDriver(Of WindowsElement)
			Get
				Return desktopSession_Conflict
			End Get
		End Property

		Public Function FindElementByAbsoluteXPath(ByVal xPath As String, Optional ByVal nTryCount As Integer = 10) As WindowsElement
			Dim uiTarget As WindowsElement = Nothing
			Dim index = xPath.IndexOf(value:= "/"c, startIndex:= 1)
			xPath = xPath.Substring(startIndex:= index)
'INSTANT VB WARNING: An assignment within expression was extracted from the following statement:
'ORIGINAL LINE: while (nTryCount-- > 0)
			Do While nTryCount > 0
				nTryCount -= 1
				Try
					uiTarget = desktopSession_Conflict.FindElementByXPath(xpath:= $"/{xPath}")
				Catch
					Console.WriteLine($"Find failed: ""{xPath}""")
				End Try

				If uiTarget IsNot Nothing Then
					Exit Do
				End If
				Thread.Sleep(millisecondsTimeout:= 100)
			Loop
			nTryCount -= 1

			Return uiTarget
		End Function

		Public Function Manage() As IOptions
			Return Me.desktopSession_Conflict.Manage()
		End Function

		Public Sub CloseApp()
			Me.desktopSession_Conflict.CloseApp()
		End Sub
	End Class

End Namespace
