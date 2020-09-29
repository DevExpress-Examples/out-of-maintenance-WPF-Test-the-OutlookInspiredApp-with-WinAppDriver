
//******************************************************************************
//
// Copyright (c) 2019 Microsoft Corporation. All rights reserved.
//
// This code is licensed under the MIT License (MIT).
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//
//******************************************************************************

using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace outlookinspiredapp.uitest
{
    public class DesktopSession
    {
        const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723/";
        WindowsDriver<WindowsElement> desktopSession;
        public DesktopSession(WindowsDriver<WindowsElement> source)
        {
            desktopSession = source;
        }

        public WindowsDriver<WindowsElement> DesktopSessionElement
        {
            get { return desktopSession; }
        }

        public WindowsElement FindElementByAbsoluteXPath(string xPath, int nTryCount = 10)
        {
            WindowsElement uiTarget = null;
            var index = xPath.IndexOf(value: '/', startIndex: 1);
            xPath = xPath.Substring(startIndex: index);
            while (nTryCount-- > 0)
            {
                try
                {
                    uiTarget = desktopSession.FindElementByXPath(xpath: $"/{xPath}");
                }
                catch {
                    Console.WriteLine($@"Find failed: ""{xPath}""");
                }

                if (uiTarget != null)
                    break;
                Thread.Sleep(millisecondsTimeout: 100);
            }

            return uiTarget;
        }

        public IOptions Manage()
        {
            return this.desktopSession.Manage();
        }

        public void CloseApp()
        {
            this.desktopSession.CloseApp();
        }
    }

}
