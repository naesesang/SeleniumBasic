using Selenium.Core;
using System;
using System.Net;
using System.Runtime.InteropServices;

namespace Selenium {

    /// <summary>
    /// Web driver for PhantomJS (Headless browser)
    /// </summary>
    /// <example>
    /// 
    /// VBScript:
    /// <code lang="vbs">	
    /// Class Script
    ///     Dim driver
    ///     
    ///     Sub Class_Initialize
    ///         Set driver = CreateObject("Selenium.PhantomJSDriver")
    ///         driver.Get "http://www.google.com"
    ///     End Sub
    /// 
    ///     Sub Class_Terminate
    ///         driver.Quit
    ///     End Sub
    /// End Class
    /// 
    /// Set s = New Script
    /// </code>
    /// 
    /// VBA:
    /// <code lang="vbs">	
    /// Public Sub Script()
    ///   Dim driver As New PhantomJSDriver
    ///   driver.Get "http://www.google.com"
    ///   ...
    ///   driver.Quit
    /// End Sub
    /// </code>
    /// 
    /// </example>
    [ProgId("Selenium.PhantomJSDriver")]
    [Guid("0277FC34-FD1B-4616-BB19-0809389E78C4")]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class PhantomJSDriver : WebDriver, ComInterfaces._WebDriver {

        const string BROWSER_NAME = "phantomjs";

        public PhantomJSDriver()
            : base(BROWSER_NAME) { }

        internal static IDriverService StartService(WebDriver wd) {
            // PhantomJS is discontinued and only speaks the legacy JSON Wire protocol.
            // The class is kept so that existing VBA references still compile.
            throw new SeleniumError("PhantomJS is no longer supported. Use ChromeDriver or EdgeDriver"
                + " with AddArgument \"--headless=new\" for a headless browser.");
        }

        internal static void ExtendCapabilities(WebDriver wd, bool remote) {
            var capa = wd.Capabilities;
            capa.BrowserName = "phantomjs";
            if (wd.Arguments.Count > 0)
                capa["phantomjs.cli.args"] = wd.Arguments;
        }

    }

}
