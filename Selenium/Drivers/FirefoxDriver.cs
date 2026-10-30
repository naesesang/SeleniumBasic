using Selenium.Core;
using Selenium.Internal;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace Selenium {

    /// <summary>
    /// Web driver for Firefox
    /// </summary>
    /// 
    /// <example>
    /// 
    /// VBScript:
    /// <code lang="vbs">	
    /// Class Script
    ///     Dim driver
    ///     
    ///     Sub Class_Initialize
    ///         Set driver = CreateObject("Selenium.FirefoxDriver")
    ///         driver.Get "http://www.google.com"
    ///     End Sub
    /// 
    ///     Sub Class_Terminate
    ///         driver.Quit
    ///         ...
    ///     End Sub
    /// End Class
    /// 
    /// Set s = New Script
    /// </code>
    /// 
    /// VBA:
    /// <code lang="vbs">	
    /// Public Sub Script()
    ///   Dim driver As New FirefoxDriver
    ///   driver.Get "http://www.mywebsite.com"
    ///   ...
    ///   driver.Quit
    /// End Sub
    /// </code>
    /// 
    /// </example>
    [ProgId("Selenium.FirefoxDriver")]
    [Guid("0277FC34-FD1B-4616-BB19-14DB1E4916D4")]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class FirefoxDriver : WebDriver, ComInterfaces._WebDriver {

        const string BROWSER_NAME = "firefox";
        const string APP_PROFILES_FOLDER = @"Mozilla\Firefox";

        public FirefoxDriver()
            : base(BROWSER_NAME) { }

        internal static IDriverService StartService(WebDriver wd) {
            ExtendCapabilities(wd, false);

            var svc = new DriverService();
            svc.AddArgument("--port=" + svc.IPEndPoint.Port.ToString());
            svc.Start("geckodriver.exe");
            return svc;
        }

        internal static void ExtendCapabilities(WebDriver wd, bool remote) {
            Capabilities capa = wd.Capabilities;

            Dictionary opts;
            if (!capa.TryGetValue("moz:firefoxOptions", out opts))
                capa["moz:firefoxOptions"] = opts = new Dictionary();

            var args = new List();
            foreach (object arg in wd.Arguments)
                args.Add(arg);

            if (wd.Profile != null) {
                args.Add("-profile");
                args.Add(remote ? wd.Profile : ExpandProfile(wd));
            }

            if (args.Count != 0)
                opts["args"] = args;

            if (wd.Preferences.Count != 0)
                opts["prefs"] = wd.Preferences;

            string binary = wd.Binary;
            if (binary == null)
                capa.TryGetValue("firefox_binary", out binary);
            if (binary != null)
                opts["binary"] = binary;
        }

        /// <summary>
        /// Installs the extensions once the session is started (no capability for it with geckodriver).
        /// </summary>
        internal static void InstallExtensions(WebDriver wd, RemoteSession session) {
            foreach (DriverExtension ext in wd.Extensions)
                session.Send(RequestMethod.POST, "/moz/addon/install", "addon", ext, "temporary", true);
        }

        /// <summary>
        /// Returns the profile directory from a path or from the name of an existing profile.
        /// </summary>
        private static string ExpandProfile(WebDriver wd) {
            string profile = wd.Profile;
            if (IOExt.IsPath(profile)) {
                profile = IOExt.ExpandPath(profile);
                Directory.CreateDirectory(profile);
                return profile;
            }

            string appdata_dir;
            if (!wd.Capabilities.TryGetValue("firefox_appdata", out appdata_dir))
                appdata_dir = Path.Combine(IOExt.AppDataFolder, APP_PROFILES_FOLDER);

            string profile_dir;
            if (!Directory.Exists(appdata_dir) || !ParseExistingProfiles(appdata_dir).TryGetValue(profile, out profile_dir))
                throw new SeleniumError("Firefox profile not found: {0}", profile);
            return profile_dir;
        }

        /// <summary>
        /// Extract the name and path for each profile define in the provided INI file
        /// </summary>
        /// <param name="profiles_dir">Folder of the INI file (ApplicationData\Mozilla\Firefox)</param>
        /// <returns>Dictionary Key:name Value:path</returns>
        private static Dictionary ParseExistingProfiles(string profiles_dir) {
            Dictionary profiles = new Dictionary();
            var inifile = IniFile.Load(profiles_dir + @"\profiles.ini");
            foreach (var section in inifile) {
                if (!section.Key.StartsWith(@"Profile", StringComparison.OrdinalIgnoreCase))
                    continue;

                var properties = section.Value;
                string path = properties["Path"].Replace('/', '\\');
                if (properties["IsRelative"] == "1")
                    path = Path.Combine(profiles_dir, path);
                profiles.Add(properties["Name"], path);
            }
            return profiles;
        }

    }

}
