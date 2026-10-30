using Microsoft.Win32;
using Selenium.Internal;
using Selenium.Serializer;
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;

namespace Selenium.Core {

    /// <summary>
    /// Finds the driver matching the installed browser, and downloads it when missing.
    /// </summary>
    /// <remarks>
    /// Lookup order: the driver next to this assembly when its version matches the browser,
    /// then the drivers cache (%LOCALAPPDATA%\SeleniumBasic\drivers), then a download into the cache.
    /// </remarks>
    static class DriverManager {

        const string CHROME_DRIVERS_URL = "https://googlechromelabs.github.io/chrome-for-testing/latest-patch-versions-per-build-with-downloads.json";
        const string EDGE_DRIVER_URL = "https://msedgedriver.microsoft.com/{0}/edgedriver_{1}.zip";
        const string GECKO_RELEASE_URL = "https://api.github.com/repos/mozilla/geckodriver/releases/latest";

        const int DOWNLOAD_TIMEOUT = 120000;

        /// <summary>
        /// Returns the path of chromedriver for the installed Chrome (or the given binary).
        /// </summary>
        public static string GetChromeDriver(string browserBinary) {
            return GetDriver("chromedriver.exe", "chrome", "chrome.exe"
                , @"Google\Chrome\Application\chrome.exe", browserBinary, DownloadChromeDriver);
        }

        /// <summary>
        /// Returns the path of msedgedriver for the installed Edge (or the given binary).
        /// </summary>
        public static string GetEdgeDriver(string browserBinary) {
            return GetDriver("msedgedriver.exe", "edge", "msedge.exe"
                , @"Microsoft\Edge\Application\msedge.exe", browserBinary, DownloadEdgeDriver);
        }

        /// <summary>
        /// Returns the path of geckodriver. Any version works with a recent Firefox.
        /// </summary>
        public static string GetGeckoDriver() {
            const string filename = "geckodriver.exe";

            string local = Path.Combine(IOExt.GetAssemblyDirectory(), filename);
            if (File.Exists(local))
                return local;

            string cacheDir = GetCacheDirectory("gecko");
            string cached = FindNewestCachedDriver(cacheDir, filename);
            if (cached != null)
                return cached;

            return WithDownloadLock(() => {
                cached = FindNewestCachedDriver(cacheDir, filename);
                return cached ?? DownloadGeckoDriver(cacheDir, filename);
            });
        }


        #region Lookup

        delegate string Downloader(string browserVersion, string targetDir, string filename);

        private static string GetDriver(string filename, string cacheName, string appName
            , string defaultInstallPath, string browserBinary, Downloader download) {

            string local = Path.Combine(IOExt.GetAssemblyDirectory(), filename);
            string browserVersion = GetBrowserVersion(appName, defaultInstallPath, browserBinary);

            // Browser not found: keep the previous behaviour and let the driver report it
            if (browserVersion == null) {
                if (File.Exists(local))
                    return local;
                throw new SeleniumError("Browser not found: {0}", appName);
            }

            if (File.Exists(local) && IsCompatible(GetDriverVersion(local), browserVersion))
                return local;

            string cacheDir = Path.Combine(GetCacheDirectory(cacheName), browserVersion);
            string cached = Path.Combine(cacheDir, filename);
            if (File.Exists(cached))
                return cached;

            try {
                return WithDownloadLock(() => {
                    if (File.Exists(cached))
                        return cached;
                    return download(browserVersion, cacheDir, filename);
                });
            } catch (Exception ex) {
                // Offline or no matching driver: fall back to the bundled driver if any
                if (File.Exists(local))
                    return local;
                throw new SeleniumError("Failed to get {0} for {1} {2}: {3}"
                    , filename, appName, browserVersion, ex.Message);
            }
        }

        /// <summary>
        /// Chromium drivers only support the browser of the same major version.
        /// </summary>
        private static bool IsCompatible(string driverVersion, string browserVersion) {
            if (driverVersion == null)
                return false;
            return GetMajor(driverVersion) == GetMajor(browserVersion);
        }

        private static string GetMajor(string version) {
            int i = version.IndexOf('.');
            return i == -1 ? version : version.Substring(0, i);
        }

        /// <summary>
        /// Returns the version of the browser executable, or null if not installed.
        /// </summary>
        private static string GetBrowserVersion(string appName, string defaultInstallPath, string browserBinary) {
            string path = browserBinary;
            if (path == null) {
                path = GetAppPath(Registry.LocalMachine, appName) ?? GetAppPath(Registry.CurrentUser, appName);
            }
            if (path == null) {
                foreach (var folder in new[] {
                    Environment.SpecialFolder.ProgramFiles,
                    Environment.SpecialFolder.ProgramFilesX86,
                    Environment.SpecialFolder.LocalApplicationData }) {
                    string candidate = Path.Combine(Environment.GetFolderPath(folder), defaultInstallPath);
                    if (File.Exists(candidate)) {
                        path = candidate;
                        break;
                    }
                }
            }
            if (path == null || !File.Exists(path))
                return null;
            return FileVersionInfo.GetVersionInfo(path).ProductVersion;
        }

        private static string GetAppPath(RegistryKey root, string appName) {
            using (var key = root.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" + appName)) {
                if (key == null)
                    return null;
                string path = key.GetValue(null) as string;
                return string.IsNullOrEmpty(path) ? null : path.Trim('"', ' ');
            }
        }

        /// <summary>
        /// Returns the version printed by "driver --version", or null.
        /// </summary>
        private static string GetDriverVersion(string driverPath) {
            try {
                var info = new ProcessStartInfo(driverPath, "--version") {
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };
                using (var p = Process.Start(info)) {
                    string output = p.StandardOutput.ReadToEnd();
                    p.WaitForExit(5000);
                    var match = Regex.Match(output, @"\d+(\.\d+)+");
                    return match.Success ? match.Value : null;
                }
            } catch {
                return null;
            }
        }

        private static string GetCacheDirectory(string name) {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                , @"SeleniumBasic\drivers\" + name);
        }

        private static string FindNewestCachedDriver(string cacheDir, string filename) {
            if (!Directory.Exists(cacheDir))
                return null;
            string newest = null;
            DateTime newestTime = DateTime.MinValue;
            foreach (string dir in Directory.GetDirectories(cacheDir)) {
                string path = Path.Combine(dir, filename);
                if (File.Exists(path) && File.GetLastWriteTimeUtc(path) > newestTime) {
                    newest = path;
                    newestTime = File.GetLastWriteTimeUtc(path);
                }
            }
            return newest;
        }

        /// <summary>
        /// Serializes the downloads between processes (Excel instances, scripts...).
        /// </summary>
        private static string WithDownloadLock(Func<string> action) {
            using (var mutex = new Mutex(false, @"Local\SeleniumBasic.DriverDownload")) {
                bool owned = false;
                try {
                    try {
                        owned = mutex.WaitOne(DOWNLOAD_TIMEOUT);
                    } catch (AbandonedMutexException) {
                        owned = true;
                    }
                    return action();
                } finally {
                    if (owned)
                        mutex.ReleaseMutex();
                }
            }
        }

        #endregion


        #region Downloads

        private static string Platform {
            get {
                return Environment.Is64BitOperatingSystem ? "win64" : "win32";
            }
        }

        private static string DownloadChromeDriver(string browserVersion, string targetDir, string filename) {
            // The drivers are published per build: 156.0.8078.79 -> "156.0.8078"
            string build = Regex.Match(browserVersion, @"^\d+\.\d+\.\d+").Value;

            var versions = (Dictionary)DownloadJson(CHROME_DRIVERS_URL);
            Dictionary builds = (Dictionary)versions["builds"];
            Dictionary release;
            if (!builds.TryGetValue(build, out release))
                throw new SeleniumError("No chromedriver published for Chrome {0}", browserVersion);

            var downloads = (Dictionary)release["downloads"];
            List drivers;
            if (!downloads.TryGetValue("chromedriver", out drivers))
                throw new SeleniumError("No chromedriver published for Chrome {0}", browserVersion);

            foreach (Dictionary item in drivers) {
                if ((string)item["platform"] == Platform)
                    return DownloadAndExtract((string)item["url"], targetDir, filename);
            }
            throw new SeleniumError("No chromedriver published for {0}", Platform);
        }

        private static string DownloadEdgeDriver(string browserVersion, string targetDir, string filename) {
            // msedgedriver is published for each Edge version
            string url = string.Format(EDGE_DRIVER_URL, browserVersion, Platform == "win64" ? "win64" : "win32");
            return DownloadAndExtract(url, targetDir, filename);
        }

        private static string DownloadGeckoDriver(string cacheDir, string filename) {
            var release = (Dictionary)DownloadJson(GECKO_RELEASE_URL);
            string version = (string)release["tag_name"];
            string suffix = "-" + Platform + ".zip";
            foreach (Dictionary asset in (List)release["assets"]) {
                string name = (string)asset["name"];
                if (name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) {
                    string url = (string)asset["browser_download_url"];
                    return DownloadAndExtract(url, Path.Combine(cacheDir, version), filename);
                }
            }
            throw new SeleniumError("No geckodriver published for {0}", Platform);
        }

        private static object DownloadJson(string url) {
            using (var response = CreateRequest(url).GetResponse())
            using (var stream = response.GetResponseStream()) {
                return JSON.Parse(stream);
            }
        }

        /// <summary>
        /// Downloads a zip and extracts the driver into the target directory.
        /// </summary>
        private static string DownloadAndExtract(string url, string targetDir, string filename) {
            string tempZip = Path.Combine(DriverService.GetTempFolder(), Guid.NewGuid().ToString("N") + ".zip");
            string tempExe = null;
            try {
                using (var response = CreateRequest(url).GetResponse())
                using (var stream = response.GetResponseStream())
                using (var file = File.Create(tempZip)) {
                    stream.CopyTo(file);
                }

                Directory.CreateDirectory(targetDir);
                string target = Path.Combine(targetDir, filename);
                tempExe = target + ".tmp";

                using (var zip = System.IO.Compression.ZipFile.OpenRead(tempZip)) {
                    ZipArchiveEntry entry = null;
                    foreach (var e in zip.Entries) {
                        if (string.Equals(e.Name, filename, StringComparison.OrdinalIgnoreCase)) {
                            entry = e;
                            break;
                        }
                    }
                    if (entry == null)
                        throw new SeleniumError("{0} not found in {1}", filename, url);
                    entry.ExtractToFile(tempExe, true);
                }

                // Moved in place last so an interrupted download never leaves a broken driver
                if (File.Exists(target))
                    File.Delete(target);
                File.Move(tempExe, target);
                tempExe = null;
                return target;
            } finally {
                TryDelete(tempZip);
                if (tempExe != null)
                    TryDelete(tempExe);
            }
        }

        private static HttpWebRequest CreateRequest(string url) {
            RemoteServer.EnableTls12();
            var request = (HttpWebRequest)WebRequest.Create(url);
            request.UserAgent = "SeleniumBasic";  // required by the GitHub API
            request.Timeout = DOWNLOAD_TIMEOUT;
            request.ReadWriteTimeout = DOWNLOAD_TIMEOUT;
            // RemoteServer disables the default proxy for local drivers, use the system one here
            request.Proxy = WebRequest.GetSystemWebProxy();
            return request;
        }

        private static void TryDelete(string path) {
            try {
                if (File.Exists(path))
                    File.Delete(path);
            } catch { }
        }

        #endregion

    }

}
