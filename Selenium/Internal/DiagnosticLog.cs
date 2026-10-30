using System;
using System.IO;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;

namespace Selenium.Internal {

    /// <summary>
    /// Opt-in diagnostic log for problems happening inside the host process (Excel...).
    /// Enabled only when %LOCALAPPDATA%\SeleniumBasic\debug.log exists: create an empty file to enable it.
    /// Uses mscorlib only, so it still works when another framework assembly fails to load.
    /// </summary>
    static class DiagnosticLog {

        private static readonly object _lock = new object();
        private static string _path = null;
        private static bool _initialized = false;

        public static void Initialize() {
            lock (_lock) {
                if (_initialized)
                    return;
                _initialized = true;

                string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                    , @"SeleniumBasic\debug.log");
                if (!File.Exists(path))
                    return;
                _path = path;
            }

            AppDomain domain = AppDomain.CurrentDomain;
            Write("---- session " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
            Write("process: " + Environment.CommandLine + " (64bit=" + Environment.Is64BitProcess + ")");
            Write("runtime: " + RuntimeEnvironment.GetSystemVersion() + " " + RuntimeEnvironment.GetRuntimeDirectory());
            Write("appdomain: " + domain.FriendlyName + " base=" + domain.BaseDirectory
                + " config=" + domain.SetupInformation.ConfigurationFile);
            Write("selenium: " + typeof(DiagnosticLog).Assembly.FullName + " @ " + typeof(DiagnosticLog).Assembly.Location);
            foreach (Assembly assembly in domain.GetAssemblies())
                Write("loaded: " + Describe(assembly));

            domain.AssemblyLoad += (sender, e) => Write("load: " + Describe(e.LoadedAssembly));
            domain.AssemblyResolve += (sender, e) => {
                Write("resolve: " + e.Name + " requested by "
                    + (e.RequestingAssembly == null ? "?" : e.RequestingAssembly.FullName));
                return null;
            };
            domain.FirstChanceException += OnFirstChanceException;
        }

        private static void OnFirstChanceException(object sender, FirstChanceExceptionEventArgs e) {
            Exception ex = e.Exception;
            string text = "exception: " + ex.GetType().FullName + " 0x" + ex.HResult.ToString("X8") + " " + ex.Message;
            var fileLoad = ex as FileLoadException;
            if (fileLoad != null)
                text += "\n  file: " + fileLoad.FileName + "\n  fusion: " + fileLoad.FusionLog;
            var notFound = ex as FileNotFoundException;
            if (notFound != null)
                text += "\n  file: " + notFound.FileName + "\n  fusion: " + notFound.FusionLog;
            var badImage = ex as BadImageFormatException;
            if (badImage != null)
                text += "\n  file: " + badImage.FileName + "\n  fusion: " + badImage.FusionLog;
            text += "\n  at: " + ex.StackTrace;
            Write(text);
        }

        private static string Describe(Assembly assembly) {
            string location;
            try {
                location = assembly.IsDynamic ? "(dynamic)" : assembly.Location;
            } catch (Exception ex) {
                location = "(" + ex.Message + ")";
            }
            return assembly.FullName + " @ " + location + " [" + assembly.ImageRuntimeVersion + "]";
        }

        private static void Write(string line) {
            string path = _path;
            if (path == null)
                return;
            lock (_lock) {
                try {
                    File.AppendAllText(path, DateTime.Now.ToString("HH:mm:ss.fff ") + line + Environment.NewLine);
                } catch { }
            }
        }

    }

}
