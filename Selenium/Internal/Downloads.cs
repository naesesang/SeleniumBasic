using System;
using System.IO;
using System.Runtime.InteropServices;

namespace Selenium.Internal {

    /// <summary>
    /// Waits for the files downloaded by the browser.
    /// </summary>
    static class Downloads {

        // Temporary files of the downloads in progress (Chrome/Edge, Firefox, Firefox/Edge)
        static readonly string[] PARTIAL_EXTENSIONS = { ".crdownload", ".part", ".tmp" };

        /// <summary>
        /// Returns the user's Downloads folder, the browsers' default.
        /// </summary>
        public static string GetDefaultFolder() {
            IntPtr pathPtr;
            Guid downloadsId = new Guid("374DE290-123F-4565-9164-39C4925E467B");
            if (SHGetKnownFolderPath(ref downloadsId, 0, IntPtr.Zero, out pathPtr) == 0) {
                try {
                    return Marshal.PtrToStringUni(pathPtr);
                } finally {
                    Marshal.FreeCoTaskMem(pathPtr);
                }
            }
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
        }

        /// <summary>
        /// Waits for a file matching the pattern to be completely downloaded after the given time.
        /// </summary>
        /// <param name="folder">Download folder</param>
        /// <param name="pattern">File name or wildcard pattern, all files when null</param>
        /// <param name="sinceUtc">Only the files written after this time are considered</param>
        /// <param name="timeout">Timeout in milliseconds</param>
        /// <returns>Full path of the downloaded file</returns>
        public static string WaitForFile(string folder, string pattern, DateTime sinceUtc, int timeout) {
            if (string.IsNullOrEmpty(pattern))
                pattern = "*";
            DateTime endTime = DateTime.UtcNow.AddMilliseconds(timeout);
            string candidate = null;
            long candidateSize = -1;

            while (true) {
                string found = null;
                bool pending = false;
                DateTime foundTime = DateTime.MinValue;

                if (Directory.Exists(folder)) {
                    foreach (string path in Directory.GetFiles(folder)) {
                        DateTime writeTime;
                        try {
                            writeTime = File.GetLastWriteTimeUtc(path);
                        } catch (IOException) {
                            continue;
                        }
                        if (writeTime < sinceUtc)
                            continue;
                        if (IsPartial(path)) {
                            pending = true;
                            continue;
                        }
                        if (!MatchesPattern(Path.GetFileName(path), pattern))
                            continue;
                        if (writeTime > foundTime) {
                            found = path;
                            foundTime = writeTime;
                        }
                    }
                }

                // Done when the newest match keeps the same size between two polls
                // and no other download is still in progress
                if (found != null && !pending) {
                    long size = GetSize(found);
                    if (found == candidate && size == candidateSize && size >= 0 && IsReadable(found))
                        return found;
                    candidate = found;
                    candidateSize = size;
                } else {
                    candidate = null;
                    candidateSize = -1;
                }

                if (DateTime.UtcNow > endTime) {
                    throw new Errors.TimeoutError(
                        "No download of \"{0}\" completed in {1} within {2} seconds."
                        , pattern, folder, timeout / 1000);
                }
                System.Threading.Thread.Sleep(250);
            }
        }

        private static bool IsPartial(string path) {
            string ext = Path.GetExtension(path);
            foreach (string partial in PARTIAL_EXTENSIONS) {
                if (string.Equals(ext, partial, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        private static bool MatchesPattern(string fileName, string pattern) {
            string regex = "^" + System.Text.RegularExpressions.Regex.Escape(pattern)
                .Replace(@"\*", ".*").Replace(@"\?", ".") + "$";
            return System.Text.RegularExpressions.Regex.IsMatch(fileName, regex
                , System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }

        private static long GetSize(string path) {
            try {
                return new FileInfo(path).Length;
            } catch (IOException) {
                return -1;
            }
        }

        private static bool IsReadable(string path) {
            try {
                using (File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
                    return true;
                }
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            }
        }

        [DllImport("shell32.dll")]
        private static extern int SHGetKnownFolderPath(ref Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr pszPath);

    }

}
