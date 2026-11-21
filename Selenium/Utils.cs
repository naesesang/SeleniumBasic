using Selenium.Internal;
using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Selenium {

    /// <summary>
    /// 
    /// </summary>
    [ProgId("Selenium.Utils")]
    [Guid("0277FC34-FD1B-4616-BB19-A34FCBA29598")]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class Utils : ComInterfaces._Utils {

        /// <summary>
        /// Takes a screenshoot of the desktop.
        /// </summary>
        /// <returns></returns>
        public Image TakeScreenShot() {
            var bounds = System.Windows.Forms.Screen.GetBounds(System.Drawing.Point.Empty);
            var bitmap = new System.Drawing.Bitmap(bounds.Width, bounds.Height, System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            using (var g = System.Drawing.Graphics.FromImage(bitmap)) {
                g.CopyFromScreen(System.Drawing.Point.Empty, System.Drawing.Point.Empty, bounds.Size);
                return new Image(bitmap);
            }
        }

        /// <summary>
        /// Sends keystrokes to the active application using the windows SendKeys methode.
        /// </summary>
        /// <param name="keys">The string of keystrokes to Send.</param>
        public void SendKeysNat(string keys) {
            System.Windows.Forms.SendKeys.Send(keys);
        }

        /// <summary>
        /// Indicates whether the regular expression finds a match in the input string using the regular expression specified in the pattern parameter.
        /// </summary>
        /// <param name="input">The string to search for a match.</param>
        /// <param name="pattern">The regular expression pattern to match.</param>
        /// <returns>true if the regular expression finds a match; otherwise, false.</returns>
        public bool IsMatch(string input, string pattern) {
            return Regex.IsMatch(input, pattern);
        }

        /// <summary>
        /// Searches the specified input string for an occurrence of the regular expression supplied in the pattern parameter.
        /// </summary>
        /// <param name="input">The string to search for a match.</param>
        /// <param name="pattern">The regular expression pattern to match.</param>
        /// <param name="group">Group value to return</param>
        /// <returns>Matching string or groups </returns>
        public string Match(string input, string pattern, short group = 0) {
            return StringExt.Match(input, pattern, group);
        }

        /// <summary>
        /// Searches the specified input string for an occurrence of the regular expression supplied in the pattern parameter.
        /// </summary>
        /// <param name="input">The string to search for a match.</param>
        /// <param name="pattern">The regular expression pattern to match.</param>
        /// <param name="group">Group value to return</param>
        /// <returns></returns>
        public List Matches(string input, string pattern, short group = 0) {
            return StringExt.Matches(input, pattern, group);
        }

        /// <summary>
        /// Within a specified input string, replaces all strings that match a specified regular expression with a specified replacement string.
        /// </summary>
        /// <param name="input">The string to search for a match.</param>
        /// <param name="pattern">The regular expression pattern to match.</param>
        /// <param name="replacement">The replacement string.</param>
        /// <returns>A new string that is identical to the input string, except that a replacement string takes the place of each matched string.</returns>
        public string Replace(string input, string pattern, string replacement) {
            return Regex.Replace(input, pattern, replacement);
        }

        /// <summary>
        /// Parses a JSON string into Dictionary and List objects.
        /// </summary>
        public object ParseJson(string json) {
            return WebDriver.ParseUserJson(json);
        }

        /// <summary>
        /// Converts a value (Dictionary, List, array, string, number...) to a JSON string.
        /// </summary>
        public string ToJson(object value) {
            return Serializer.JSON.Serialize(value, 64).ToString();
        }

        /// <summary>
        /// Decodes a Base64 string (optionally a data URL) and saves it to a file.
        /// </summary>
        /// <returns>Full path of the file</returns>
        public string SaveBase64ToFile(string base64, string filePath) {
            if (base64 == null)
                throw new Errors.ArgumentError("The Base64 string is empty.");
            int comma = base64.IndexOf(',');
            if (base64.StartsWith("data:", StringComparison.OrdinalIgnoreCase) && comma != -1)
                base64 = base64.Substring(comma + 1);
            filePath = System.IO.Path.GetFullPath(IOExt.ExpandPath(filePath));
            string dir = System.IO.Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir))
                System.IO.Directory.CreateDirectory(dir);
            System.IO.File.WriteAllBytes(filePath, Convert.FromBase64String(base64.Trim()));
            return filePath;
        }

        /// <summary>
        /// Reads a file and returns its content as a Base64 string.
        /// </summary>
        public string FileToBase64(string filePath) {
            filePath = System.IO.Path.GetFullPath(IOExt.ExpandPath(filePath));
            if (!System.IO.File.Exists(filePath))
                throw new Errors.FileNotFoundError(filePath);
            return Convert.ToBase64String(System.IO.File.ReadAllBytes(filePath));
        }

    }

}
