using Selenium;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Selenium.ComInterfaces {
#pragma warning disable 1591

    [Guid("0277FC34-FD1B-4616-BB19-01D514FE0B1A")]
    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Description("")]
    public interface _Utils {

        [DispId(265), Description("Get a screenshot of the desktop")]
        Image TakeScreenShot();

        [DispId(403), Description("Sends keystrokes to the active application using the windows SendKeys methode.")]
        void SendKeysNat(string keys);

        [DispId(407), Description("Indicates whether the regular expression finds a match in the input string using the regular expression specified in the pattern parameter.")]
        bool IsMatch(string input, string pattern);

        [DispId(410), Description("Searches the specified input string for an occurrence of the regular expression supplied in the pattern parameter.")]
        string Match(string input, string pattern, short group = 0);

        [DispId(413), Description("Searches the specified input string for all the occurrences of the regular expression supplied in the pattern parameter.")]
        List Matches(string input, string pattern, short group = 0);

        [DispId(423), Description("Searches the specified input string for an occurrence of the regular expression supplied in the pattern parameter and replace it.")]
        string Replace(string input, string pattern, string replacement);

        [DispId(430), Description("Parses a JSON string into Dictionary and List objects.")]
        object ParseJson(string json);

        [DispId(433), Description("Converts a value (Dictionary, List, array, string, number...) to a JSON string.")]
        string ToJson(object value);

        [DispId(436), Description("Decodes a Base64 string and saves it to a file. Returns the full path of the file.")]
        string SaveBase64ToFile(string base64, string filePath);

        [DispId(439), Description("Reads a file and returns its content as a Base64 string.")]
        string FileToBase64(string filePath);

    }

}
