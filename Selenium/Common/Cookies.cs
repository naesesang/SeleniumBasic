using Selenium.Core;
using System;
using System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Selenium {

    /// <summary>
    /// Cookie collection
    /// </summary>
    [ProgId("Selenium.Cookies")]
    [Guid("0277FC34-FD1B-4616-BB19-313080DBB4DA")]
    [Description("Collection of cookies (One-based)")]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class Cookies : List, ComInterfaces._Cookies {

        #region Static methods

        /// <summary>
        /// Get all cookies visible to the current page.
        /// </summary>
        /// <param name="session"></param>
        /// <returns></returns>
        internal static List GetAll(RemoteSession session) {
            return (List)session.Send(RequestMethod.GET, "/cookie");
        }

        /// <summary>
        /// Adds a new cookie.
        /// </summary>
        /// <param name="session"></param>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <param name="path"></param>
        /// <param name="domain"></param>
        /// <param name="expiry"></param>
        /// <param name="secure"></param>
        /// <param name="httpOnly"></param>
        internal static void AddOne(RemoteSession session, string name, string value, string path, string domain, string expiry, bool secure, bool httpOnly, string sameSite = null) {
            var dict = new Dictionary();
            dict.Add("name", name);
            dict.Add("value", value);
            if (!string.IsNullOrEmpty(path))
                dict.Add("path", path);
            if (!string.IsNullOrEmpty(domain))
                dict.Add("domain", domain);
            if (!string.IsNullOrEmpty(expiry)) {
                var dt = (long)(DateTime.Parse(expiry) - Cookie.BASE_TIME).TotalSeconds;
                dict.Add("expiry", dt);
            }
            if (secure == true)
                dict.Add("secure", true);
            if (httpOnly == true)
                dict.Add("httpOnly", true);
            if (!string.IsNullOrEmpty(sameSite))
                dict.Add("sameSite", sameSite);
            session.Send(RequestMethod.POST, "/cookie", "cookie", dict);
        }

        /// <summary>
        /// Delete all cookies visible to the current page.
        /// </summary>
        internal static void DeleteAll(RemoteSession session) {
            session.Send(RequestMethod.DELETE, "/cookie");
        }

        #endregion


        private readonly RemoteSession _session;

        internal Cookies(RemoteSession session, List rawCookies)
            : base(rawCookies.Count) {
            _session = session;

            _count = _items.Length;
            for (int i = 0; i < _count; i++) {
                _items[i] = new Cookie(session, (Dictionary)rawCookies[i]);
            }
        }

        /// <summary>
        /// Returns the cookie at index (Zero based)
        /// </summary>
        /// <param name="index">Zero based index</param>
        /// <returns></returns>
        public new Cookie this[int index] {
            get {
                return (Cookie)base[index];
            }
        }

        /// <summary>
        /// Returns the cookie for the given name
        /// </summary>
        /// <param name="name">Cookie name</param>
        /// <returns></returns>
        public Cookie this[string name] {
            get {
                for (int i = _count; i-- > 0; ) {
                    Cookie item = (Cookie)_items[i];
                    if (string.Compare(name, item.Name, true) == 0)
                        return item;
                }
                throw new SeleniumError("Cookie not found: " + name);
            }
        }

        /// <summary>
        /// Gets an item at the provided index (One based).
        /// </summary>
        /// <param name="identifier">Name or index (based one)</param>
        /// <returns>Cookie</returns>
        Cookie ComInterfaces._Cookies.this[object identifier] {
            get {
                if (identifier is string)
                    return this[(string)identifier];
                if (identifier is short)
                    return this[(short)identifier - 1];
                if (identifier is int)
                    return this[(int)identifier - 1];
                throw new Errors.ArgumentError("Invalid type for argument identifier.");
            }
        }

        /// <summary>
        /// Saves the cookies to a JSON file, to restore a login later with LoadFromFile.
        /// </summary>
        /// <param name="filePath">File path</param>
        public void SaveToFile(string filePath) {
            filePath = System.IO.Path.GetFullPath(Internal.IOExt.ExpandPath(filePath));
            var list = new List(_count);
            for (int i = 0; i < _count; i++)
                list.Add(((Cookie)_items[i]).Raw);
            string dir = System.IO.Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir))
                System.IO.Directory.CreateDirectory(dir);
            System.IO.File.WriteAllText(filePath, Serializer.JSON.Serialize(list, 16).ToString(), new System.Text.UTF8Encoding(false));
        }

        /// <summary>
        /// Adds the cookies of a JSON file saved by SaveToFile to the browser.
        /// Open a page of the same site first: a browser only accepts the cookies of the current domain.
        /// </summary>
        /// <param name="filePath">File path</param>
        /// <returns>Number of cookies added</returns>
        public int LoadFromFile(string filePath) {
            filePath = System.IO.Path.GetFullPath(Internal.IOExt.ExpandPath(filePath));
            if (!System.IO.File.Exists(filePath))
                throw new Errors.FileNotFoundError(filePath);
            var list = WebDriver.ParseUserJson(System.IO.File.ReadAllText(filePath)) as List;
            if (list == null)
                throw new SeleniumError("The file does not contain a list of cookies: {0}", filePath);
            int added = 0;
            foreach (object item in list) {
                var cookie = item as Dictionary;
                if (cookie == null)
                    continue;
                try {
                    _session.Send(RequestMethod.POST, "/cookie", "cookie", cookie);
                    added++;
                } catch (Errors.InvalidCookieDomainError) {
                } catch (Errors.UnableToSetCookieError) {
                } catch (Errors.InvalidArgumentError) {
                }
            }
            return added;
        }

        IEnumerator ComInterfaces._Cookies._NewEnum() {
            return base.GetEnumerator();
        }

    }

}
