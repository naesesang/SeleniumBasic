using Selenium.Core;
using Selenium.Internal;
using Selenium.Serializer;
using System;
using System.IO;
using System.Reflection;
using System.Text;

namespace Selenium {

    // Features added in 3.2: browser options, page helpers, scrolling, waits, windows and DevTools.
    public partial class WebDriver {

        #region Browser options (before Start)

        private bool _headless = false;
        private bool _incognito = false;
        private bool _guest = false;
        private string _userAgent = null;

        private void EnsureNotStarted(string name) {
            if (_session != null)
                throw new Errors.InvalidOperationError("Call {0} before Start.", name);
        }

        /// <summary>
        /// Runs the browser without a window. Call before Start.
        /// </summary>
        public void SetHeadless(bool enable = true) {
            EnsureNotStarted("SetHeadless");
            _headless = enable;
        }

        /// <summary>
        /// Runs the browser in incognito / InPrivate / private mode. Call before Start.
        /// </summary>
        public void SetIncognito(bool enable = true) {
            EnsureNotStarted("SetIncognito");
            _incognito = enable;
        }

        /// <summary>
        /// Runs Chrome or Edge as a guest, without the user's profile data. Call before Start.
        /// </summary>
        public void SetGuest(bool enable = true) {
            EnsureNotStarted("SetGuest");
            _guest = enable;
        }

        /// <summary>
        /// Replaces the User-Agent sent by the browser. Call before Start.
        /// </summary>
        public void SetUserAgent(string userAgent) {
            EnsureNotStarted("SetUserAgent");
            _userAgent = string.IsNullOrEmpty(userAgent) ? null : userAgent;
        }

        /// <summary>
        /// Hides the "controlled by automated test software" bar in Chrome and Edge. Call before Start.
        /// </summary>
        public void HideAutomationInfo() {
            EnsureNotStarted("HideAutomationInfo");
            if (!this.ExcludeSwitches.Contains("enable-automation"))
                this.ExcludeSwitches.Add("enable-automation");
        }

        /// <summary>
        /// Accepts the invalid or self-signed TLS certificates. Call before Start.
        /// </summary>
        public void SetAcceptInsecureCerts(bool accept = true) {
            EnsureNotStarted("SetAcceptInsecureCerts");
            this.Capabilities["acceptInsecureCerts"] = accept;
        }

        /// <summary>
        /// Sets when Get returns: "normal" (page loaded), "eager" (DOM ready) or "none". Call before Start.
        /// </summary>
        public void SetPageLoadStrategy(string strategy) {
            EnsureNotStarted("SetPageLoadStrategy");
            strategy = (strategy ?? "").Trim().ToLowerInvariant();
            if (strategy != "normal" && strategy != "eager" && strategy != "none")
                throw new Errors.ArgumentError("Invalid page load strategy: {0}. Use normal, eager or none.", strategy);
            this.Capabilities["pageLoadStrategy"] = strategy;
        }

        /// <summary>
        /// Sets how unexpected alerts are handled: "dismiss", "accept", "dismiss and notify",
        /// "accept and notify" or "ignore". Call before Start.
        /// </summary>
        public void SetUnhandledPromptBehavior(string behavior) {
            EnsureNotStarted("SetUnhandledPromptBehavior");
            behavior = (behavior ?? "").Trim().ToLowerInvariant();
            switch (behavior) {
                case "dismiss":
                case "accept":
                case "dismiss and notify":
                case "accept and notify":
                case "ignore":
                    this.Capabilities["unhandledPromptBehavior"] = behavior;
                    break;
                default:
                    throw new Errors.ArgumentError("Invalid prompt behavior: {0}", behavior);
            }
        }

        /// <summary>
        /// Keeps Chrome or Edge open after Quit and after the script ends. Call before Start.
        /// </summary>
        public void SetDetach(bool detach = true) {
            EnsureNotStarted("SetDetach");
            this.Detach = detach;
        }

        /// <summary>
        /// Requests a WebDriver BiDi connection. Read WebSocketUrl after Start. Call before Start.
        /// </summary>
        public void EnableBiDi() {
            EnsureNotStarted("EnableBiDi");
            this.Capabilities["webSocketUrl"] = true;
        }

        /// <summary>
        /// WebDriver BiDi WebSocket URL of the session (requires EnableBiDi before Start), or an empty string.
        /// </summary>
        public string WebSocketUrl {
            get {
                return session.capabilities.Get("webSocketUrl", "") as string ?? "";
            }
        }

        /// <summary>
        /// Adds the browser specific arguments and preferences of the options above.
        /// </summary>
        private void ApplyStartOptions(string browser) {
            bool firefox = browser == "firefox";
            if (_headless)
                AddArgumentOnce(firefox ? "-headless" : "--headless=new");
            if (_incognito)
                AddArgumentOnce(firefox ? "-private" : browser == "MicrosoftEdge" ? "--inprivate" : "--incognito");
            if (_guest && !firefox)
                AddArgumentOnce("--guest");
            if (_userAgent != null) {
                if (firefox) {
                    this.Preferences["general.useragent.override"] = _userAgent;
                } else {
                    for (int i = this.Arguments.Count; i-- > 0; ) {
                        if (((string)this.Arguments[i]).StartsWith("--user-agent="))
                            this.Arguments.RemoveAt(i);
                    }
                    this.Arguments.Add("--user-agent=" + _userAgent);
                }
            }
        }

        private void AddArgumentOnce(string argument) {
            if (!this.Arguments.Contains(argument))
                this.Arguments.Add(argument);
        }

        #endregion


        #region Session information

        /// <summary>
        /// Id of the WebDriver session.
        /// </summary>
        public string SessionId {
            get {
                return session.Id;
            }
        }

        /// <summary>
        /// Status of the driver (ready, message, build...).
        /// </summary>
        public Dictionary Status() {
            var response = session.server.Send(RequestMethod.GET, "/status");
            object value;
            if (response.TryGetValue("value", out value) && value is Dictionary)
                return (Dictionary)value;
            return response;
        }

        /// <summary>
        /// Name of the current frame (window.name), or an empty string for the top page.
        /// </summary>
        public string CurrentFrameName {
            get {
                return session.javascript.Execute("return window.self === window.top ? '' : window.name;", null, false) as string ?? "";
            }
        }

        /// <summary>
        /// User-Agent of the browser.
        /// </summary>
        public string UserAgent {
            get {
                return (string)session.javascript.Execute("return navigator.userAgent;", null, false);
            }
        }

        /// <summary>
        /// HTTP status code of the current page (200, 404...), or 0 when unknown (local file, data URL, older browser).
        /// </summary>
        public int StatusCode {
            get {
                object value = session.javascript.Execute(
                    "var e=performance.getEntriesByType('navigation')[0];"
                    + "return e && e.responseStatus ? e.responseStatus : 0;", null, false);
                return Convert.ToInt32(value);
            }
        }

        #endregion


        #region DevTools

        /// <summary>
        /// Executes a Chrome DevTools Protocol command (Chrome and Edge) and returns its result.
        /// </summary>
        /// <param name="command">Command, ex: "Network.clearBrowserCache"</param>
        /// <param name="parameters">Optional - Dictionary or JSON string, ex: "{""latitude"":37.5}"</param>
        /// <returns>Dictionary with the result</returns>
        public object ExecuteCDP(string command, object parameters = null) {
            return Cdp.Execute(session, command, parameters);
        }

        #endregion


        #region Page

        /// <summary>
        /// Opens an HTML string in the browser.
        /// </summary>
        public void OpenHtml(string html) {
            string data = Convert.ToBase64String(Encoding.UTF8.GetBytes(html ?? ""));
            this.Get("data:text/html;charset=utf-8;base64," + data);
        }

        /// <summary>
        /// Opens a local file in the browser.
        /// </summary>
        public void OpenFile(string path) {
            string fullPath = Path.GetFullPath(IOExt.ExpandPath(path));
            if (!File.Exists(fullPath))
                throw new Errors.FileNotFoundError(fullPath);
            this.Get(new Uri(fullPath).AbsoluteUri);
        }

        /// <summary>
        /// Returns the current page as an HTML document object (MSHTML, like CreateObject("htmlfile")).
        /// The scripts of the page are not run.
        /// </summary>
        public object PageToHtmlDocument() {
            string html = this.PageSource();
            Type type = Type.GetTypeFromProgID("htmlfile", true);
            object doc = Activator.CreateInstance(type);
            type.InvokeMember("designMode", BindingFlags.SetProperty, null, doc, new object[] { "on" });
            type.InvokeMember("write", BindingFlags.InvokeMethod, null, doc, new object[] { html });
            type.InvokeMember("close", BindingFlags.InvokeMethod, null, doc, null);
            return doc;
        }

        /// <summary>
        /// Parses the JSON displayed by the current page (ex: an API URL) into Dictionary and List objects.
        /// </summary>
        public object PageToJson() {
            string text = (string)session.javascript.Execute(
                "var p=document.querySelector('body > pre');"
                + "return p ? p.textContent : document.body ? document.body.innerText : document.documentElement.textContent;"
                , null, false);
            return ParseUserJson(text);
        }

        /// <summary>
        /// Returns the current page as an XML document object (MSXML2.DOMDocument.6.0).
        /// </summary>
        public object PageToXmlDocument() {
            string xml = (string)session.javascript.Execute(
                "return new XMLSerializer().serializeToString(document);", null, false);
            Type type = Type.GetTypeFromProgID("MSXML2.DOMDocument.6.0", true);
            object doc = Activator.CreateInstance(type);
            type.InvokeMember("async", BindingFlags.SetProperty, null, doc, new object[] { false });
            bool loaded = (bool)type.InvokeMember("loadXML", BindingFlags.InvokeMethod, null, doc, new object[] { xml });
            if (!loaded) {
                object error = type.InvokeMember("parseError", BindingFlags.GetProperty, null, doc, null);
                object reason = error.GetType().InvokeMember("reason", BindingFlags.GetProperty, null, error, null);
                throw new SeleniumError("The page is not valid XML: {0}", reason);
            }
            return doc;
        }

        internal static object ParseUserJson(string text) {
            if (string.IsNullOrEmpty(text))
                throw new Errors.ArgumentError("The JSON text is empty.");
            try {
                return JSON.Parse(text.Trim(), 64);
            } catch (Exception ex) {
                throw new SeleniumError("Invalid JSON: {0}", ex.Message);
            }
        }

        /// <summary>
        /// Downloads a URL with the browser's cookies and saves it. Returns the path of the file.
        /// </summary>
        /// <param name="url">URL, absolute or relative to the current page</param>
        /// <param name="path">Optional - File path, or folder to keep the server's file name. Default is the Downloads folder.</param>
        /// <param name="timeout">Optional - Timeout in milliseconds. Default is 60 seconds.</param>
        public string DownloadResource(string url, string path = null, int timeout = 60000) {
            return PageHelper.DownloadResource(session, url, path, timeout);
        }

        #endregion


        #region Waits

        /// <summary>
        /// Waits until the page is loaded and no network request (fetch/XHR) happened for idleTime milliseconds.
        /// </summary>
        public void WaitForNetworkIdle(int idleTime = 500, int timeout = 30000) {
            PageHelper.WaitForNetworkIdle(session, idleTime, timeout);
        }

        /// <summary>
        /// Waits until the page is loaded and the DOM did not change for idleTime milliseconds.
        /// </summary>
        public void WaitForDomIdle(int idleTime = 500, int timeout = 30000) {
            PageHelper.WaitForDomIdle(session, null, idleTime, timeout);
        }

        /// <summary>
        /// Waits until the page is loaded and jQuery has no active request. Returns at once without jQuery.
        /// </summary>
        public void WaitForJQueryIdle(int timeout = 30000) {
            PageHelper.WaitForJQueryIdle(session, timeout);
        }

        #endregion


        #region Scrolling

        /// <summary>
        /// Scrolls the page to the top.
        /// </summary>
        public void ScrollToTop() {
            PageHelper.Scroll(session, null, "top", 0, 0);
        }

        /// <summary>
        /// Scrolls the page to the bottom.
        /// </summary>
        public void ScrollToBottom() {
            PageHelper.Scroll(session, null, "bottom", 0, 0);
        }

        /// <summary>
        /// Scrolls the page to a position in pixels.
        /// </summary>
        public void ScrollTo(int x, int y) {
            PageHelper.Scroll(session, null, "to", x, y);
        }

        /// <summary>
        /// Scrolls the page by an offset in pixels.
        /// </summary>
        public void ScrollBy(int x, int y) {
            PageHelper.Scroll(session, null, "by", x, y);
        }

        /// <summary>
        /// Scrolls down until no more content loads (infinite scroll). Returns the number of scrolls that loaded content.
        /// </summary>
        /// <param name="maxScrolls">Optional - Maximum number of scrolls. Default is 100.</param>
        /// <param name="pause">Optional - Time to wait for the content after each scroll, in milliseconds. Default is 1000.</param>
        public int ScrollUntilEnd(int maxScrolls = 100, int pause = 1000) {
            return PageHelper.ScrollUntilEnd(session, null, maxScrolls, pause);
        }

        /// <summary>
        /// Height of the page content in pixels.
        /// </summary>
        public int ScrollHeight {
            get {
                return PageHelper.ScrollSize(session, null, "scrollHeight");
            }
        }

        /// <summary>
        /// Width of the page content in pixels.
        /// </summary>
        public int ScrollWidth {
            get {
                return PageHelper.ScrollSize(session, null, "scrollWidth");
            }
        }

        #endregion


        #region Windows and tabs

        /// <summary>
        /// Opens a new tab, switches to it and optionally loads a URL.
        /// </summary>
        public Window NewTab(string url = null) {
            return OpenNewWindow("tab", url);
        }

        /// <summary>
        /// Opens a new window, switches to it and optionally loads a URL.
        /// </summary>
        public Window NewWindow(string url = null) {
            return OpenNewWindow("window", url);
        }

        private Window OpenNewWindow(string type, string url) {
            var result = (Dictionary)session.Send(RequestMethod.POST, "/window/new", "type", type);
            Window window = session.windows.SwitchToHandle((string)result["handle"]);
            if (!string.IsNullOrEmpty(url))
                this.Get(url);
            return window;
        }

        /// <summary>
        /// Switches to a window opened since Start or the last window switch (ex: by a link with target="_blank").
        /// </summary>
        public Window SwitchToNewWindow(int timeout = -1, bool raise = true) {
            try {
                return session.windows.SwitchToNewWindow(timeout);
            } catch (Errors.NoSuchWindowError) {
                if (raise)
                    throw;
                return null;
            }
        }

        /// <summary>
        /// Switches to the window whose URL matches the pattern (wildcards * and ?).
        /// </summary>
        public Window SwitchToWindowByUrl(string urlPattern, int timeout = -1, bool raise = true) {
            try {
                return session.windows.SwitchToWindowByUrl(urlPattern, timeout);
            } catch (Errors.NoSuchWindowError) {
                if (raise)
                    throw;
                return null;
            }
        }

        #endregion

    }

}
