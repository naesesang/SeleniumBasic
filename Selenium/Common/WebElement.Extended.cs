using Selenium.Core;
using Selenium.Internal;
using System;
using System.Collections;
using System.IO;
using System.Text;

namespace Selenium {

    // Features added in 3.2: navigation in the tree, attributes, focus, accessibility,
    // highlight, upload, resource download and scrolling.
    public partial class WebElement {

        private object Script(string script, params object[] args) {
            var arguments = new object[args.Length + 1];
            arguments[0] = this.SerializeJson();
            Array.Copy(args, 0, arguments, 1, args.Length);
            return _session.javascript.Execute(script, arguments, true);
        }

        #region Tree

        /// <summary>
        /// Parent element, or Nothing for the root element.
        /// </summary>
        public WebElement Parent {
            get {
                return Script("return arguments[0].parentElement;") as WebElement;
            }
        }

        /// <summary>
        /// Child elements.
        /// </summary>
        public WebElements Children() {
            var result = Script("return Array.prototype.slice.call(arguments[0].children);");
            return result as WebElements ?? new WebElements();
        }

        #endregion


        #region Attributes and properties

        /// <summary>
        /// Sets an HTML attribute of the element.
        /// </summary>
        public void SetAttribute(string name, string value) {
            Script("arguments[0].setAttribute(arguments[1],arguments[2]);", name, value ?? "");
        }

        /// <summary>
        /// Removes an HTML attribute of the element.
        /// </summary>
        public void RemoveAttribute(string name) {
            Script("arguments[0].removeAttribute(arguments[1]);", name);
        }

        /// <summary>
        /// Returns True when the element has the HTML attribute.
        /// </summary>
        public bool HasAttribute(string name) {
            return (bool)Script("return arguments[0].hasAttribute(arguments[1]);", name);
        }

        /// <summary>
        /// Returns a JavaScript property of the element (ex: "value", "checked"), without falling back to the attribute.
        /// </summary>
        public object Property(string name) {
            return Send(RequestMethod.GET, "/property/" + name);
        }

        /// <summary>
        /// Returns an HTML attribute as written in the page, without falling back to the property. Nothing when missing.
        /// </summary>
        public object DomAttribute(string name) {
            return Send(RequestMethod.GET, "/attribute/" + name);
        }

        #endregion


        #region Focus and accessibility

        /// <summary>
        /// Gives the focus to the element.
        /// </summary>
        public void Focus() {
            Script("arguments[0].focus();");
        }

        /// <summary>
        /// True when the element has the focus.
        /// </summary>
        public bool HasFocus {
            get {
                return (bool)Script(
                    "var a=document.activeElement;"
                    + "while(a&&a.shadowRoot&&a.shadowRoot.activeElement)a=a.shadowRoot.activeElement;"
                    + "return a===arguments[0];");
            }
        }

        /// <summary>
        /// Accessible name of the element, as read by a screen reader.
        /// </summary>
        public string AriaLabel {
            get {
                return Send(RequestMethod.GET, "/computedlabel") as string ?? "";
            }
        }

        /// <summary>
        /// Accessible role of the element (ex: button, link, textbox).
        /// </summary>
        public string AriaRole {
            get {
                return Send(RequestMethod.GET, "/computedrole") as string ?? "";
            }
        }

        #endregion


        #region Highlight

        /// <summary>
        /// Highlights the element with a colored outline.
        /// </summary>
        /// <param name="color">Optional - CSS color. Default is red.</param>
        public void Highlight(string color = "red") {
            Script("var e=arguments[0];"
                + "if(e.__sbOutline===undefined){e.__sbOutline=e.style.outline;e.__sbOffset=e.style.outlineOffset;}"
                + "e.style.outline='3px solid '+arguments[1];e.style.outlineOffset='-1px';"
                , string.IsNullOrEmpty(color) ? "red" : color);
        }

        /// <summary>
        /// Removes the highlight of the element.
        /// </summary>
        public void Unhighlight() {
            Script("var e=arguments[0];"
                + "if(e.__sbOutline!==undefined){e.style.outline=e.__sbOutline;e.style.outlineOffset=e.__sbOffset;"
                + "delete e.__sbOutline;delete e.__sbOffset;}");
        }

        #endregion


        #region Files

        /// <summary>
        /// Selects files in a file input (input type="file"). Accepts several paths or an array.
        /// </summary>
        /// <example>
        /// <code lang="vbs">
        /// driver.FindElementByCss("input[type=file]").UploadFile "C:\a.pdf", "C:\b.pdf"
        /// </code>
        /// </example>
        public void UploadFile(params object[] filePaths) {
            var paths = new StringBuilder();
            foreach (string path in Flatten(filePaths)) {
                string fullPath = Path.GetFullPath(IOExt.ExpandPath(path));
                if (!File.Exists(fullPath))
                    throw new Errors.FileNotFoundError(fullPath);
                if (paths.Length != 0)
                    paths.Append('\n');
                paths.Append(fullPath);
            }
            if (paths.Length == 0)
                throw new Errors.ArgumentError("No file to upload.");
            Send(RequestMethod.POST, "/value", "text", paths.ToString());
        }

        private static System.Collections.Generic.IEnumerable<string> Flatten(IEnumerable items) {
            foreach (object item in items) {
                if (item == null)
                    continue;
                var text = item as string;
                if (text != null) {
                    if (text.Length != 0)
                        yield return text;
                } else if (item is IEnumerable) {
                    foreach (string nested in Flatten((IEnumerable)item))
                        yield return nested;
                } else {
                    yield return item.ToString();
                }
            }
        }

        /// <summary>
        /// Downloads the resource of the element (image src, link href...) with the browser's cookies and saves it.
        /// </summary>
        /// <param name="path">Optional - File path, or folder to keep the server's file name. Default is the Downloads folder.</param>
        /// <param name="timeout">Optional - Timeout in milliseconds. Default is 60 seconds.</param>
        /// <returns>Full path of the file</returns>
        public string DownloadResource(string path = null, int timeout = 60000) {
            string url = Script("var e=arguments[0];"
                + "return e.currentSrc||e.src||e.href||e.data||e.getAttribute('src')||e.getAttribute('href')||'';") as string;
            if (string.IsNullOrEmpty(url))
                throw new SeleniumError("The element has no src or href to download.");
            return PageHelper.DownloadResource(_session, url, path, timeout);
        }

        #endregion


        #region Scrolling and waiting

        /// <summary>
        /// Scrolls the content of the element to the top.
        /// </summary>
        public void ScrollToTop() {
            PageHelper.Scroll(_session, this, "top", 0, 0);
        }

        /// <summary>
        /// Scrolls the content of the element to the bottom.
        /// </summary>
        public void ScrollToBottom() {
            PageHelper.Scroll(_session, this, "bottom", 0, 0);
        }

        /// <summary>
        /// Scrolls the content of the element to a position in pixels.
        /// </summary>
        public void ScrollTo(int x, int y) {
            PageHelper.Scroll(_session, this, "to", x, y);
        }

        /// <summary>
        /// Scrolls the content of the element by an offset in pixels.
        /// </summary>
        public void ScrollBy(int x, int y) {
            PageHelper.Scroll(_session, this, "by", x, y);
        }

        /// <summary>
        /// Scrolls the content of the element down until no more content loads. Returns the number of scrolls that loaded content.
        /// </summary>
        public int ScrollUntilEnd(int maxScrolls = 100, int pause = 1000) {
            return PageHelper.ScrollUntilEnd(_session, this, maxScrolls, pause);
        }

        /// <summary>
        /// Height of the content of the element in pixels.
        /// </summary>
        public int ScrollHeight {
            get {
                return PageHelper.ScrollSize(_session, this, "scrollHeight");
            }
        }

        /// <summary>
        /// Width of the content of the element in pixels.
        /// </summary>
        public int ScrollWidth {
            get {
                return PageHelper.ScrollSize(_session, this, "scrollWidth");
            }
        }

        /// <summary>
        /// Waits until the content of the element did not change for idleTime milliseconds.
        /// </summary>
        public void WaitForDomIdle(int idleTime = 500, int timeout = 30000) {
            PageHelper.WaitForDomIdle(_session, this, idleTime, timeout);
        }

        #endregion

    }

}
