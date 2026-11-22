using Selenium.Core;
using System;
using System.IO;
using System.Threading;

namespace Selenium.Internal {

    /// <summary>
    /// Scrolling, waiting and resource download through JavaScript, shared by WebDriver and WebElement.
    /// </summary>
    static class PageHelper {

        // arguments[0]: the element to scroll, or null for the page
        const string JS_TARGET = "var t=arguments[0]||document.scrollingElement||document.documentElement;";

        public static void Scroll(RemoteSession session, WebElement element, string operation, int x, int y) {
            string script;
            switch (operation) {
                case "top": script = "t.scrollTop=0;"; break;
                case "bottom": script = "t.scrollTop=t.scrollHeight;"; break;
                case "to": script = "t.scrollTo(arguments[1],arguments[2]);"; break;
                case "by": script = "t.scrollBy(arguments[1],arguments[2]);"; break;
                default: throw new ArgumentException(operation);
            }
            session.javascript.Execute(JS_TARGET + script, new object[] { Arg(element), x, y }, false);
        }

        public static int ScrollSize(RemoteSession session, WebElement element, string property) {
            object value = session.javascript.Execute(JS_TARGET + "return t." + property + ";"
                , new object[] { Arg(element) }, false);
            return Convert.ToInt32(value);
        }

        /// <summary>
        /// Scrolls to the bottom until the content stops growing. Returns the number of scrolls that loaded more content.
        /// </summary>
        public static int ScrollUntilEnd(RemoteSession session, WebElement element, int maxScrolls, int pause) {
            const string script = JS_TARGET + "t.scrollTop=t.scrollHeight;return t.scrollHeight;";
            var args = new object[] { Arg(element) };
            int count = 0;
            int height = Convert.ToInt32(session.javascript.Execute(script, args, false));
            while (count < maxScrolls) {
                // Content can arrive late: check twice before concluding the end is reached
                int newHeight = height;
                for (int attempt = 0; attempt < 2 && newHeight == height; attempt++) {
                    SysWaiter.Wait(pause);    // Ctrl+Pause interrupts the wait
                    newHeight = Convert.ToInt32(session.javascript.Execute(script, args, false));
                }
                if (newHeight <= height)
                    break;
                height = newHeight;
                count++;
            }
            return count;
        }

        /// <summary>
        /// Waits until no DOM mutation happened for the given idle time.
        /// </summary>
        public static void WaitForDomIdle(RemoteSession session, WebElement element, int idleTime, int timeout) {
            const string install =
                "var t=arguments[0]||document.documentElement;"
                + "if(!t.__sbObserver){"
                + " t.__sbLast=Date.now();"
                + " t.__sbObserver=new MutationObserver(function(){t.__sbLast=Date.now();});"
                + " t.__sbObserver.observe(t,{childList:true,subtree:true,attributes:true,characterData:true});"
                + "}"
                + "return Date.now()-t.__sbLast;";
            var args = new object[] { Arg(element) };
            WaitUntil(timeout, "DOM idle", () =>
                Convert.ToInt64(session.javascript.Execute(install, args, false)) >= idleTime
                && IsDocumentComplete(session));
        }

        /// <summary>
        /// Waits until the page is loaded and no fetch/XHR request was pending or completed for the given idle time.
        /// </summary>
        public static void WaitForNetworkIdle(RemoteSession session, int idleTime, int timeout) {
            const string script =
                "var w=window;"
                + "if(!w.__sbNet){"
                + " var n=w.__sbNet={pending:0,last:Date.now()};"
                + " var done=function(){n.pending--;n.last=Date.now();};"
                + " var send=XMLHttpRequest.prototype.send;"
                + " XMLHttpRequest.prototype.send=function(){n.pending++;n.last=Date.now();"
                + "  this.addEventListener('loadend',done);return send.apply(this,arguments);};"
                + " if(w.fetch){var f=w.fetch;w.fetch=function(){n.pending++;n.last=Date.now();"
                + "  return f.apply(this,arguments).then(function(r){done();return r;},function(e){done();throw e;});};}"
                + "}"
                + "var r=performance.getEntriesByType('resource');"
                + "var count=r.length, end=0;"
                + "for(var i=0;i<r.length;i++){if(r[i].responseEnd>end)end=r[i].responseEnd;}"
                + "var idle=Math.min(Date.now()-w.__sbNet.last, performance.now()-end);"
                + "return [document.readyState, w.__sbNet.pending, idle];";
            WaitUntil(timeout, "network idle", () => {
                var state = (List)session.javascript.Execute(script, null, false);
                return (string)state[0] == "complete"
                    && Convert.ToInt32(state[1]) <= 0
                    && Convert.ToDouble(state[2]) >= idleTime;
            });
        }

        public static void WaitForJQueryIdle(RemoteSession session, int timeout) {
            WaitUntil(timeout, "jQuery idle", () =>
                IsDocumentComplete(session)
                && (bool)session.javascript.Execute("return !window.jQuery || window.jQuery.active === 0;", null, false));
        }

        private static bool IsDocumentComplete(RemoteSession session) {
            return "complete" == session.javascript.Execute("return document.readyState;", null, false) as string;
        }

        private static void WaitUntil(int timeout, string what, Func<bool> condition) {
            DateTime endTime = DateTime.UtcNow.AddMilliseconds(timeout);
            while (!condition()) {
                if (DateTime.UtcNow > endTime)
                    throw new Errors.TimeoutError("Timed out after {0} ms waiting for {1}.", timeout, what);
                SysWaiter.Wait(100);    // Ctrl+Pause interrupts the wait
            }
        }

        /// <summary>
        /// Downloads a URL with the browser's cookies and session, and saves it.
        /// </summary>
        /// <param name="session">Session</param>
        /// <param name="url">URL, relative to the current page</param>
        /// <param name="path">File path, or folder (existing or ending with \) to keep the server's file name</param>
        /// <param name="timeout">Timeout in milliseconds</param>
        /// <returns>Full path of the saved file</returns>
        public static string DownloadResource(RemoteSession session, string url, string path, int timeout) {
            const string script =
                "var url=new URL(arguments[0],document.baseURI).href;"
                + "fetch(url,{credentials:'include'}).then(function(r){"
                + " if(!r.ok){callback({error:'HTTP '+r.status+' '+url});return;}"
                + " var cd=r.headers.get('content-disposition')||'';"
                + " return r.blob().then(function(b){var fr=new FileReader();"
                // The data URL prefix is kept: the JSON parser would turn a bare PNG Base64 string into an Image
                + "  fr.onload=function(){callback({url:r.url,disposition:cd,data:fr.result});};"
                + "  fr.onerror=function(){callback({error:'Failed to read '+url});};"
                + "  fr.readAsDataURL(b);});"
                + "}).catch(function(e){callback({error:String(e)+' '+url});});";

            int previous = session.timeouts.Script;
            Dictionary result;
            try {
                session.timeouts.Script = timeout;
                result = (Dictionary)session.javascript.ExecuteAsync(script, new object[] { url }, false);
            } finally {
                session.timeouts.Script = previous;
            }

            object error;
            if (result.TryGetValue("error", out error))
                throw new SeleniumError("Failed to download the resource: {0}", error);

            string finalUrl = (string)result["url"];
            string target = ResolveTarget(path, finalUrl, result.Get("disposition", "") as string);
            string dir = Path.GetDirectoryName(target);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);
            string dataUrl = (string)result["data"];
            File.WriteAllBytes(target, Convert.FromBase64String(dataUrl.Substring(dataUrl.IndexOf(',') + 1)));
            return target;
        }

        private static string ResolveTarget(string path, string url, string disposition) {
            if (string.IsNullOrEmpty(path))
                path = Downloads.GetDefaultFolder() + "\\";
            path = IOExt.ExpandPath(path.EndsWith("\\") || path.EndsWith("/") ? path + "." : path);
            bool isFolder = Directory.Exists(path) || path.EndsWith("\\.") || path.EndsWith("/.");
            if (!isFolder)
                return Path.GetFullPath(path);
            string folder = Path.GetFullPath(path.TrimEnd('.'));
            return Path.Combine(folder, GetFileName(url, disposition));
        }

        private static string GetFileName(string url, string disposition) {
            string name = null;
            const System.Text.RegularExpressions.RegexOptions IGNORE_CASE = System.Text.RegularExpressions.RegexOptions.IgnoreCase;
            // The encoded name (filename*=UTF-8''...) takes precedence over the plain one
            var match = System.Text.RegularExpressions.Regex.Match(disposition ?? "", @"filename\*=[^']*'[^']*'([^;]+)", IGNORE_CASE);
            if (match.Success) {
                name = Uri.UnescapeDataString(match.Groups[1].Value.Trim().Trim('"'));
            } else {
                match = System.Text.RegularExpressions.Regex.Match(disposition ?? "", @"filename=""?([^"";]+)""?", IGNORE_CASE);
                if (match.Success)
                    name = match.Groups[1].Value.Trim();
            }
            if (string.IsNullOrEmpty(name)) {
                try {
                    name = Uri.UnescapeDataString(Path.GetFileName(new Uri(url).AbsolutePath));
                } catch (UriFormatException) { }
            }
            if (string.IsNullOrEmpty(name))
                name = "download";
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }

        private static object Arg(WebElement element) {
            return element == null ? null : (object)element.SerializeJson();
        }

    }

}
