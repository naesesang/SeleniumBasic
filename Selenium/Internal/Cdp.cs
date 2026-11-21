using Selenium.Core;
using Selenium.Serializer;

namespace Selenium.Internal {

    /// <summary>
    /// Sends Chrome DevTools Protocol commands through chromedriver or msedgedriver.
    /// </summary>
    static class Cdp {

        /// <summary>
        /// Returns true when the session runs Chrome or Edge.
        /// </summary>
        public static bool IsChromium(RemoteSession session) {
            string name = BrowserName(session);
            return name == "chrome" || name == "msedge" || name == "MicrosoftEdge";
        }

        /// <summary>
        /// Executes a CDP command and returns its result.
        /// </summary>
        /// <param name="session">Session</param>
        /// <param name="command">Command (ex: "Network.clearBrowserCache")</param>
        /// <param name="parameters">Null, a Dictionary, or a JSON string</param>
        public static object Execute(RemoteSession session, string command, object parameters) {
            if (!IsChromium(session))
                throw new Errors.InvalidOperationError("Chrome DevTools commands require Chrome or Edge.");

            var data = new Dictionary();
            data.Add("cmd", command);
            data.Add("params", ToDictionary(parameters));

            string vendor = BrowserName(session) == "chrome" ? "goog" : "ms";
            return session.Send(RequestMethod.POST, "/" + vendor + "/cdp/execute", data);
        }

        internal static Dictionary ToDictionary(object parameters) {
            if (parameters == null)
                return new Dictionary();
            var dict = parameters as Dictionary;
            if (dict != null)
                return dict;
            var text = parameters as string;
            if (text != null) {
                if (text.Trim().Length == 0)
                    return new Dictionary();
                dict = JSON.Parse(text) as Dictionary;
                if (dict != null)
                    return dict;
            }
            throw new Errors.ArgumentError("The parameters must be a Dictionary or a JSON object string.");
        }

        private static string BrowserName(RemoteSession session) {
            var capabilities = session.capabilities;
            if (capabilities == null)
                return null;
            return capabilities.Get("browserName", null) as string;
        }

    }

}
