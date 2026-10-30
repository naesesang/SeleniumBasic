using System;

namespace Selenium.Core {

    /// <summary>
    /// Converts the legacy (JSON Wire) capabilities to W3C capabilities.
    /// </summary>
    /// <remarks>
    /// A W3C remote end rejects any top level capability that is neither a standard
    /// capability nor a vendor extension ("prefix:name"), so the other legacy keys are left out.
    /// </remarks>
    static class W3CCapabilities {

        public static Dictionary FromLegacy(Dictionary legacy) {
            var w3c = new Dictionary();
            string browserName = legacy.Get("browserName", null) as string;

            foreach (DictionaryItem item in legacy) {
                string key = item.Key;
                object value = item.Value;
                if (value == null)
                    continue;

                switch (key) {
                    case "browserName":
                    case "pageLoadStrategy":
                    case "acceptInsecureCerts":
                    case "setWindowRect":
                    case "strictFileInteractability":
                    case "timeouts":
                    case "unhandledPromptBehavior":
                        w3c.Set(key, value);
                        break;
                    case "browserVersion":
                    case "version":
                        if (!string.IsNullOrEmpty(value as string))
                            w3c.Set("browserVersion", value);
                        break;
                    case "platform":
                    case "platformName":
                        string platform = FormatPlatform(value as string);
                        if (platform != null)
                            w3c.Set("platformName", platform);
                        break;
                    case "acceptSslCerts":
                        w3c.Set("acceptInsecureCerts", value);
                        break;
                    case "unexpectedAlertBehaviour":
                        w3c.Set("unhandledPromptBehavior", value);
                        break;
                    case "proxy":
                        var proxy = value as Dictionary;
                        if (proxy != null)
                            w3c.Set(key, FormatProxy(proxy));
                        break;
                    case "chromeOptions":
                        w3c.Set(browserName == "MicrosoftEdge" ? "ms:edgeOptions" : "goog:chromeOptions", value);
                        break;
                    case "edgeOptions":
                        w3c.Set("ms:edgeOptions", value);
                        break;
                    case "loggingPrefs":
                        if (browserName == "chrome")
                            w3c.Set("goog:loggingPrefs", value);
                        break;
                    default:
                        if (key.IndexOf(':') > 0)
                            w3c.Set(key, value);
                        break;
                }
            }
            return w3c;
        }

        /// <summary>
        /// Returns the W3C platform name or null for any platform.
        /// </summary>
        private static string FormatPlatform(string platform) {
            if (string.IsNullOrEmpty(platform))
                return null;
            string name = platform.ToLowerInvariant();
            if (name == "any")
                return null;
            if (name.StartsWith("win") || name == "xp" || name == "vista")
                return "windows";
            if (name.StartsWith("mac") || name.StartsWith("os x"))
                return "mac";
            return name;
        }

        /// <summary>
        /// Returns a copy of the proxy without the keys unknown to the W3C protocol.
        /// </summary>
        private static Dictionary FormatProxy(Dictionary proxy) {
            var w3c = new Dictionary();
            foreach (DictionaryItem item in proxy) {
                switch (item.Key) {
                    case "autodetect":
                    case "socksUsername":
                    case "socksPassword":
                        break;
                    case "proxyType":
                        w3c.Set(item.Key, Convert.ToString(item.Value).ToLowerInvariant());
                        break;
                    default:
                        w3c.Set(item.Key, item.Value);
                        break;
                }
            }
            if (w3c.ContainsKey("socksProxy") && !w3c.ContainsKey("socksVersion"))
                w3c.Set("socksVersion", 5);
            return w3c;
        }

    }

}
