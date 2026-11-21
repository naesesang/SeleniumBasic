using Selenium.Core;
using Selenium.Serializer;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Selenium {

    /// <summary>
    /// Shadow root of an element (Shadow DOM). Finds the elements inside it.
    /// </summary>
    /// <example>
    /// <code lang="vbs">
    /// Set root = driver.FindElementByCss("my-component").ShadowRoot
    /// root.FindElementByCss("button").Click
    /// </code>
    /// </example>
    [ProgId("Selenium.ShadowRoot")]
    [Guid("0277FC34-FD1B-4616-BB19-3B8E5F0D72A1")]
    [Description("Shadow root of an element (Shadow DOM)")]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class ShadowRoot : SearchContext, ComInterfaces._ShadowRoot, IJsonObject {

        /// <summary>
        /// Key identifying a shadow root reference in the W3C protocol.
        /// </summary>
        internal const string SHADOW_KEY = "shadow-6066-11e4-a52e-4f735466cecf";

        internal static bool TryParse(Dictionary dict, out string id) {
            return dict.TryGetValue(SHADOW_KEY, out id);
        }

        private readonly RemoteSession _session;
        internal readonly string Id;

        internal ShadowRoot(RemoteSession session, string id) {
            _session = session;
            this.Id = id;
        }

        internal override RemoteSession session {
            get {
                return _session;
            }
        }

        internal override string uri {
            get {
                return "/shadow/" + this.Id;
            }
        }

        /// <summary>
        /// </summary>
        public Dictionary SerializeJson() {
            var dict = new Dictionary();
            dict.Add(SHADOW_KEY, this.Id);
            return dict;
        }

    }

}
