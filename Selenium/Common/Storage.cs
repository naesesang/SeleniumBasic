using Selenium.Core;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Selenium {

    /// <summary>
    /// Storage object
    /// </summary>
    [ProgId("Selenium.Storage")]
    [Guid("0277FC34-FD1B-4616-BB19-4425731F2305")]
    [Description("Storage object")]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class Storage : ComInterfaces._Storage {

        internal enum StorageType {
            SessionStorage,
            LocalStorage
        }

        private readonly RemoteSession _session;
        private readonly string _storage;

        /// <summary>
        /// </summary>
        /// <param name="session"></param>
        /// <param name="storageType"></param>
        internal Storage(RemoteSession session, StorageType storageType) {
            _session = session;
            // No storage command in the W3C protocol: the storage is accessed from the page
            _storage = storageType == StorageType.LocalStorage ? "localStorage" : "sessionStorage";
        }

        /// <summary>
        /// Get or set the storage item for the given key.
        /// </summary>
        /// <param name="key">The key to set.</param>
        /// <param name="value">The value to set.</param>
        /// <returns>object</returns>
        public string this[string key] {
            get {
                return (string)Execute("return {0}.getItem(arguments[0]);", key);
            }
            set {
                this.Set(key, value);
            }
        }

        /// <summary>
        /// Sets the storage item
        /// </summary>
        /// <param name="key">Key</param>
        /// <param name="value">Value</param>
        public void Set(string key, string value) {
            Execute("{0}.setItem(arguments[0], arguments[1]);", key, value);
        }

        /// <summary>
        /// Get all keys of the storage.
        /// </summary>
        /// <returns>The list of keys.</returns>
        public List Keys() {
            return (List)Execute("return Object.keys({0});");
        }

        /// <summary>
        /// Get the number of items in the storage.
        /// </summary>
        /// <returns>{number} The number of items in the storage.</returns>
        public int Count() {
            return Convert.ToInt32(Execute("return {0}.length;"));
        }

        /// <summary>
        /// Clear the storage.
        /// </summary>
        public void Clear() {
            Execute("{0}.clear();");
        }

        /// <summary>
        /// Remove the storage item for the given key.
        /// </summary>
        /// <param name="key">Key of the item</param>
        public void Remove(string key) {
            Execute("{0}.removeItem(arguments[0]);", key);
        }

        private object Execute(string script, params object[] arguments) {
            return _session.javascript.Execute(string.Format(script, _storage), arguments, false);
        }

    }

}
