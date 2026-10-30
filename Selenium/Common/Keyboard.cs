using Selenium.Core;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Selenium {

    /// <summary>
    /// Keyboard object
    /// </summary>
    [ProgId("Selenium.Keyboard")]
    [Guid("0277FC34-FD1B-4616-BB19-988CE6BB6E1F")]
    [Description("Keyboard interface.")]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class Keyboard : ComInterfaces._Keyboard {

        const char KEY_NULL = '\xE000';
        const char KEY_SHIFT = '\xE008';
        const char KEY_ALT = '\xE009';
        const char KEY_CTRL = '\xE00A';
        const char KEY_META = '\xE03D';

        private RemoteSession _session;
        private Keys _keys;
        private readonly System.Collections.Generic.List<char> _pressedModifiers;

        internal Keyboard(RemoteSession session) {
            _session = session;
            _keys = new Keys();
            _pressedModifiers = new System.Collections.Generic.List<char>(4);
        }

        /// <summary>
        /// Returns a list of pressable keys.
        /// </summary>
        public Keys Keys {
            get {
                return _keys;
            }
        }

        /// <summary>
        /// Sends a sequence of key strokes to the active element.
        /// </summary>
        /// <param name="keysOrModifiers">Modifier keys (Ctrl, Shit or Alt) or keys</param>
        /// <param name="keys">Keys</param>
        /// <returns>Self</returns>
        public Keyboard SendKeys(string keysOrModifiers, string keys = null) {
            if (keys != null)
                keysOrModifiers = string.Concat(keysOrModifiers, keys, keysOrModifiers);

            // Like the legacy /keys command: a modifier key toggles its state
            // and the NULL key releases all the pressed modifiers.
            var actions = new List();
            for (int i = 0; i < keysOrModifiers.Length; i++) {
                char c = keysOrModifiers[i];
                if (c == KEY_NULL) {
                    foreach (char modifier in _pressedModifiers)
                        actions.Add(W3CActions.KeyUp(modifier.ToString()));
                    _pressedModifiers.Clear();
                } else if (IsModifier(c)) {
                    if (_pressedModifiers.Remove(c)) {
                        actions.Add(W3CActions.KeyUp(c.ToString()));
                    } else {
                        _pressedModifiers.Add(c);
                        actions.Add(W3CActions.KeyDown(c.ToString()));
                    }
                } else {
                    string key = c.ToString();
                    if (char.IsHighSurrogate(c) && i + 1 < keysOrModifiers.Length)
                        key += keysOrModifiers[++i];
                    actions.Add(W3CActions.KeyDown(key));
                    actions.Add(W3CActions.KeyUp(key));
                }
            }
            W3CActions.PerformKeys(_session, actions);
            return this;
        }

        /// <summary>
        /// Presses and holds modifier keys (Control, Alt and Shift).
        /// </summary>
        /// <param name="modifierKeys">The modifier key to Send. Values are defined in Keys class.</param>
        /// <returns>Self</returns>
        public Keyboard KeyDown(string modifierKeys) {
            check_keys_are_modifiers(modifierKeys);
            var actions = new List();
            foreach (char key in modifierKeys) {
                if (!_pressedModifiers.Contains(key))
                    _pressedModifiers.Add(key);
                actions.Add(W3CActions.KeyDown(key.ToString()));
            }
            W3CActions.PerformKeys(_session, actions);
            return this;
        }

        /// <summary>
        /// Release modifier keys (Control, Alt and Shift).
        /// </summary>
        /// <param name="modifierKeys">The modifier key to Send. Values are defined in Keys class.</param>
        /// <returns>Self</returns>
        public Keyboard KeyUp(string modifierKeys) {
            check_keys_are_modifiers(modifierKeys);
            var actions = new List();
            foreach (char key in modifierKeys) {
                _pressedModifiers.Remove(key);
                actions.Add(W3CActions.KeyUp(key.ToString()));
            }
            W3CActions.PerformKeys(_session, actions);
            return this;
        }


        #region private support methods

        private static bool IsModifier(char key) {
            return key == KEY_SHIFT || key == KEY_ALT || key == KEY_CTRL || key == KEY_META;
        }

        private void check_keys_are_modifiers(string keys) {
            foreach (char key in keys) {
                if (key != KEY_SHIFT && key != KEY_ALT && key != KEY_CTRL)
                    throw new Errors.InvalideModifierKeyError();
            }
        }

        #endregion

    }

}
