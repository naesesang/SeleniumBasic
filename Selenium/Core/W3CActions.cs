namespace Selenium.Core {

    /// <summary>
    /// Sends input actions with the W3C "Perform Actions" command.
    /// </summary>
    /// <remarks>
    /// The remote end keeps the state of an input source (pointer position, pressed keys and buttons)
    /// between two commands, so a sequence can be sent in several calls.
    /// </remarks>
    static class W3CActions {

        const string POINTER_ID = "mouse";
        const string KEYBOARD_ID = "keyboard";

        public static Dictionary PointerMove(object origin, int x, int y) {
            var action = new Dictionary();
            action.Add("type", "pointerMove");
            action.Add("duration", 0);
            action.Add("origin", origin);
            action.Add("x", x);
            action.Add("y", y);
            return action;
        }

        public static Dictionary PointerDown(MouseButton button) {
            return Action("pointerDown", "button", (int)button);
        }

        public static Dictionary PointerUp(MouseButton button) {
            return Action("pointerUp", "button", (int)button);
        }

        public static Dictionary KeyDown(string key) {
            return Action("keyDown", "value", key);
        }

        public static Dictionary KeyUp(string key) {
            return Action("keyUp", "value", key);
        }

        public static void PerformPointer(RemoteSession session, params Dictionary[] actions) {
            var parameters = new Dictionary();
            parameters.Add("pointerType", "mouse");
            Perform(session, "pointer", POINTER_ID, parameters, actions);
        }

        /// <summary>
        /// Scrolls with the wheel. The origin is "viewport" or a serialized element.
        /// </summary>
        public static void PerformWheel(RemoteSession session, object origin, int x, int y, int deltaX, int deltaY) {
            var action = new Dictionary();
            action.Add("type", "scroll");
            action.Add("duration", 0);
            action.Add("origin", origin);
            action.Add("x", x);
            action.Add("y", y);
            action.Add("deltaX", deltaX);
            action.Add("deltaY", deltaY);
            Perform(session, "wheel", "wheel", null, new object[] { action });
        }

        public static void PerformKeys(RemoteSession session, List actions) {
            if (actions.Count == 0)
                return;
            Perform(session, "key", KEYBOARD_ID, null, actions.Values());
        }

        private static Dictionary Action(string type, string key, object value) {
            var action = new Dictionary();
            action.Add("type", type);
            action.Add(key, value);
            return action;
        }

        private static void Perform(RemoteSession session, string type, string id, Dictionary parameters, object[] actions) {
            var source = new Dictionary();
            source.Add("type", type);
            source.Add("id", id);
            if (parameters != null)
                source.Add("parameters", parameters);
            source.Add("actions", actions);
            session.Send(RequestMethod.POST, "/actions", "actions", new object[] { source });
        }

    }

}
