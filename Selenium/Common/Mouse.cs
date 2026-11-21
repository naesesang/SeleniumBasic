using Selenium.Core;
using Selenium;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Selenium {

    /// <summary>
    /// Mouse interraction.
    /// </summary>
    [ProgId("Selenium.Mouse")]
    [Guid("0277FC34-FD1B-4616-BB19-94EBFCC1BAFC")]
    [Description("Mouse interface.")]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public class Mouse : ComInterfaces._Mouse {

        private RemoteSession _session;

        internal Mouse(RemoteSession session) {
            _session = session;
        }

        /// <summary>
        /// Move the mouse to the specificed element.
        /// </summary>
        /// <param name="element">Opaque ID assigned to the element to move to, as described in the WebElement JSON Object. If not specified or is null, the offset is relative to current position of the mouse.</param>
        public Mouse moveTo(WebElement element) {
            ScrollIntoViewIfNeeded(element);
            W3CActions.PerformPointer(_session, W3CActions.PointerMove(element.SerializeJson(), 0, 0));
            return this;
        }

        /// <summary>
        /// Moves the mouse to a position in the viewport.
        /// </summary>
        internal void MoveToViewport(int x, int y) {
            W3CActions.PerformPointer(_session, W3CActions.PointerMove("viewport", x, y));
        }

        /// <summary>
        /// Move the mouse by an offset of the specificed element.
        /// </summary>
        /// <param name="element">Opaque ID assigned to the element to move to, as described in the WebElement JSON Object. If not specified or is null, the offset is relative to current position of the mouse.</param>
        /// <param name="xoffset">{number} X offset to move to, relative to the top-left corner of the element. If not specified, the mouse will move to the middle of the element.</param>
        /// <param name="yoffset"> {number} Y offset to move to, relative to the top-left corner of the element. If not specified, the mouse will move to the middle of the element.</param>
        public Mouse MoveTo(WebElement element, int xoffset = 0, int yoffset = 0) {
            Dictionary move;
            if (element == null) {
                move = W3CActions.PointerMove("pointer", xoffset, yoffset);
            } else if (xoffset == 0 && yoffset == 0) {
                return moveTo(element);
            } else {
                // The W3C offsets are relative to the center of the element, not its top-left corner
                ScrollIntoViewIfNeeded(element);
                var size = element.Size();
                move = W3CActions.PointerMove(element.SerializeJson()
                    , xoffset - size.Width / 2, yoffset - size.Height / 2);
            }
            W3CActions.PerformPointer(_session, move);
            return this;
        }

        /// <summary>
        /// Click any mouse button (at the coordinates set by the last moveto command).
        /// </summary>
        /// <param name="button">{number} Which button, enum: {LEFT = 0, MIDDLE = 1 , RIGHT = 2}. Defaults to the left mouse button if not specified.</param>
        public Mouse Click(MouseButton button = MouseButton.Left) {
            W3CActions.PerformPointer(_session
                , W3CActions.PointerDown(button), W3CActions.PointerUp(button));
            return this;
        }

        /// <summary>
        /// Click and hold the left mouse button (at the coordinates set by the last moveto command).
        /// </summary>
        /// <param name="button">{number} Which button, enum: {LEFT = 0, MIDDLE = 1 , RIGHT = 2}. Defaults to the left mouse button if not specified.</param>
        public Mouse ClickAndHold(MouseButton button = MouseButton.Left) {
            W3CActions.PerformPointer(_session, W3CActions.PointerDown(button));
            return this;
        }

        /// <summary>
        /// Releases the mouse button previously held (where the mouse is currently at).
        /// </summary>
        /// <param name="button">{number} Which button, enum: {LEFT = 0, MIDDLE = 1 , RIGHT = 2}. Defaults to the left mouse button if not specified.</param>
        public Mouse Release(MouseButton button = MouseButton.Left) {
            W3CActions.PerformPointer(_session, W3CActions.PointerUp(button));
            return this;
        }

        /// <summary>
        /// Double-clicks at the current mouse coordinates (set by moveto).
        /// </summary>
        public Mouse ClickDouble() {
            W3CActions.PerformPointer(_session
                , W3CActions.PointerDown(MouseButton.Left), W3CActions.PointerUp(MouseButton.Left)
                , W3CActions.PointerDown(MouseButton.Left), W3CActions.PointerUp(MouseButton.Left));
            return this;
        }

        /// <summary>
        /// The legacy moveto scrolled the element into view, a W3C pointer move fails instead.
        /// </summary>
        private void ScrollIntoViewIfNeeded(WebElement element) {
            _session.javascript.Execute(
                "var e=arguments[0],r=e.getBoundingClientRect();"
                + "if(r.bottom<0||r.right<0||r.top>=innerHeight||r.left>=innerWidth)"
                + "e.scrollIntoView({block:'center',inline:'center'});"
                , new[] { element }, false);
        }

    }

}
