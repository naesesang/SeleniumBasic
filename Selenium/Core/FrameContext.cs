using System;

namespace Selenium.Core {

    /// <summary>
    /// Context manager
    /// </summary>
    class FrameContext {

        private RemoteSession _session;

        /// <summary>
        /// Contructor over a session
        /// </summary>
        /// <param name="session"></param>
        internal FrameContext(RemoteSession session) {
            _session = session;
        }

        /// <summary>
        /// Switches focus to the specified frame, by index, name or WebElement.
        /// </summary>
        /// <param name="identifier">The name, id, or WebElement of the frame to switch.</param>
        /// <param name="timeout">Optional timeout in milliseconds</param>
        public void SwitchToFrame(object identifier, int timeout) {
            if (identifier == null)
                throw new Errors.ArgumentError("Invalid type for argument identifier");

            try {
                SendSwitchToFrame(identifier);
            } catch (Errors.NoSuchFrameError) {
                if (timeout == 0)
                    throw;
                var endTime = _session.GetEndTime(timeout);
                while (true) {
                    SysWaiter.Wait();
                    try {
                        SendSwitchToFrame(identifier);
                        break;
                    } catch (Errors.NoSuchFrameError) {
                        if (DateTime.UtcNow > endTime)
                            throw;
                    }
                }
            }
        }

        /// <summary>
        /// The W3C protocol only accepts an index or an element, so a name or id
        /// is first resolved to the frame element.
        /// </summary>
        private void SendSwitchToFrame(object identifier) {
            var element = identifier as WebElement;
            var name = identifier as string;
            if (name != null) {
                string css = "frame[name=" + By.CssQuote(name) + "],iframe[name=" + By.CssQuote(name) + "]"
                    + ",frame[id=" + By.CssQuote(name) + "],iframe[id=" + By.CssQuote(name) + "]";
                var locator = By.FormatLocator(Strategy.Css, css);
                List frames = (List)_session.Send(RequestMethod.POST, "/elements", locator);
                if (frames.Count == 0)
                    throw new Errors.NoSuchFrameError(name);
                element = new WebElement(_session, (Dictionary)frames[0]);
            }
            if (element != null) {
                identifier = element.SerializeJson();
            }
            _session.Send(RequestMethod.POST, "/frame", "id", identifier);
        }

        /// <summary>
        /// Select the parent frame of the currently selected frame.
        /// </summary>
        public void SwitchToParentFrame() {
            _session.Send(RequestMethod.POST, "/frame/parent");
        }

        /// <summary>
        /// Selects either the first frame on the page or the main document when a page contains iFrames.
        /// </summary>
        public void SwitchToDefaultContent() {
            _session.Send(RequestMethod.POST, "/frame", "id", null);
        }

    }

}
