using Selenium.Core;
using System;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Runtime.InteropServices;

namespace Selenium.Internal {

    class EndPointExt : IDisposable {

        /// <summary>
        /// Lock a new TCP end point that can be unlocked later.
        /// </summary>
        public static EndPointExt Create(IPAddress address, bool disableInheritence) {
            Socket socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            socket.NoDelay = true;
            socket.ReceiveBufferSize = 0;
            socket.SendBufferSize = 0;
            socket.Bind(new IPEndPoint(address, 0));

            //Disable inheritance to the child processes so the main process can close the
            //socket once a child process is launched.
            if (disableInheritence)
                Native.SetHandleInformation(socket.Handle, 1, 0);

            return new EndPointExt(socket, (IPEndPoint)socket.LocalEndPoint);
        }

        internal static EndPointExt Parse(string address) {
            string[] parts = address.Split(':');
            IPAddress ipAddress = IPAddress.Parse(parts[0]);
            int ipPort = int.Parse(parts[1]);
            IPEndPoint endPoint = new IPEndPoint(ipAddress, ipPort);
            return new EndPointExt(null, endPoint);
        }

        Socket _socket;
        IPEndPoint _ipEndPoint;

        public EndPointExt(Socket socket, IPEndPoint endpoint) {
            _socket = socket;
            _ipEndPoint = endpoint;
        }

        public void Dispose() {
            if (_socket == null)
                return;
            _socket.Close();
            _socket = null;
        }

        public IPEndPoint IPEndPoint {
            get {
                return _ipEndPoint;
            }
        }
        public override string ToString() {
            return _ipEndPoint.ToString();
        }

        /// <summary>
        /// Waits for a local port to accept connections.
        /// </summary>
        /// <remarks>
        /// The readiness is checked by connecting to the port. The former check read the TCP table
        /// (GetTcpTable) and waited for the LISTEN state, which some Windows builds (KB5127936)
        /// reported as 0, so Start failed after 10s although the driver was running.
        /// </remarks>
        /// <param name="timeout">Timeout in milliseconds</param>
        /// <param name="delay">Time to wait in milliseconds before checking again</param>
        /// <param name="hasExited">Optional - Returns true when the process that should listen has exited, to fail at once</param>
        /// <returns>True when the port accepts connections, false after the timeout</returns>
        public bool WaitForListening(int timeout, int delay, Func<bool> hasExited = null) {
            // Release the port reserved for the driver
            if (_socket != null) {
                _socket.Close();
                _socket = null;
            }

            DateTime endtime = DateTime.UtcNow.AddMilliseconds(timeout);
            while (true) {
                // Windows reports a refused loopback connection after about 1s (SYN retries):
                // a short attempt keeps the polling frequent
                if (TryConnect(_ipEndPoint, 250))
                    return true;
                if (hasExited != null && hasExited())
                    return false;
                if (DateTime.UtcNow > endtime)
                    return false;
                SysWaiter.Wait(delay);
            }
        }

        /// <summary>
        /// Returns true when a TCP connection to the end point succeeds within the timeout.
        /// </summary>
        private static bool TryConnect(IPEndPoint endPoint, int timeout) {
            // A socket that failed to connect cannot be reused: create one per attempt
            using (var socket = new Socket(endPoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp)) {
                socket.NoDelay = true;
                try {
                    IAsyncResult result = socket.BeginConnect(endPoint, null, null);
                    if (!result.AsyncWaitHandle.WaitOne(timeout))
                        return false;
                    socket.EndConnect(result);
                    socket.Shutdown(SocketShutdown.Both);
                    return true;
                } catch (SocketException) {
                    return false;
                } catch (ObjectDisposedException) {
                    return false;
                }
            }
        }

        static class Native {

            const string KERNEL32 = "kernel32.dll";

            [DllImport(KERNEL32)]
            internal static extern bool SetHandleInformation(IntPtr hObject, int dwMask, int dwFlags);

        }

    }

}
