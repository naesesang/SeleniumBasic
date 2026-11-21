using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Selenium.ComInterfaces {
#pragma warning disable 1591

    [Guid("0277FC34-FD1B-4616-BB19-2276E80F5CF7")]
    [ComVisible(true), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface _Cookie {

        [DispId(30), Description("Gets the name of the cookie.")]
        string Name { get; }

        [DispId(32), Description("Gets the value of the cookie.")]
        string Value { get; }

        [DispId(34), Description("Gets the path of the cookie.")]
        string Path { get; }

        [DispId(36), Description("Gets the domain of the cookie.")]
        string Domain { get; }

        [DispId(38), Description("Gets a value indicating whether the cookie is secure.")]
        bool Secure { get; }

        [DispId(42), Description("Gets the SameSite policy of the cookie: Strict, Lax, None or an empty string.")]
        string SameSite { get; }

        [DispId(44), Description("Gets a value indicating whether the cookie is hidden from the scripts.")]
        bool HttpOnly { get; }

        [DispId(40), Description("Gets the expiration date of the cookie.")]
        string Expiry { get; }

        [DispId(245), Description("Deletes the specified cookie from the page.")]
        void Delete();

    }

}
