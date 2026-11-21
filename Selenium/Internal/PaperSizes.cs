using System;
using System.Globalization;

namespace Selenium.Internal {

    /// <summary>
    /// Paper sizes in centimeters for PrintToPDF.
    /// </summary>
    static class PaperSizes {

        /// <summary>
        /// Returns false when no size is given. Raises an error for an unknown size.
        /// </summary>
        public static bool TryGet(string name, out double width, out double height) {
            width = height = 0;
            if (string.IsNullOrEmpty(name) || name.Trim().Length == 0)
                return false;

            switch (name.Trim().ToUpperInvariant()) {
                case "A3": width = 29.7; height = 42.0; return true;
                case "A4": width = 21.0; height = 29.7; return true;
                case "A5": width = 14.8; height = 21.0; return true;
                case "B4": width = 25.7; height = 36.4; return true;   // JIS B4, the common B4 in Korea and Japan
                case "B5": width = 18.2; height = 25.7; return true;   // JIS B5
                case "LETTER": width = 21.59; height = 27.94; return true;
                case "LEGAL": width = 21.59; height = 35.56; return true;
                case "TABLOID": width = 27.94; height = 43.18; return true;
            }

            // Custom size: "width x height" in cm
            string[] parts = name.ToLowerInvariant().Replace("cm", "").Split('x', '*', '×');
            if (parts.Length == 2
                && double.TryParse(parts[0].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out width)
                && double.TryParse(parts[1].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out height)
                && width > 0 && height > 0)
                return true;

            throw new Errors.ArgumentError("Invalid paper size: {0}. Use A3, A4, A5, B4, B5, Letter, Legal, Tabloid or \"width x height\" in cm.", name);
        }

    }

}
