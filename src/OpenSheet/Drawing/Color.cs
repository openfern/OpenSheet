using System;
using System.Globalization;

namespace OpenSheet.Drawing
{
    /// <summary>An opaque RGB color. Kept tiny on purpose so the library needs no graphics dependency.</summary>
    public readonly struct Color : IEquatable<Color>
    {
        public Color(byte r, byte g, byte b) { R = r; G = g; B = b; }

        public byte R { get; }
        public byte G { get; }
        public byte B { get; }

        public static Color FromRgb(byte r, byte g, byte b) => new Color(r, g, b);

        /// <summary>Parses "RRGGBB" or "#RRGGBB".</summary>
        public static Color FromHex(string hex)
        {
            if (hex == null) throw new ArgumentNullException(nameof(hex));
            var s = hex.Trim().TrimStart('#');
            if (s.Length != 6 || !int.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var v))
                throw new ArgumentException($"'{hex}' is not a color in RRGGBB form.", nameof(hex));
            return new Color((byte)(v >> 16), (byte)(v >> 8), (byte)v);
        }

        /// <summary>"#RRGGBB" in upper case.</summary>
        public string ToHex() => $"#{R:X2}{G:X2}{B:X2}";

        public bool Equals(Color other) => R == other.R && G == other.G && B == other.B;
        public override bool Equals(object obj) => obj is Color c && Equals(c);
        public override int GetHashCode() => (R << 16) | (G << 8) | B;
        public override string ToString() => ToHex();
    }
}
