using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using NPOI.SS.Formula.Eval;

namespace OpenSheet.Internal.Formulas
{
    static partial class FunctionPack
    {
        static double Bits(Ctx c, ValueEval v)
        {
            var x = c.Num(v);
            if (x < 0 || x != Math.Floor(x) || x >= 281474976710656.0) throw XlError.Num;
            return x;
        }

        static void DefineEngineering()
        {
            Add("BITAND", (c, a) => { Need(a, 2, 2); return N((long)Bits(c, a[0]) & (long)Bits(c, a[1])); });
            Add("BITOR", (c, a) => { Need(a, 2, 2); return N((long)Bits(c, a[0]) | (long)Bits(c, a[1])); });
            Add("BITXOR", (c, a) => { Need(a, 2, 2); return N((long)Bits(c, a[0]) ^ (long)Bits(c, a[1])); });
            Add("BITLSHIFT", (c, a) => Shift(c, a, left: true));
            Add("BITRSHIFT", (c, a) => Shift(c, a, left: false));
            Add("GESTEP", (c, a) => { Need(a, 1, 2); return N(c.Num(a[0]) >= c.Num(a, 1, 0) ? 1 : 0); });
            Add("ERF", (c, a) =>
            {
                Need(a, 1, 2);
                var lo = c.Num(a[0]);
                return N(a.Length == 1 ? Special.Erf(lo) : Special.Erf(c.Num(a[1])) - Special.Erf(lo));
            });
            Add("ERF.PRECISE", (c, a) => { Need(a, 1, 1); return N(Special.Erf(c.Num(a[0]))); });
            Add("ERFC", (c, a) => { Need(a, 1, 1); return N(Special.Erfc(c.Num(a[0]))); });
            Add("ERFC.PRECISE", (c, a) => { Need(a, 1, 1); return N(Special.Erfc(c.Num(a[0]))); });

            Add("IMABS", (c, a) => { Need(a, 1, 1); var z = ComplexArg(c, a[0]); return N(Hypot(z.Re, z.Im)); });
            Add("IMREAL", (c, a) => { Need(a, 1, 1); return N(ComplexArg(c, a[0]).Re); });
            Add("IMAGINARY", (c, a) => { Need(a, 1, 1); return N(ComplexArg(c, a[0]).Im); });
            Corrections["IMREAL"] = Map["IMREAL"];       // NPOI answers with text
            Corrections["IMAGINARY"] = Map["IMAGINARY"];
            Add("IMARGUMENT", (c, a) => { Need(a, 1, 1); var z = ComplexArg(c, a[0]); if (z.Re == 0 && z.Im == 0) throw XlError.Div0; return N(Math.Atan2(z.Im, z.Re)); });
            Add("IMCONJUGATE", (c, a) => { Need(a, 1, 1); var z = ComplexArg(c, a[0]); return S(new Cx(z.Re, -z.Im, z.Suffix).ToString()); });
            Add("IMSUM", (c, a) => ComplexFold(c, a, (x, y) => new Cx(x.Re + y.Re, x.Im + y.Im, x.Suffix)));
            Add("IMSUB", (c, a) => { Need(a, 2, 2); var x = ComplexArg(c, a[0]); var y = ComplexArg(c, a[1]); return S(new Cx(x.Re - y.Re, x.Im - y.Im, x.Suffix).ToString()); });
            Add("IMPRODUCT", (c, a) => ComplexFold(c, a, (x, y) => new Cx(x.Re * y.Re - x.Im * y.Im, x.Re * y.Im + x.Im * y.Re, x.Suffix)));
            Add("IMDIV", (c, a) =>
            {
                Need(a, 2, 2);
                var x = ComplexArg(c, a[0]); var y = ComplexArg(c, a[1]);
                var d = y.Re * y.Re + y.Im * y.Im;
                if (d == 0) throw XlError.Num;
                return S(new Cx((x.Re * y.Re + x.Im * y.Im) / d, (x.Im * y.Re - x.Re * y.Im) / d, x.Suffix).ToString());
            });
            Add("CONVERT", (c, a) =>
            {
                Need(a, 3, 3);
                return N(Units.Convert(c.Num(a[0]), c.Str(a[1]), c.Str(a[2])));
            });
        }

        static ValueEval Shift(Ctx c, ValueEval[] a, bool left)
        {
            Need(a, 2, 2);
            var n = (long)Bits(c, a[0]); var s = (int)Math.Truncate(c.Num(a[1]));
            if (Math.Abs(s) > 53) throw XlError.Num;
            var toLeft = left == (s >= 0);
            s = Math.Abs(s);
            var r = toLeft ? n << s : n >> s;
            if (r >= 281474976710656L) throw XlError.Num;
            return N(r);
        }

        static ValueEval ComplexFold(Ctx c, ValueEval[] a, Func<Cx, Cx, Cx> f)
        {
            if (a.Length == 0) throw XlError.Value;
            var items = a.SelectMany(x => x is AreaEval ? c.Area(x).Flat() : new[] { x }).Select(v => ComplexArg(c, v)).ToList();
            return S(items.Skip(1).Aggregate(items[0], f).ToString());
        }

        static Cx ComplexArg(Ctx c, ValueEval v)
        {
            v = c.Scalar(v);
            if (v is BoolEval) throw XlError.Value;
            if (v is StringEval s && s.StringValue.Trim().Length == 0) throw XlError.Num;
            return Cx.Parse(c.Str(v));
        }

        static double Hypot(double x, double y)
        {
            x = Math.Abs(x); y = Math.Abs(y);
            if (x < y) { var t = x; x = y; y = t; }
            return x == 0 ? 0 : x * Math.Sqrt(1 + (y / x) * (y / x));
        }

        /// <summary>A complex number written the way Excel's IM* functions take it: "3+4i", "-2j", "i", "5".</summary>
        struct Cx
        {

            public Cx(double re, double im, char suffix) { Re = re; Im = im; Suffix = suffix; }
            public double Re { get; }
            public double Im { get; }
            public char Suffix { get; }

            public static Cx Parse(string text)
            {
                text = text.Trim();
                if (text.Length == 0) return new Cx(0, 0, 'i');
                // split "a+bi" at the last sign that is not part of an exponent
                var last = text[text.Length - 1];
                if (last != 'i' && last != 'j' && last != 'I' && last != 'J')
                    return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var r) ? new Cx(r, 0, 'i') : throw XlError.Num;
                var body = text.Substring(0, text.Length - 1); var suffix = char.ToLowerInvariant(last);
                var cut = -1;
                for (var i = body.Length - 1; i > 0; i--)
                    if ((body[i] == '+' || body[i] == '-') && char.ToLowerInvariant(body[i - 1]) != 'e') { cut = i; break; }
                string rePart = cut > 0 ? body.Substring(0, cut) : "", imPart = cut > 0 ? body.Substring(cut) : body;
                double Number(string s, double ifEmpty)
                {
                    if (s == "" || s == "+") return ifEmpty;
                    if (s == "-") return -ifEmpty;
                    return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : throw XlError.Num;
                }
                return new Cx(rePart == "" ? 0 : Number(rePart, 0), Number(imPart, 1), suffix);
            }

            public override string ToString()
            {
                string F(double d) => Math.Abs(d) < 1e-15 ? "0" : d.ToString("G15", CultureInfo.InvariantCulture);
                if (Im == 0) return F(Re);
                var im = Im == 1 ? "" : Im == -1 ? "-" : F(Im);
                if (Re == 0) return im + Suffix;
                return F(Re) + (Im > 0 && !im.StartsWith("-", StringComparison.Ordinal) ? "+" : "") + im + Suffix;
            }
        }

        // ---- CONVERT ----

        static class Units
        {
            sealed class Unit
            {
                public Unit(string group, double factor, bool metric = false, bool binary = false) { Group = group; Factor = factor; Metric = metric; Binary = binary; }
                public string Group { get; }
                public double Factor { get; }       // size of one unit in the group's base unit
                public bool Metric { get; }         // accepts SI prefixes
                public bool Binary { get; }         // accepts ki, Mi, ... prefixes too
            }

            static readonly Dictionary<string, Unit> Table = new Dictionary<string, Unit>(StringComparer.Ordinal)
            {
                // mass, base gram
                ["g"] = new Unit("mass", 1, true), ["sg"] = new Unit("mass", 14593.8424189287), ["lbm"] = new Unit("mass", 453.59237), ["u"] = new Unit("mass", 1.66053886282693e-24, true),
                ["ozm"] = new Unit("mass", 28.349523125), ["grain"] = new Unit("mass", 0.06479891), ["cwt"] = new Unit("mass", 45359.237), ["shweight"] = new Unit("mass", 45359.237),
                ["uk_cwt"] = new Unit("mass", 50802.34544), ["lcwt"] = new Unit("mass", 50802.34544), ["hweight"] = new Unit("mass", 50802.34544), ["stone"] = new Unit("mass", 6350.29318),
                ["ton"] = new Unit("mass", 907184.74), ["brton"] = new Unit("mass", 1016046.9088), ["LTON"] = new Unit("mass", 1016046.9088), ["uk_ton"] = new Unit("mass", 1016046.9088),
                // distance, base metre
                ["m"] = new Unit("distance", 1, true), ["mi"] = new Unit("distance", 1609.344), ["Nmi"] = new Unit("distance", 1852), ["in"] = new Unit("distance", 0.0254), ["ft"] = new Unit("distance", 0.3048),
                ["yd"] = new Unit("distance", 0.9144), ["ang"] = new Unit("distance", 1e-10, true), ["ell"] = new Unit("distance", 1.143), ["ly"] = new Unit("distance", 9.4607304725808e15),
                ["parsec"] = new Unit("distance", 3.08567758128155e16), ["pc"] = new Unit("distance", 3.08567758128155e16), ["Pica"] = new Unit("distance", 0.0254 / 72), ["Picapt"] = new Unit("distance", 0.0254 / 72),
                ["pica"] = new Unit("distance", 0.0254 / 6), ["survey_mi"] = new Unit("distance", 1609.3472186944), ["us_survey_mi"] = new Unit("distance", 1609.3472186944),
                // time, base second
                ["yr"] = new Unit("time", 31557600), ["day"] = new Unit("time", 86400), ["d"] = new Unit("time", 86400), ["hr"] = new Unit("time", 3600), ["mn"] = new Unit("time", 60), ["min"] = new Unit("time", 60), ["sec"] = new Unit("time", 1, true), ["s"] = new Unit("time", 1, true),
                // pressure, base pascal
                ["Pa"] = new Unit("pressure", 1, true), ["p"] = new Unit("pressure", 1, true), ["atm"] = new Unit("pressure", 101325, true), ["at"] = new Unit("pressure", 101325, true), ["mmHg"] = new Unit("pressure", 133.322363925, true),
                ["psi"] = new Unit("pressure", 6894.75729316836), ["Torr"] = new Unit("pressure", 133.322368421053),
                // force, base newton
                ["N"] = new Unit("force", 1, true), ["dyn"] = new Unit("force", 1e-5, true), ["dy"] = new Unit("force", 1e-5, true), ["lbf"] = new Unit("force", 4.4482216152605), ["pond"] = new Unit("force", 0.00980665, true),
                // energy, base joule
                ["J"] = new Unit("energy", 1, true), ["e"] = new Unit("energy", 1e-7, true), ["c"] = new Unit("energy", 4.184), ["cal"] = new Unit("energy", 4.1868), ["eV"] = new Unit("energy", 1.602176487e-19, true),
                ["ev"] = new Unit("energy", 1.602176487e-19, true), ["HPh"] = new Unit("energy", 2684519.53769617), ["hh"] = new Unit("energy", 2684519.53769617), ["Wh"] = new Unit("energy", 3600, true),
                ["wh"] = new Unit("energy", 3600, true), ["flb"] = new Unit("energy", 1.3558179483314), ["BTU"] = new Unit("energy", 1055.05585262), ["btu"] = new Unit("energy", 1055.05585262),
                // power, base watt
                ["W"] = new Unit("power", 1, true), ["w"] = new Unit("power", 1, true), ["HP"] = new Unit("power", 745.69987158227), ["h"] = new Unit("power", 745.69987158227), ["PS"] = new Unit("power", 735.49875),
                // magnetism, base tesla
                ["T"] = new Unit("magnetism", 1, true), ["ga"] = new Unit("magnetism", 1e-4, true),
                // volume, base litre... stored in cubic metres
                ["l"] = new Unit("volume", 0.001, true), ["L"] = new Unit("volume", 0.001, true), ["lt"] = new Unit("volume", 0.001, true), ["tsp"] = new Unit("volume", 4.92892159375e-6), ["tbs"] = new Unit("volume", 1.478676478125e-5),
                ["oz"] = new Unit("volume", 2.95735295625e-5), ["cup"] = new Unit("volume", 2.365882365e-4), ["pt"] = new Unit("volume", 4.73176473e-4), ["us_pt"] = new Unit("volume", 4.73176473e-4),
                ["uk_pt"] = new Unit("volume", 5.68261225e-4), ["qt"] = new Unit("volume", 9.46352946e-4), ["gal"] = new Unit("volume", 3.785411784e-3), ["m3"] = new Unit("volume", 1, true),
                ["ang3"] = new Unit("volume", 1e-30, true), ["barrel"] = new Unit("volume", 0.158987294928), ["bushel"] = new Unit("volume", 0.03523907016688), ["regton"] = new Unit("volume", 2.8316846592),
                ["GRT"] = new Unit("volume", 2.8316846592), ["Straight"] = new Unit("volume", 2.8316846592), ["MTON"] = new Unit("volume", 1.13267386368), ["in3"] = new Unit("volume", 1.6387064e-5), ["ft3"] = new Unit("volume", 0.028316846592),
                ["yd3"] = new Unit("volume", 0.764554857984), ["mi3"] = new Unit("volume", 4168181825.44058), ["Nmi3"] = new Unit("volume", 6352182208), ["ly3"] = new Unit("volume", 8.46786664623715e47),
                // area, base square metre
                ["m2"] = new Unit("area", 1, true), ["in2"] = new Unit("area", 0.00064516), ["ft2"] = new Unit("area", 0.09290304), ["yd2"] = new Unit("area", 0.83612736), ["mi2"] = new Unit("area", 2589988.110336),
                ["ha"] = new Unit("area", 10000, true), ["ar"] = new Unit("area", 100, true), ["ang2"] = new Unit("area", 1e-20, true), ["uk_acre"] = new Unit("area", 4046.8564224), ["us_acre"] = new Unit("area", 4046.87260987425), ["Morgen"] = new Unit("area", 2500),
                ["Nmi2"] = new Unit("area", 3429904), ["ly2"] = new Unit("area", 8.95054210748189e31), ["Pica2"] = new Unit("area", 1.24344e-7), ["Picapt2"] = new Unit("area", 1.24344e-7),
                // speed, base metre per second
                ["m/s"] = new Unit("speed", 1, true), ["m/sec"] = new Unit("speed", 1, true), ["m/h"] = new Unit("speed", 1.0 / 3600, true), ["m/hr"] = new Unit("speed", 1.0 / 3600, true), ["mph"] = new Unit("speed", 0.44704),
                ["kn"] = new Unit("speed", 1852.0 / 3600), ["admkn"] = new Unit("speed", 1853.184 / 3600),
                // information, base bit
                ["bit"] = new Unit("information", 1, true, true), ["byte"] = new Unit("information", 8, true, true),
            };

            // prefix → exponent of ten (or the binary power for the "i" prefixes)
            static readonly Dictionary<string, double> Prefixes = new Dictionary<string, double>(StringComparer.Ordinal)
            {
                ["Y"] = 1e24, ["Z"] = 1e21, ["E"] = 1e18, ["P"] = 1e15, ["T"] = 1e12, ["G"] = 1e9, ["M"] = 1e6, ["k"] = 1e3, ["h"] = 1e2, ["da"] = 1e1, ["e"] = 1e1,
                ["d"] = 1e-1, ["c"] = 1e-2, ["m"] = 1e-3, ["u"] = 1e-6, ["n"] = 1e-9, ["p"] = 1e-12, ["f"] = 1e-15, ["a"] = 1e-18, ["z"] = 1e-21, ["y"] = 1e-24,
            };

            static readonly Dictionary<string, double> BinaryPrefixes = new Dictionary<string, double>(StringComparer.Ordinal)
            {
                ["Yi"] = Math.Pow(2, 80), ["Zi"] = Math.Pow(2, 70), ["Ei"] = Math.Pow(2, 60), ["Pi"] = Math.Pow(2, 50), ["Ti"] = Math.Pow(2, 40), ["Gi"] = Math.Pow(2, 30), ["Mi"] = Math.Pow(2, 20), ["ki"] = Math.Pow(2, 10),
            };

            static bool Find(string name, out Unit unit, out double prefix)
            {
                prefix = 1;
                if (Table.TryGetValue(name, out unit)) return true;
                foreach (var kv in BinaryPrefixes.Where(p => name.StartsWith(p.Key, StringComparison.Ordinal)))
                    if (Table.TryGetValue(name.Substring(kv.Key.Length), out unit) && unit.Binary) { prefix = kv.Value; return true; }
                foreach (var kv in Prefixes.OrderByDescending(p => p.Key.Length).Where(p => name.StartsWith(p.Key, StringComparison.Ordinal)))
                    if (name.Length > kv.Key.Length && Table.TryGetValue(name.Substring(kv.Key.Length), out unit) && unit.Metric) { prefix = kv.Value; return true; }
                unit = null; return false;
            }

            static readonly string[] Temperatures = { "C", "cel", "F", "fah", "K", "kel", "Rank", "Reau" };

            public static double Convert(double value, string from, string to)
            {
                if (Array.IndexOf(Temperatures, from) >= 0 || Array.IndexOf(Temperatures, to) >= 0)
                {
                    if (Array.IndexOf(Temperatures, from) < 0 || Array.IndexOf(Temperatures, to) < 0) throw XlError.NA;
                    return FromKelvin(ToKelvin(value, from), to);
                }
                if (!Find(from, out var u1, out var p1) || !Find(to, out var u2, out var p2) || u1.Group != u2.Group) throw XlError.NA;
                // prefixed squared / cubed units scale by the power of the prefix
                var power1 = from.EndsWith("2", StringComparison.Ordinal) && u1.Group == "area" ? 2 : from.EndsWith("3", StringComparison.Ordinal) && u1.Group == "volume" ? 3 : 1;
                var power2 = to.EndsWith("2", StringComparison.Ordinal) && u2.Group == "area" ? 2 : to.EndsWith("3", StringComparison.Ordinal) && u2.Group == "volume" ? 3 : 1;
                return value * Math.Pow(p1, power1) * u1.Factor / (Math.Pow(p2, power2) * u2.Factor);
            }

            static double ToKelvin(double v, string u)
            {
                switch (u)
                {
                    case "C": case "cel": return v + 273.15;
                    case "F": case "fah": return (v - 32) * 5 / 9 + 273.15;
                    case "Rank": return v * 5 / 9;
                    case "Reau": return v * 5 / 4 + 273.15;
                    default: return v;
                }
            }

            static double FromKelvin(double k, string u)
            {
                switch (u)
                {
                    case "C": case "cel": return k - 273.15;
                    case "F": case "fah": return (k - 273.15) * 9 / 5 + 32;
                    case "Rank": return k * 9 / 5;
                    case "Reau": return (k - 273.15) * 4 / 5;
                    default: return k;
                }
            }
        }
    }
}
