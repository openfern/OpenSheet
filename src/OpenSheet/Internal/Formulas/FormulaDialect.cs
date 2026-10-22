using System;
using System.Collections.Generic;
using System.Text;

namespace OpenSheet.Internal.Formulas
{
    /// <summary>
    /// Excel stores functions added after Excel 2007 with a prefix (<c>_xlfn.VAR.P</c>) and shows them without it.
    /// Callers of OpenSheet use the plain name, so formulas are translated on the way in and out.
    /// </summary>
    static class FormulaDialect
    {
        // Functions Excel writes with the "_xlfn." prefix (Excel 2010 and later).
        static readonly HashSet<string> Future = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "ACOT", "ACOTH", "AGGREGATE", "ARABIC", "ARRAYTOTEXT", "BASE", "BETA.DIST", "BETA.INV", "BINOM.DIST", "BINOM.DIST.RANGE", "BINOM.INV",
            "BITAND", "BITLSHIFT", "BITOR", "BITRSHIFT", "BITXOR", "BYCOL", "BYROW", "CEILING.MATH", "CEILING.PRECISE", "CHISQ.DIST", "CHISQ.DIST.RT",
            "CHISQ.INV", "CHISQ.INV.RT", "CHISQ.TEST", "CHOOSECOLS", "CHOOSEROWS", "COMBINA", "CONCAT", "CONFIDENCE.NORM", "CONFIDENCE.T", "COT", "COTH",
            "COVARIANCE.P", "COVARIANCE.S", "CSC", "CSCH", "DAYS", "DECIMAL", "DROP", "ERF.PRECISE", "ERFC.PRECISE", "EXPAND", "EXPON.DIST", "F.DIST",
            "F.DIST.RT", "F.INV", "F.INV.RT", "F.TEST", "FILTER", "FILTERXML", "FLOOR.MATH", "FLOOR.PRECISE", "FORECAST.ETS", "FORECAST.ETS.CONFINT",
            "FORECAST.ETS.SEASONALITY", "FORECAST.ETS.STAT", "FORECAST.LINEAR", "FORMULATEXT", "GAMMA", "GAMMA.DIST", "GAMMA.INV", "GAMMALN.PRECISE",
            "GAUSS", "HSTACK", "HYPGEOM.DIST", "IFNA", "IFS", "IMCOSH", "IMCOT", "IMCSC", "IMCSCH", "IMSEC", "IMSECH", "IMSINH", "IMTAN", "ISFORMULA",
            "ISOMITTED", "ISOWEEKNUM", "LAMBDA", "LET", "LOGNORM.DIST", "LOGNORM.INV", "MAKEARRAY", "MAP", "MAXIFS", "MINIFS", "MODE.MULT", "MODE.SNGL",
            "MUNIT", "NEGBINOM.DIST", "NETWORKDAYS.INTL", "NORM.DIST", "NORM.INV", "NORM.S.DIST", "NORM.S.INV", "NUMBERVALUE", "PDURATION",
            "PERCENTILE.EXC", "PERCENTILE.INC", "PERCENTRANK.EXC", "PERCENTRANK.INC", "PERMUTATIONA", "PHI", "POISSON.DIST", "QUARTILE.EXC",
            "QUARTILE.INC", "QUERYSTRING", "RANDARRAY", "RANK.AVG", "RANK.EQ", "REDUCE", "RRI", "SCAN", "SEC", "SECH", "SEQUENCE", "SHEET", "SHEETS",
            "SKEW.P", "SORT", "SORTBY", "STDEV.P", "STDEV.S", "SWITCH", "T.DIST", "T.DIST.2T", "T.DIST.RT", "T.INV", "T.INV.2T", "T.TEST", "TAKE",
            "TEXTAFTER", "TEXTBEFORE", "TEXTJOIN", "TEXTSPLIT", "TOCOL", "TOROW", "UNICHAR", "UNICODE", "UNIQUE", "VALUETOTEXT", "VAR.P", "VAR.S",
            "VSTACK", "WEBSERVICE", "WEIBULL.DIST", "WORKDAY.INTL", "WRAPCOLS", "WRAPROWS", "XLOOKUP", "XMATCH", "XOR", "Z.TEST",
        };

        const string Prefix = "_xlfn.";

        /// <summary>Plain formula as a person writes it → the form Excel stores.</summary>
        public static string ToStored(string formula)
        {
            return Rewrite(formula, (name, hadPrefix) => !hadPrefix && Future.Contains(name) ? Prefix + name : null);
        }

        /// <summary>The form Excel stores → the plain formula a person reads.</summary>
        public static string ToDisplay(string stored)
        {
            if (stored == null || stored.IndexOf("_xl", StringComparison.OrdinalIgnoreCase) < 0) return stored;
            return Rewrite(stored, (name, hadPrefix) => hadPrefix ? name : null);
        }

        // Walks the text, leaves string literals and quoted sheet names alone, and offers each "name(" to the callback.
        static string Rewrite(string text, Func<string, bool, string> replace)
        {
            if (string.IsNullOrEmpty(text)) return text;
            var sb = new StringBuilder(text.Length + 16);
            var i = 0;
            while (i < text.Length)
            {
                var ch = text[i];
                if (ch == '"' || ch == '\'')
                {
                    var j = i + 1;
                    while (j < text.Length)
                    {
                        if (text[j] == ch) { if (j + 1 < text.Length && text[j + 1] == ch) { j += 2; continue; } break; }
                        j++;
                    }
                    var end = Math.Min(j + 1, text.Length);
                    sb.Append(text, i, end - i);
                    i = end;
                    continue;
                }
                if (IsStart(ch) && (i == 0 || !IsPart(text[i - 1])))
                {
                    var j = i;
                    while (j < text.Length && IsPart(text[j])) j++;
                    var k = j;
                    while (k < text.Length && text[k] == ' ') k++;
                    var token = text.Substring(i, j - i);
                    if (k < text.Length && text[k] == '(')
                    {
                        var hadPrefix = token.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase) || token.StartsWith("_xlws.", StringComparison.OrdinalIgnoreCase);
                        var name = hadPrefix ? token.Substring(6) : token;
                        var swapped = replace(name, hadPrefix);
                        sb.Append(swapped ?? token);
                    }
                    else sb.Append(token);
                    i = j;
                    continue;
                }
                sb.Append(ch);
                i++;
            }
            return sb.ToString();
        }

        static bool IsStart(char c) => char.IsLetter(c) || c == '_';
        static bool IsPart(char c) => char.IsLetterOrDigit(c) || c == '_' || c == '.';
    }
}
