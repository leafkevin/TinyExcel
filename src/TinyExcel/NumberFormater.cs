using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace TinyExcel;

/// <summary>
/// Parse ECMA-376 number format strings and format values like Excel and other spreadsheet softwares.
/// </summary>
public class NumberFormater
{
    public static string Format(object value, string formatString, CultureInfo culture, bool isDate1904)
    {
        var sections = Parser.ParseSections(formatString, out bool hasSyntaxError);
        if (hasSyntaxError)
            return ToCompatibleString(value, culture);

        var section = Evaluator.GetSection(sections, value);
        if (section == null)
            return ToCompatibleString(value, culture);

        try
        {
            return Format(value, section, culture, isDate1904);
        }
        catch (InvalidCastException)
        {
            // TimeSpan cast exception
            return ToCompatibleString(value, culture);
        }
        catch (FormatException)
        {
            // Convert.ToDouble/ToDateTime exceptions
            return ToCompatibleString(value, culture);
        }
    }
    private static string Format(object value, Section node, CultureInfo culture, bool isDate1904)
    {
        switch (node.Type)
        {
            case SectionType.Number:
                // Hide sign under certain conditions and section index
                var number = Convert.ToDouble(value, culture);
                if ((node.SectionIndex == 0 && node.Condition != null) || node.SectionIndex == 1)
                    number = Math.Abs(number);

                return FormatNumber(number, node.Number, culture);

            case SectionType.Date:
                if (ExcelDateTime.TryConvert(value, isDate1904, culture, out var excelDateTime))
                    return FormatDate(excelDateTime, node.GeneralTextDateDurationParts, culture);
                else throw new FormatException("Unexpected date value");

            case SectionType.Duration:
                if (value is TimeSpan ts)
                    return FormatTimeSpan(ts, node.GeneralTextDateDurationParts, culture);
                else
                {
                    var d = Convert.ToDouble(value);
                    return FormatTimeSpan(TimeSpan.FromDays(d), node.GeneralTextDateDurationParts, culture);
                }

            case SectionType.General:
            case SectionType.Text:
                return FormatGeneralText(ToCompatibleString(value, culture), node.GeneralTextDateDurationParts);

            case SectionType.Exponential:
                return FormatExponential(Convert.ToDouble(value, culture), node, culture);

            case SectionType.Fraction:
                return FormatFraction(Convert.ToDouble(value, culture), node, culture);

            default:
                throw new InvalidOperationException("Unknown number format section");
        }
    }
    private static string FormatGeneralText(string text, List<string> tokens)
    {
        var result = new StringBuilder();
        for (var i = 0; i < tokens.Count; i++)
        {
            var token = tokens[i];
            if (Token.IsGeneral(token) || token == "@")
            {
                result.Append(text);
            }
            else
            {
                FormatLiteral(token, result);
            }
        }
        return result.ToString();
    }
    private static string FormatTimeSpan(TimeSpan timeSpan, List<string> tokens, CultureInfo culture)
    {
        // NOTE/TODO: assumes there is exactly one [hh], [mm] or [ss] using the integer part of TimeSpan.TotalXXX when formatting.
        // The timeSpan input is then truncated to the remainder fraction, which is used to format mm and/or ss.
        var result = new StringBuilder();
        var containsMilliseconds = false;
        for (var i = tokens.Count - 1; i >= 0; i--)
        {
            if (tokens[i].StartsWith(".0"))
            {
                containsMilliseconds = true;
                break;
            }
        }

        for (var i = 0; i < tokens.Count; i++)
        {
            var token = tokens[i];

            if (token.StartsWith("m", StringComparison.OrdinalIgnoreCase))
            {
                var value = timeSpan.Minutes;
                var digits = token.Length;
                result.Append(value.ToString("D" + digits));
            }
            else if (token.StartsWith("s", StringComparison.OrdinalIgnoreCase))
            {
                // If format does not include ms, then include ms in seconds and round before printing
                var formatMs = containsMilliseconds ? 0 : timeSpan.Milliseconds / 1000D;
                var value = (int)Math.Round(timeSpan.Seconds + formatMs, 0, MidpointRounding.AwayFromZero);
                var digits = token.Length;
                result.Append(value.ToString("D" + digits));
            }
            else if (token.StartsWith("[h", StringComparison.OrdinalIgnoreCase))
            {
                var value = (int)timeSpan.TotalHours;
                var digits = token.Length - 2;
                result.Append(value.ToString("D" + digits));
                timeSpan = new TimeSpan(0, 0, Math.Abs(timeSpan.Minutes), Math.Abs(timeSpan.Seconds), Math.Abs(timeSpan.Milliseconds));
            }
            else if (token.StartsWith("[m", StringComparison.OrdinalIgnoreCase))
            {
                var value = (int)timeSpan.TotalMinutes;
                var digits = token.Length - 2;
                result.Append(value.ToString("D" + digits));
                timeSpan = new TimeSpan(0, 0, 0, Math.Abs(timeSpan.Seconds), Math.Abs(timeSpan.Milliseconds));
            }
            else if (token.StartsWith("[s", StringComparison.OrdinalIgnoreCase))
            {
                var value = (int)timeSpan.TotalSeconds;
                var digits = token.Length - 2;
                result.Append(value.ToString("D" + digits));
                timeSpan = new TimeSpan(0, 0, 0, 0, Math.Abs(timeSpan.Milliseconds));
            }
            else if (token.StartsWith(".0"))
            {
                var value = timeSpan.Milliseconds;
                var digits = token.Length - 1;
                result.Append("." + value.ToString("D" + digits));
            }
            else
            {
                FormatLiteral(token, result);
            }
        }

        return result.ToString();
    }
    private static string FormatDate(ExcelDateTime date, List<string> tokens, CultureInfo culture)
    {
        var containsAmPm = ContainsAmPm(tokens);

        var result = new StringBuilder();
        for (var i = 0; i < tokens.Count; i++)
        {
            var token = tokens[i];

            if (token.StartsWith("y", StringComparison.OrdinalIgnoreCase))
            {
                // year
                var digits = token.Length;
                if (digits < 2)
                    digits = 2;
                if (digits == 3)
                    digits = 4;

                var year = date.Year;
                if (digits == 2)
                    year = year % 100;

                result.Append(year.ToString("D" + digits));
            }
            else if (token.StartsWith("m", StringComparison.OrdinalIgnoreCase))
            {
                // If  "m" or "mm" code is used immediately after the "h" or "hh" code (for hours) or immediately before 
                // the "ss" code (for seconds), the application shall display minutes instead of the month. 
                if (LookBackDatePart(tokens, i - 1, "h") || LookAheadDatePart(tokens, i + 1, "s"))
                {
                    var digits = token.Length;
                    result.Append(date.Minute.ToString("D" + digits));
                }
                else
                {
                    var digits = token.Length;
                    if (digits == 3)
                    {
                        result.Append(culture.DateTimeFormat.AbbreviatedMonthNames[date.Month - 1]);
                    }
                    else if (digits == 4)
                    {
                        result.Append(culture.DateTimeFormat.MonthNames[date.Month - 1]);
                    }
                    else if (digits == 5)
                    {
                        result.Append(culture.DateTimeFormat.MonthNames[date.Month - 1][0]);
                    }
                    else
                    {
                        result.Append(date.Month.ToString("D" + digits));
                    }
                }
            }
            else if (token.StartsWith("d", StringComparison.OrdinalIgnoreCase))
            {
                var digits = token.Length;
                if (digits == 3)
                {
                    // Sun-Sat
                    result.Append(culture.DateTimeFormat.AbbreviatedDayNames[(int)date.DayOfWeek]);
                }
                else if (digits == 4)
                {
                    // Sunday-Saturday
                    result.Append(culture.DateTimeFormat.DayNames[(int)date.DayOfWeek]);
                }
                else
                {
                    result.Append(date.Day.ToString("D" + digits));
                }
            }
            else if (token.StartsWith("h", StringComparison.OrdinalIgnoreCase))
            {
                var digits = token.Length;
                if (containsAmPm)
                    result.Append(((date.Hour + 11) % 12 + 1).ToString("D" + digits));
                else
                    result.Append(date.Hour.ToString("D" + digits));
            }
            else if (token.StartsWith("s", StringComparison.OrdinalIgnoreCase))
            {
                var digits = token.Length;
                result.Append(date.Second.ToString("D" + digits));
            }
            else if (token.StartsWith("g", StringComparison.OrdinalIgnoreCase))
            {
                var era = culture.DateTimeFormat.Calendar.GetEra(date.AdjustedDateTime);
                var digits = token.Length;
                if (digits < 3)
                {
                    result.Append(culture.DateTimeFormat.GetAbbreviatedEraName(era));
                }
                else
                {
                    result.Append(culture.DateTimeFormat.GetEraName(era));
                }
            }
            else if (string.Compare(token, "am/pm", StringComparison.OrdinalIgnoreCase) == 0)
            {
                var ampm = date.ToString("tt", CultureInfo.InvariantCulture);
                result.Append(ampm.ToUpperInvariant());
            }
            else if (string.Compare(token, "a/p", StringComparison.OrdinalIgnoreCase) == 0)
            {
                var ampm = date.ToString("%t", CultureInfo.InvariantCulture);
                if (char.IsUpper(token[0]))
                {
                    result.Append(ampm.ToUpperInvariant());
                }
                else
                {
                    result.Append(ampm.ToLowerInvariant());
                }
            }
            else if (token.StartsWith(".0"))
            {
                var value = date.Millisecond;
                var digits = token.Length - 1;
                result.Append("." + value.ToString("D" + digits));
            }
            else if (token == "/")
            {
#if NETSTANDARD1_0
                    result.Append(DateTime.MaxValue.ToString("/d", culture)[0]);
#else
                result.Append(culture.DateTimeFormat.DateSeparator);
#endif
            }
            else if (token == ",")
            {
                while (i < tokens.Count - 1 && tokens[i + 1] == ",")
                {
                    i++;
                }

                result.Append(",");
            }
            else
            {
                FormatLiteral(token, result);
            }
        }

        return result.ToString();
    }
    private static bool LookAheadDatePart(List<string> tokens, int fromIndex, string startsWith)
    {
        for (var i = fromIndex; i < tokens.Count; i++)
        {
            var token = tokens[i];
            if (token.StartsWith(startsWith, StringComparison.OrdinalIgnoreCase))
                return true;
            if (Token.IsDatePart(token))
                return false;
        }

        return false;
    }
    private static bool LookBackDatePart(List<string> tokens, int fromIndex, string startsWith)
    {
        for (var i = fromIndex; i >= 0; i--)
        {
            var token = tokens[i];
            if (token.StartsWith(startsWith, StringComparison.OrdinalIgnoreCase))
                return true;
            if (Token.IsDatePart(token))
                return false;
        }

        return false;
    }
    private static bool ContainsAmPm(List<string> tokens)
    {
        foreach (var token in tokens)
        {
            if (string.Compare(token, "am/pm", StringComparison.OrdinalIgnoreCase) == 0)
            {
                return true;
            }

            if (string.Compare(token, "a/p", StringComparison.OrdinalIgnoreCase) == 0)
            {
                return true;
            }
        }

        return false;
    }
    private static string FormatNumber(double value, DecimalSection format, CultureInfo culture)
    {
        bool thousandSeparator = format.ThousandSeparator;
        value = value / format.ThousandDivisor;
        value = value * format.PercentMultiplier;

        var result = new StringBuilder();
        FormatNumber(value, format.BeforeDecimal, format.DecimalSeparator, format.AfterDecimal, thousandSeparator, culture, result);
        return result.ToString();
    }
    private static void FormatNumber(double value, List<string> beforeDecimal, bool decimalSeparator, List<string> afterDecimal, bool thousandSeparator, CultureInfo culture, StringBuilder result)
    {
        int signitificantDigits = 0;
        if (afterDecimal != null)
            signitificantDigits = GetDigitCount(afterDecimal);

        var valueString = Math.Abs(value).ToString("F" + signitificantDigits, CultureInfo.InvariantCulture);
        var valueStrings = valueString.Split('.');
        var thousandsString = valueStrings[0];
        var decimalString = valueStrings.Length > 1 ? valueStrings[1].TrimEnd('0') : "";

        if (value < 0)
        {
            result.Append("-");
        }

        if (beforeDecimal != null)
        {
            FormatThousands(thousandsString, thousandSeparator, false, beforeDecimal, culture, result);
        }

        if (decimalSeparator)
        {
            result.Append(culture.NumberFormat.NumberDecimalSeparator);
        }

        if (afterDecimal != null)
        {
            FormatDecimals(decimalString, afterDecimal, result);
        }
    }
    /// <summary>
    /// Prints right-aligned, left-padded integer before the decimal separator. With optional most-significant zero.
    /// </summary>
    private static void FormatThousands(string valueString, bool thousandSeparator, bool significantZero, List<string> tokens, CultureInfo culture, StringBuilder result)
    {
        var significant = false;
        var formatDigits = GetDigitCount(tokens);
        valueString = valueString.PadLeft(formatDigits, '0');

        // Print literals occurring before any placeholders
        var tokenIndex = 0;
        for (; tokenIndex < tokens.Count; tokenIndex++)
        {
            var token = tokens[tokenIndex];
            if (Token.IsPlaceholder(token))
                break;
            else
                FormatLiteral(token, result);
        }

        // Print value digits until there are as many digits remaining as there are placeholders
        var digitIndex = 0;
        for (; digitIndex < (valueString.Length - formatDigits); digitIndex++)
        {
            significant = true;
            result.Append(valueString[digitIndex]);

            if (thousandSeparator)
                FormatThousandSeparator(valueString, digitIndex, culture, result);
        }

        // Print remaining value digits and format literals
        for (; tokenIndex < tokens.Count; ++tokenIndex)
        {
            var token = tokens[tokenIndex];
            if (Token.IsPlaceholder(token))
            {
                var c = valueString[digitIndex];
                if (c != '0' || (significantZero && digitIndex == valueString.Length - 1)) significant = true;

                FormatPlaceholder(token, c, significant, result);

                if (thousandSeparator && (significant || token.Equals("0")))
                    FormatThousandSeparator(valueString, digitIndex, culture, result);

                digitIndex++;
            }
            else
            {
                FormatLiteral(token, result);
            }
        }
    }
    private static void FormatThousandSeparator(string valueString, int digit, CultureInfo culture, StringBuilder result)
    {
        var positionInTens = valueString.Length - 1 - digit;
        if (positionInTens > 0 && (positionInTens % 3) == 0)
        {
            result.Append(culture.NumberFormat.NumberGroupSeparator);
        }
    }
    /// <summary>
    /// Prints left-aligned, right-padded integer after the decimal separator. Does not print significant zero.
    /// </summary>
    private static void FormatDecimals(string valueString, List<string> tokens, StringBuilder result)
    {
        var significant = true;
        var unpaddedDigits = valueString.Length;
        var formatDigits = GetDigitCount(tokens);

        valueString = valueString.PadRight(formatDigits, '0');

        // Print all format digits
        var valueIndex = 0;
        for (var tokenIndex = 0; tokenIndex < tokens.Count; ++tokenIndex)
        {
            var token = tokens[tokenIndex];
            if (Token.IsPlaceholder(token))
            {
                var c = valueString[valueIndex];
                significant = valueIndex < unpaddedDigits;

                FormatPlaceholder(token, c, significant, result);
                valueIndex++;
            }
            else
            {
                FormatLiteral(token, result);
            }
        }
    }
    private static string FormatExponential(double value, Section format, CultureInfo culture)
    {
        // The application shall display a number to the right of 
        // the "E" symbol that corresponds to the number of places that 
        // the decimal point was moved. 

        var baseDigits = 0;
        if (format.Exponential.BeforeDecimal != null)
        {
            baseDigits = GetDigitCount(format.Exponential.BeforeDecimal);
        }

        var exponent = (int)Math.Floor(Math.Log10(Math.Abs(value)));
        var mantissa = value / Math.Pow(10, exponent);

        var shift = Math.Abs(exponent) % baseDigits;
        if (shift > 0)
        {
            if (exponent < 0)
                shift = (baseDigits - shift);

            mantissa *= Math.Pow(10, shift);
            exponent -= shift;
        }

        var result = new StringBuilder();
        FormatNumber(mantissa, format.Exponential.BeforeDecimal, format.Exponential.DecimalSeparator, format.Exponential.AfterDecimal, false, culture, result);

        result.Append(format.Exponential.ExponentialToken[0]);

        if (format.Exponential.ExponentialToken[1] == '+' && exponent >= 0)
        {
            result.Append("+");
        }
        else if (exponent < 0)
        {
            result.Append("-");
        }

        FormatThousands(Math.Abs(exponent).ToString(CultureInfo.InvariantCulture), false, false, format.Exponential.Power, culture, result);
        return result.ToString();
    }
    private static string FormatFraction(double value, Section format, CultureInfo culture)
    {
        int integral = 0;
        int numerator, denominator;

        bool sign = value < 0;

        if (format.Fraction.IntegerPart != null)
        {
            integral = (int)Math.Truncate(value);
            value = Math.Abs(value - integral);
        }

        if (format.Fraction.DenominatorConstant != 0)
        {
            denominator = format.Fraction.DenominatorConstant;
            var rr = Math.Round(value * denominator);
            var b = Math.Floor(rr / denominator);
            numerator = (int)(rr - b * denominator);
        }
        else
        {
            var denominatorDigits = Math.Min(GetDigitCount(format.Fraction.Denominator), 7);
            GetFraction(value, (int)Math.Pow(10, denominatorDigits) - 1, out numerator, out denominator);
        }

        // Don't hide fraction if at least one zero in the numerator format
        var numeratorZeros = GetZeroCount(format.Fraction.Numerator);
        var hideFraction = (format.Fraction.IntegerPart != null && numerator == 0 && numeratorZeros == 0);

        var result = new StringBuilder();

        if (sign)
            result.Append("-");

        // Print integer part with significant zero if fraction part is hidden
        if (format.Fraction.IntegerPart != null)
            FormatThousands(Math.Abs(integral).ToString("F0", CultureInfo.InvariantCulture), false, hideFraction, format.Fraction.IntegerPart, culture, result);

        var numeratorString = Math.Abs(numerator).ToString("F0", CultureInfo.InvariantCulture);
        var denominatorString = denominator.ToString("F0", CultureInfo.InvariantCulture);

        var fraction = new StringBuilder();

        FormatThousands(numeratorString, false, true, format.Fraction.Numerator, culture, fraction);

        fraction.Append("/");

        if (format.Fraction.DenominatorPrefix != null)
            FormatThousands("", false, false, format.Fraction.DenominatorPrefix, culture, fraction);

        if (format.Fraction.DenominatorConstant != 0)
            fraction.Append(format.Fraction.DenominatorConstant.ToString());
        else
            FormatDenominator(denominatorString, format.Fraction.Denominator, fraction);

        if (format.Fraction.DenominatorSuffix != null)
            FormatThousands("", false, false, format.Fraction.DenominatorSuffix, culture, fraction);

        if (hideFraction)
            result.Append(new string(' ', fraction.ToString().Length));
        else
            result.Append(fraction.ToString());

        if (format.Fraction.FractionSuffix != null)
            FormatThousands("", false, false, format.Fraction.FractionSuffix, culture, result);

        return result.ToString();
    }
    // Adapted from ssf.js 'frac()' helper
    private static void GetFraction(double x, int D, out int nom, out int den)
    {
        var sgn = x < 0 ? -1 : 1;
        var B = x * sgn;
        var P_2 = 0.0;
        var P_1 = 1.0;
        var P = 0.0;
        var Q_2 = 1.0;
        var Q_1 = 0.0;
        var Q = 0.0;
        var A = Math.Floor(B);
        while (Q_1 < D)
        {
            A = Math.Floor(B);
            P = A * P_1 + P_2;
            Q = A * Q_1 + Q_2;
            if ((B - A) < 0.00000005) break;
            B = 1 / (B - A);
            P_2 = P_1; P_1 = P;
            Q_2 = Q_1; Q_1 = Q;
        }
        if (Q > D) { if (Q_1 > D) { Q = Q_2; P = P_2; } else { Q = Q_1; P = P_1; } }
        nom = (int)(sgn * P);
        den = (int)Q;
    }
    /// <summary>
    /// Prints left-aligned, left-padded fraction integer denominator.
    /// Assumes tokens contain only placeholders, valueString has fewer or equal number of digits as tokens.
    /// </summary>
    private static void FormatDenominator(string valueString, List<string> tokens, StringBuilder result)
    {
        var formatDigits = GetDigitCount(tokens);
        valueString = valueString.PadLeft(formatDigits, '0');

        bool significant = false;
        var valueIndex = 0;
        for (var tokenIndex = 0; tokenIndex < tokens.Count; ++tokenIndex)
        {
            var token = tokens[tokenIndex];
            char c;
            if (valueIndex < valueString.Length)
            {
                c = GetLeftAlignedValueDigit(token, valueString, valueIndex, significant, out valueIndex);

                if (c != '0')
                    significant = true;
            }
            else
            {
                c = '0';
                significant = false;
            }

            FormatPlaceholder(token, c, significant, result);
        }
    }
    /// <summary>
    /// Returns the first digit from valueString. If the token is '?' 
    /// returns the first significant digit from valueString, or '0' if there are no significant digits.
    /// The out valueIndex parameter contains the offset to the next digit in valueString.
    /// </summary>
    private static char GetLeftAlignedValueDigit(string token, string valueString, int startIndex, bool significant, out int valueIndex)
    {
        char c;
        valueIndex = startIndex;
        if (valueIndex < valueString.Length)
        {
            c = valueString[valueIndex];
            valueIndex++;

            if (c != '0')
                significant = true;

            if (token == "?" && !significant)
            {
                // Eat insignificant zeros to left align denominator
                while (valueIndex < valueString.Length)
                {
                    c = valueString[valueIndex];
                    valueIndex++;

                    if (c != '0')
                    {
                        significant = true;
                        break;
                    }
                }
            }
        }
        else
        {
            c = '0';
            significant = false;
        }

        return c;
    }
    private static void FormatPlaceholder(string token, char c, bool significant, StringBuilder result)
    {
        if (token == "0")
        {
            if (significant)
                result.Append(c);
            else
                result.Append("0");
        }
        else if (token == "#")
        {
            if (significant)
                result.Append(c);
        }
        else if (token == "?")
        {
            if (significant)
                result.Append(c);
            else
                result.Append(" ");
        }
    }
    private static int GetDigitCount(List<string> tokens)
    {
        var counter = 0;
        foreach (var token in tokens)
        {
            if (Token.IsPlaceholder(token))
            {
                counter++;
            }
        }
        return counter;
    }
    private static int GetZeroCount(List<string> tokens)
    {
        var counter = 0;
        foreach (var token in tokens)
        {
            if (token == "0")
                counter++;
        }
        return counter;
    }
    private static void FormatLiteral(string token, StringBuilder result)
    {
        string literal = string.Empty;
        if (token == ",")
        {
            ; // skip commas
        }
        else if (token.Length == 2 && (token[0] == '*' || token[0] == '\\'))
            // TODO: * = repeat to fill cell
            literal = token[1].ToString();
        else if (token.Length == 2 && token[0] == '_')
            literal = " ";
        else if (token.StartsWith("\""))
            literal = token.Substring(1, token.Length - 2);
        else literal = token;
        result.Append(literal);
    }
    private static string ToCompatibleString(object value, IFormatProvider provider)
    {
        return value switch
        {
            double d => d.ToString("G15", provider),
            float f => f.ToString("G7", provider),
            _ => Convert.ToString(value, provider)
        };
    }

    #region others
    static class Parser
    {
        public static List<Section> ParseSections(string formatString, out bool hasSyntaxError)
        {
            var tokenizer = new Tokenizer(formatString);
            var sections = new List<Section>();
            hasSyntaxError = false;
            while (true)
            {
                var section = ParseSection(tokenizer, sections.Count, out var sectionSyntaxError);

                if (sectionSyntaxError)
                    hasSyntaxError = true;

                if (section == null)
                    break;

                sections.Add(section);
            }

            return sections;
        }
        private static Section ParseSection(Tokenizer reader, int index, out bool syntaxError)
        {
            bool hasDateParts = false;
            bool hasDurationParts = false;
            bool hasGeneralPart = false;
            bool hasTextPart = false;
            bool hasPlaceholders = false;
            Condition condition = null;
            Color color = null;
            string token;
            List<string> tokens = new List<string>();

            syntaxError = false;
            while ((token = ReadToken(reader, out syntaxError)) != null)
            {
                if (token == ";")
                    break;

                hasPlaceholders |= Token.IsPlaceholder(token);

                if (Token.IsDatePart(token))
                {
                    hasDateParts |= true;
                    hasDurationParts |= Token.IsDurationPart(token);
                    tokens.Add(token);
                }
                else if (Token.IsGeneral(token))
                {
                    hasGeneralPart |= true;
                    tokens.Add(token);
                }
                else if (token == "@")
                {
                    hasTextPart |= true;
                    tokens.Add(token);
                }
                else if (token.StartsWith("["))
                {
                    // Does not add to tokens. Absolute/elapsed time tokens
                    // also start with '[', but handled as date part above
                    var expression = token.Substring(1, token.Length - 2);
                    if (TryParseCondition(expression, out var parseCondition))
                        condition = parseCondition;
                    else if (TryParseColor(expression, out var parseColor))
                        color = parseColor;
                    else if (TryParseCurrencySymbol(expression, out var parseCurrencySymbol))
                        tokens.Add("\"" + parseCurrencySymbol + "\"");
                }
                else
                {
                    tokens.Add(token);
                }
            }

            if (syntaxError || tokens.Count == 0)
            {
                return null;
            }

            if (
                (hasDateParts && (hasGeneralPart || hasTextPart)) ||
                (hasGeneralPart && (hasDateParts || hasTextPart)) ||
                (hasTextPart && (hasGeneralPart || hasDateParts)))
            {
                // Cannot mix date, general and/or text parts
                syntaxError = true;
                return null;
            }

            SectionType type;
            FractionSection fraction = null;
            ExponentialSection exponential = null;
            DecimalSection number = null;
            List<string> generalTextDateDuration = null;

            if (hasDateParts)
            {
                if (hasDurationParts)
                {
                    type = SectionType.Duration;
                }
                else
                {
                    type = SectionType.Date;
                }

                ParseMilliseconds(tokens, out generalTextDateDuration);
            }
            else if (hasGeneralPart)
            {
                type = SectionType.General;
                generalTextDateDuration = tokens;
            }
            else if (hasTextPart || !hasPlaceholders)
            {
                type = SectionType.Text;
                generalTextDateDuration = tokens;
            }
            else if (FractionSection.TryParse(tokens, out fraction))
            {
                type = SectionType.Fraction;
            }
            else if (ExponentialSection.TryParse(tokens, out exponential))
            {
                type = SectionType.Exponential;
            }
            else if (DecimalSection.TryParse(tokens, out number))
            {
                type = SectionType.Number;
            }
            else
            {
                // Unable to parse format string
                syntaxError = true;
                return null;
            }

            return new Section()
            {
                Type = type,
                SectionIndex = index,
                Color = color,
                Condition = condition,
                Fraction = fraction,
                Exponential = exponential,
                Number = number,
                GeneralTextDateDurationParts = generalTextDateDuration
            };
        }
        /// <summary>
        /// Parses as many placeholders and literals needed to format a number with optional decimals. 
        /// Returns number of tokens parsed, or 0 if the tokens didn't form a number.
        /// </summary>
        internal static int ParseNumberTokens(List<string> tokens, int startPosition, out List<string> beforeDecimal, out bool decimalSeparator, out List<string> afterDecimal)
        {
            beforeDecimal = null;
            afterDecimal = null;
            decimalSeparator = false;

            List<string> remainder = new List<string>();
            var index = 0;
            for (index = 0; index < tokens.Count; ++index)
            {
                var token = tokens[index];
                if (token == "." && beforeDecimal == null)
                {
                    decimalSeparator = true;
                    beforeDecimal = tokens.GetRange(0, index); // TODO: why not remainder? has only valid tokens...

                    remainder = new List<string>();
                }
                else if (Token.IsNumberLiteral(token))
                {
                    remainder.Add(token);
                }
                else if (token.StartsWith("["))
                {
                    // ignore
                }
                else
                {
                    break;
                }
            }

            if (remainder.Count > 0)
            {
                if (beforeDecimal != null)
                {
                    afterDecimal = remainder;
                }
                else
                {
                    beforeDecimal = remainder;
                }
            }

            return index;
        }
        private static void ParseMilliseconds(List<string> tokens, out List<string> result)
        {
            // if tokens form .0 through .000.., combine to single subsecond token
            result = new List<string>();
            for (var i = 0; i < tokens.Count; i++)
            {
                var token = tokens[i];
                if (token == ".")
                {
                    var zeros = 0;
                    while (i + 1 < tokens.Count && tokens[i + 1] == "0")
                    {
                        i++;
                        zeros++;
                    }

                    if (zeros > 0)
                        result.Add("." + new string('0', zeros));
                    else
                        result.Add(".");
                }
                else
                {
                    result.Add(token);
                }
            }
        }
        private static string ReadToken(Tokenizer reader, out bool syntaxError)
        {
            var offset = reader.Position;
            if (
                ReadLiteral(reader) ||
                reader.ReadEnclosed('[', ']') ||

                // Symbols
                reader.ReadOneOf("#?,!&%+-$€£0123456789{}():;/.@ ") ||
                reader.ReadString("e+", true) ||
                reader.ReadString("e-", true) ||
                reader.ReadString("General", true) ||

                // Date
                reader.ReadString("am/pm", true) ||
                reader.ReadString("a/p", true) ||
                reader.ReadOneOrMore('y') ||
                reader.ReadOneOrMore('Y') ||
                reader.ReadOneOrMore('m') ||
                reader.ReadOneOrMore('M') ||
                reader.ReadOneOrMore('d') ||
                reader.ReadOneOrMore('D') ||
                reader.ReadOneOrMore('h') ||
                reader.ReadOneOrMore('H') ||
                reader.ReadOneOrMore('s') ||
                reader.ReadOneOrMore('S') ||
                reader.ReadOneOrMore('g') ||
                reader.ReadOneOrMore('G'))
            {
                syntaxError = false;
                var length = reader.Position - offset;
                return reader.Substring(offset, length);
            }

            syntaxError = reader.Position < reader.Length;
            return null;
        }
        private static bool ReadLiteral(Tokenizer reader)
        {
            if (reader.Peek() == '\\' || reader.Peek() == '*' || reader.Peek() == '_')
            {
                reader.Advance(2);
                return true;
            }
            else if (reader.ReadEnclosed('"', '"'))
                return true;

            return false;
        }
        private static bool TryParseCondition(string token, out Condition result)
        {
            var tokenizer = new Tokenizer(token);

            if (tokenizer.ReadString("<=") ||
                tokenizer.ReadString("<>") ||
                tokenizer.ReadString("<") ||
                tokenizer.ReadString(">=") ||
                tokenizer.ReadString(">") ||
                tokenizer.ReadString("="))
            {
                var conditionPosition = tokenizer.Position;
                var op = tokenizer.Substring(0, conditionPosition);

                if (ReadConditionValue(tokenizer))
                {
                    var valueString = tokenizer.Substring(conditionPosition, tokenizer.Position - conditionPosition);

                    result = new Condition()
                    {
                        Operator = op,
                        Value = double.Parse(valueString, CultureInfo.InvariantCulture)
                    };
                    return true;
                }
            }

            result = null;
            return false;
        }
        private static bool ReadConditionValue(Tokenizer tokenizer)
        {
            // NFPartCondNum = [ASCII-HYPHEN-MINUS] NFPartIntNum [INTL-CHAR-DECIMAL-SEP NFPartIntNum] [NFPartExponential NFPartIntNum]
            tokenizer.ReadString("-");
            while (tokenizer.ReadOneOf("0123456789"))
            {
            }

            if (tokenizer.ReadString("."))
            {
                while (tokenizer.ReadOneOf("0123456789"))
                {
                }
            }

            if (tokenizer.ReadString("e+", true) || tokenizer.ReadString("e-", true))
            {
                if (tokenizer.ReadOneOf("0123456789"))
                {
                    while (tokenizer.ReadOneOf("0123456789"))
                    {
                    }
                }
                else return false;
            }

            return true;
        }
        private static bool TryParseColor(string token, out Color color)
        {
            // TODO: Color1..59
            var tokenizer = new Tokenizer(token);
            if (
                tokenizer.ReadString("black", true) ||
                tokenizer.ReadString("blue", true) ||
                tokenizer.ReadString("cyan", true) ||
                tokenizer.ReadString("green", true) ||
                tokenizer.ReadString("magenta", true) ||
                tokenizer.ReadString("red", true) ||
                tokenizer.ReadString("white", true) ||
                tokenizer.ReadString("yellow", true))
            {
                color = new Color()
                {
                    Value = tokenizer.Substring(0, tokenizer.Position)
                };
                return true;
            }

            color = null;
            return false;
        }
        private static bool TryParseCurrencySymbol(string token, out string currencySymbol)
        {
            if (string.IsNullOrEmpty(token)
                || !token.StartsWith("$"))
            {
                currencySymbol = null;
                return false;
            }

            if (token.Contains("-"))
                currencySymbol = token.Substring(1, token.IndexOf('-') - 1);
            else currencySymbol = token.Substring(1);

            return true;
        }
    }
    static class Token
    {
        public static bool IsExponent(string token)
        {
            return (string.Compare(token, "e+", StringComparison.OrdinalIgnoreCase) == 0)
                || (string.Compare(token, "e-", StringComparison.OrdinalIgnoreCase) == 0);
        }
        public static bool IsLiteral(string token)
        {
            return token.StartsWith("_") ||
                token.StartsWith("\\") ||
                token.StartsWith("\"") ||
                token.StartsWith("*") ||
                token == "," ||
                token == "!" ||
                token == "&" ||
                token == "%" ||
                token == "+" ||
                token == "-" ||
                token == "$" ||
                token == "€" ||
                token == "£" ||
                token == "1" ||
                token == "2" ||
                token == "3" ||
                token == "4" ||
                token == "5" ||
                token == "6" ||
                token == "7" ||
                token == "8" ||
                token == "9" ||
                token == "{" ||
                token == "}" ||
                token == "(" ||
                token == ")" ||
                token == " ";
        }
        public static bool IsNumberLiteral(string token)
            => IsPlaceholder(token) || IsLiteral(token) || token == ".";
        public static bool IsPlaceholder(string token)
            => token == "0" || token == "#" || token == "?";
        public static bool IsGeneral(string token)
            => string.Compare(token, "general", StringComparison.OrdinalIgnoreCase) == 0;
        public static bool IsDatePart(string token)
        {
            return token.StartsWith("y", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("m", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("d", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("s", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("h", StringComparison.OrdinalIgnoreCase) ||
                (token.StartsWith("g", StringComparison.OrdinalIgnoreCase) && !IsGeneral(token)) ||
                string.Compare(token, "am/pm", StringComparison.OrdinalIgnoreCase) == 0 ||
                string.Compare(token, "a/p", StringComparison.OrdinalIgnoreCase) == 0 ||
                IsDurationPart(token);
        }
        public static bool IsDurationPart(string token)
        {
            return token.StartsWith("[h", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("[m", StringComparison.OrdinalIgnoreCase) ||
                token.StartsWith("[s", StringComparison.OrdinalIgnoreCase);
        }
        public static bool IsDigit09(string token) => token == "0" || IsDigit19(token);
        public static bool IsDigit19(string token)
        {
            switch (token)
            {
                case "1":
                case "2":
                case "3":
                case "4":
                case "5":
                case "6":
                case "7":
                case "8":
                case "9": return true;
                default: return false;
            }
        }
    }
    static class Evaluator
    {
        public static Section GetSection(List<Section> sections, object value)
        {
            // Standard format has up to 4 sections:
            // Positive;Negative;Zero;Text
            switch (value)
            {
                case string s:
                    if (sections.Count >= 4)
                        return sections[3];

                    return null;

                case DateTime dt:
                    // TODO: Check date conditions need date helpers and Date1904 knowledge
                    return GetFirstSection(sections, SectionType.Date);

                case TimeSpan ts:
                    return GetNumericSection(sections, ts.TotalDays);

                case double d:
                    return GetNumericSection(sections, d);

                case int i:
                    return GetNumericSection(sections, i);

                case short s:
                    return GetNumericSection(sections, s);

                default:
                    return null;
            }
        }
        public static Section GetFirstSection(List<Section> sections, SectionType type)
        {
            foreach (var section in sections)
                if (section.Type == type)
                    return section;
            return null;
        }
        private static Section GetNumericSection(List<Section> sections, double value)
        {
            // First section applies if 
            // - Has a condition:
            // - There is 1 section, or
            // - There are 2 sections, and the value is 0 or positive, or
            // - There are >2 sections, and the value is positive
            if (sections.Count < 1)
                return null;

            var section0 = sections[0];

            if (section0.Condition != null)
            {
                if (section0.Condition.Evaluate(value))
                    return section0;
            }
            else if (sections.Count == 1 || (sections.Count == 2 && value >= 0) || (sections.Count >= 2 && value > 0))
                return section0;

            if (sections.Count < 2)
                return null;

            var section1 = sections[1];

            // First condition didnt match, or was a negative number. Second condition applies if:
            // - Has a condition, or
            // - Value is negative, or
            // - There are two sections, and the first section had a non-matching condition
            if (section1.Condition != null)
            {
                if (section1.Condition.Evaluate(value))
                    return section1;
            }
            else if (value < 0 || (sections.Count == 2 && section0.Condition != null))
                return section1;

            // Second condition didnt match, or was positive. The following 
            // sections cannot have conditions, always fall back to the third 
            // section (for zero formatting) if specified.
            if (sections.Count < 3)
                return null;

            return sections[2];
        }
    }
    class Condition
    {
        public string Operator { get; set; }
        public double Value { get; set; }
        public bool Evaluate(double lhs)
        {
            switch (Operator)
            {
                case "<":
                    return lhs < Value;
                case "<=":
                    return lhs <= Value;
                case ">":
                    return lhs > Value;
                case ">=":
                    return lhs >= Value;
                case "<>":
                    return lhs != Value;
                case "=":
                    return lhs == Value;
            }
            return false;
        }
    }
    class DecimalSection
    {
        public bool ThousandSeparator { get; set; }
        public double ThousandDivisor { get; set; }
        public double PercentMultiplier { get; set; }
        public List<string> BeforeDecimal { get; set; }
        public bool DecimalSeparator { get; set; }
        public List<string> AfterDecimal { get; set; }
        public static bool TryParse(List<string> tokens, out DecimalSection format)
        {
            if (Parser.ParseNumberTokens(tokens, 0, out var beforeDecimal, out var decimalSeparator, out var afterDecimal) == tokens.Count)
            {
                bool thousandSeparator;
                var divisor = GetTrailingCommasDivisor(tokens, out thousandSeparator);
                var multiplier = GetPercentMultiplier(tokens);

                format = new DecimalSection()
                {
                    BeforeDecimal = beforeDecimal,
                    DecimalSeparator = decimalSeparator,
                    AfterDecimal = afterDecimal,
                    PercentMultiplier = multiplier,
                    ThousandDivisor = divisor,
                    ThousandSeparator = thousandSeparator
                };
                return true;
            }
            format = null;
            return false;
        }
        static double GetPercentMultiplier(List<string> tokens)
        {
            // If there is a percentage literal in the part list, multiply the result by 100
            foreach (var token in tokens)
            {
                if (token == "%")
                    return 100;
            }

            return 1;
        }
        static double GetTrailingCommasDivisor(List<string> tokens, out bool thousandSeparator)
        {
            // This parses all comma literals in the part list:
            // Each comma after the last digit placeholder divides the result by 1000.
            // If there are any other commas, display the result with thousand separators.
            bool hasLastPlaceholder = false;
            var divisor = 1.0;

            for (var j = 0; j < tokens.Count; j++)
            {
                var tokenIndex = tokens.Count - 1 - j;
                var token = tokens[tokenIndex];

                if (!hasLastPlaceholder)
                {
                    if (Token.IsPlaceholder(token))
                    {
                        // Each trailing comma multiplies the divisor by 1000
                        for (var k = tokenIndex + 1; k < tokens.Count; k++)
                        {
                            token = tokens[k];
                            if (token == ",")
                                divisor *= 1000.0;
                            else
                                break;
                        }

                        // Continue scanning backwards from the last digit placeholder, 
                        // but now look for a thousand separator comma
                        hasLastPlaceholder = true;
                    }
                }
                else
                {
                    if (token == ",")
                    {
                        thousandSeparator = true;
                        return divisor;
                    }
                }
            }

            thousandSeparator = false;
            return divisor;
        }
    }
    /// <summary>
    /// Similar to regular .NET DateTime, but also supports 0/1 1900 and 29/2 1900.
    /// </summary>
    class ExcelDateTime
    {
        /// <summary>
        /// The closest .NET DateTime to the specified excel date. 
        /// </summary>
        public DateTime AdjustedDateTime { get; }
        /// <summary>
        /// Number of days to adjust by in post.
        /// </summary>
        public int AdjustDaysPost { get; }
        /// <summary>
        /// Constructs a new ExcelDateTime from a numeric value.
        /// </summary>
        public ExcelDateTime(double numericDate, bool isDate1904)
        {
            if (isDate1904)
            {
                numericDate += 1462.0;
                AdjustedDateTime = new DateTime(DoubleDateToTicks(numericDate), DateTimeKind.Unspecified);
            }
            else
            {
                // internal dates before 30/12/1899 should add two days to get the real date
                // internal dates on 30/12 19899 should add two days, but subtract a day post to get the real date
                // internal dates before 28/2/1900 should add one day to get the real date
                // internal dates on 28/2 1900 should use the same date, but add a day post to get the real date

                var internalDateTime = new DateTime(DoubleDateToTicks(numericDate), DateTimeKind.Unspecified);
                if (internalDateTime < Excel1900ZeroethMinDate)
                {
                    AdjustDaysPost = 0;
                    AdjustedDateTime = internalDateTime.AddDays(2);
                }

                else if (internalDateTime < Excel1900ZeroethMaxDate)
                {
                    AdjustDaysPost = -1;
                    AdjustedDateTime = internalDateTime.AddDays(2);
                }

                else if (internalDateTime < Excel1900LeapMinDate)
                {
                    AdjustDaysPost = 0;
                    AdjustedDateTime = internalDateTime.AddDays(1);
                }

                else if (internalDateTime < Excel1900LeapMaxDate)
                {
                    AdjustDaysPost = 1;
                    AdjustedDateTime = internalDateTime;
                }
                else
                {
                    AdjustDaysPost = 0;
                    AdjustedDateTime = internalDateTime;
                }
            }
        }

        static DateTime Excel1900LeapMinDate = new DateTime(1900, 2, 28);
        static DateTime Excel1900LeapMaxDate = new DateTime(1900, 3, 1);
        static DateTime Excel1900ZeroethMinDate = new DateTime(1899, 12, 30);
        static DateTime Excel1900ZeroethMaxDate = new DateTime(1899, 12, 31);

        /// <summary>
        /// Wraps a regular .NET datetime.
        /// </summary>
        /// <param name="value"></param>
        public ExcelDateTime(DateTime value)
        {
            AdjustedDateTime = value;
            AdjustDaysPost = 0;
        }

        public int Year => AdjustedDateTime.Year;
        public int Month => AdjustedDateTime.Month;
        public int Day => AdjustedDateTime.Day + AdjustDaysPost;
        public int Hour => AdjustedDateTime.Hour;
        public int Minute => AdjustedDateTime.Minute;
        public int Second => AdjustedDateTime.Second;
        public int Millisecond => AdjustedDateTime.Millisecond;
        public DayOfWeek DayOfWeek => AdjustedDateTime.DayOfWeek;

        public string ToString(string numberFormat, CultureInfo culture)
            => AdjustedDateTime.ToString(numberFormat, culture);

        public static bool TryConvert(object value, bool isDate1904, CultureInfo culture, out ExcelDateTime result)
        {
            if (value is double doubleValue)
            {
                result = new ExcelDateTime(doubleValue, isDate1904);
                return true;
            }
            if (value is int intValue)
            {
                result = new ExcelDateTime(intValue, isDate1904);
                return true;
            }
            if (value is short shortValue)
            {
                result = new ExcelDateTime(shortValue, isDate1904);
                return true;
            }
            else if (value is DateTime dateTimeValue)
            {
                result = new ExcelDateTime(dateTimeValue);
                return true;
            }

            result = null;
            return false;
        }

        // From DateTime class to enable OADate in PCL
        // Number of 100ns ticks per time unit
        private const long TicksPerMillisecond = 10000;
        private const long TicksPerSecond = TicksPerMillisecond * 1000;
        private const long TicksPerMinute = TicksPerSecond * 60;
        private const long TicksPerHour = TicksPerMinute * 60;
        private const long TicksPerDay = TicksPerHour * 24;

        private const int MillisPerSecond = 1000;
        private const int MillisPerMinute = MillisPerSecond * 60;
        private const int MillisPerHour = MillisPerMinute * 60;
        private const int MillisPerDay = MillisPerHour * 24;

        // Number of days in a non-leap year
        private const int DaysPerYear = 365;
        // Number of days in 4 years
        private const int DaysPer4Years = DaysPerYear * 4 + 1;
        // Number of days in 100 years
        private const int DaysPer100Years = DaysPer4Years * 25 - 1;
        // Number of days in 400 years
        private const int DaysPer400Years = DaysPer100Years * 4 + 1;
        // Number of days from 1/1/0001 to 12/30/1899
        private const int DaysTo1899 = DaysPer400Years * 4 + DaysPer100Years * 3 - 367;
        private const long DoubleDateOffset = DaysTo1899 * TicksPerDay;

        internal static long DoubleDateToTicks(double value)
        {
            long millis = (long)(value * MillisPerDay + (value >= 0 ? 0.5 : -0.5));

            // The interesting thing here is when you have a value like 12.5 it all positive 12 days and 12 hours from 01/01/1899
            // However if you a value of -12.25 it is minus 12 days but still positive 6 hours, almost as though you meant -11.75 all negative
            // This line below fixes up the millis in the negative case
            if (millis < 0)
                millis -= millis % MillisPerDay * 2;

            millis += DoubleDateOffset / TicksPerMillisecond;
            return millis * TicksPerMillisecond;
        }
    }
    class ExponentialSection
    {
        public List<string> BeforeDecimal { get; set; }
        public bool DecimalSeparator { get; set; }
        public List<string> AfterDecimal { get; set; }
        public string ExponentialToken { get; set; }
        public List<string> Power { get; set; }

        public static bool TryParse(List<string> tokens, out ExponentialSection format)
        {
            format = null;
            string exponentialToken;
            int partCount = Parser.ParseNumberTokens(tokens, 0, out var beforeDecimal, out var decimalSeparator, out var afterDecimal);
            if (partCount == 0)
                return false;

            int position = partCount;
            if (position < tokens.Count && Token.IsExponent(tokens[position]))
            {
                exponentialToken = tokens[position];
                position++;
            }
            else return false;

            format = new ExponentialSection()
            {
                BeforeDecimal = beforeDecimal,
                DecimalSeparator = decimalSeparator,
                AfterDecimal = afterDecimal,
                ExponentialToken = exponentialToken,
                Power = tokens.GetRange(position, tokens.Count - position)
            };
            return true;
        }
    }
    class Tokenizer
    {
        private string formatString;
        private int formatStringPosition = 0;

        public Tokenizer(string fmt)
        {
            formatString = fmt;
        }

        public int Position => formatStringPosition;
        public int Length => formatString?.Length ?? 0;
        public string Substring(int startIndex, int length) => formatString.Substring(startIndex, length);

        public int Peek(int offset = 0)
        {
            if (formatStringPosition + offset >= Length)
                return -1;
            return formatString[formatStringPosition + offset];
        }
        public int PeekUntil(int startOffset, int until)
        {
            int offset = startOffset;
            while (true)
            {
                var c = Peek(offset++);
                if (c == -1)
                    break;
                if (c == until)
                    return offset - startOffset;
            }
            return 0;
        }
        public bool PeekOneOf(int offset, string s)
        {
            foreach (var c in s)
            {
                if (Peek(offset) == c)
                    return true;
            }
            return false;
        }
        public void Advance(int characters = 1)
        {
            formatStringPosition = Math.Min(formatStringPosition + characters, formatString.Length);
        }
        public bool ReadOneOrMore(int c)
        {
            if (Peek() != c)
                return false;

            while (Peek() == c)
                Advance();

            return true;
        }
        public bool ReadOneOf(string s)
        {
            if (PeekOneOf(0, s))
            {
                Advance();
                return true;
            }
            return false;
        }
        public bool ReadString(string s, bool ignoreCase = false)
        {
            if (formatStringPosition + s.Length > Length)
                return false;

            for (var i = 0; i < s.Length; i++)
            {
                var c1 = s[i];
                var c2 = (char)Peek(i);
                if (ignoreCase)
                {
                    if (char.ToLower(c1) != char.ToLower(c2)) return false;
                }
                else
                {
                    if (c1 != c2) return false;
                }
            }

            Advance(s.Length);
            return true;
        }
        public bool ReadEnclosed(char open, char close)
        {
            if (Peek() == open)
            {
                int length = PeekUntil(1, close);
                if (length > 0)
                {
                    Advance(1 + length);
                    return true;
                }
            }
            return false;
        }
    }
    enum SectionType { General, Number, Fraction, Exponential, Date, Duration, Text, }
    class Section
    {
        public int SectionIndex { get; set; }
        public SectionType Type { get; set; }
        public Color Color { get; set; }
        public Condition Condition { get; set; }
        public ExponentialSection Exponential { get; set; }
        public FractionSection Fraction { get; set; }
        public DecimalSection Number { get; set; }
        public List<string> GeneralTextDateDurationParts { get; set; }
    }
    class Color { public string Value { get; set; } }
    class FractionSection
    {
        public List<string> IntegerPart { get; set; }
        public List<string> Numerator { get; set; }
        public List<string> DenominatorPrefix { get; set; }
        public List<string> Denominator { get; set; }
        public int DenominatorConstant { get; set; }
        public List<string> DenominatorSuffix { get; set; }
        public List<string> FractionSuffix { get; set; }
        static public bool TryParse(List<string> tokens, out FractionSection format)
        {
            List<string> numeratorParts = null;
            List<string> denominatorParts = null;

            for (var i = 0; i < tokens.Count; i++)
            {
                var part = tokens[i];
                if (part == "/")
                {
                    numeratorParts = tokens.GetRange(0, i);
                    i++;
                    denominatorParts = tokens.GetRange(i, tokens.Count - i);
                    break;
                }
            }

            if (numeratorParts == null)
            {
                format = null;
                return false;
            }

            GetNumerator(numeratorParts, out var integerPart, out var numeratorPart);

            if (!TryGetDenominator(denominatorParts, out var denominatorPrefix, out var denominatorPart, out var denominatorConstant, out var denominatorSuffix, out var fractionSuffix))
            {
                format = null;
                return false;
            }

            format = new FractionSection()
            {
                IntegerPart = integerPart,
                Numerator = numeratorPart,
                DenominatorPrefix = denominatorPrefix,
                Denominator = denominatorPart,
                DenominatorConstant = denominatorConstant,
                DenominatorSuffix = denominatorSuffix,
                FractionSuffix = fractionSuffix
            };

            return true;
        }
        static void GetNumerator(List<string> tokens, out List<string> integerPart, out List<string> numeratorPart)
        {
            var hasPlaceholder = false;
            var hasSpace = false;
            var hasIntegerPart = false;
            var numeratorIndex = -1;
            var index = tokens.Count - 1;
            while (index >= 0)
            {
                var token = tokens[index];
                if (Token.IsPlaceholder(token))
                {
                    hasPlaceholder = true;

                    if (hasSpace)
                    {
                        hasIntegerPart = true;
                        break;
                    }
                }
                else
                {
                    if (hasPlaceholder && !hasSpace)
                    {
                        // First time we get here marks the end of the integer part
                        hasSpace = true;
                        numeratorIndex = index + 1;
                    }
                }
                index--;
            }

            if (hasIntegerPart)
            {
                integerPart = tokens.GetRange(0, numeratorIndex);
                numeratorPart = tokens.GetRange(numeratorIndex, tokens.Count - numeratorIndex);
            }
            else
            {
                integerPart = null;
                numeratorPart = tokens;
            }
        }
        static bool TryGetDenominator(List<string> tokens, out List<string> denominatorPrefix, out List<string> denominatorPart, out int denominatorConstant, out List<string> denominatorSuffix, out List<string> fractionSuffix)
        {
            var index = 0;
            var hasPlaceholder = false;
            var hasConstant = false;

            var constant = new StringBuilder();

            // Read literals until the first number placeholder or digit
            while (index < tokens.Count)
            {
                var token = tokens[index];
                if (Token.IsPlaceholder(token))
                {
                    hasPlaceholder = true;
                    break;
                }
                else
                if (Token.IsDigit19(token))
                {
                    hasConstant = true;
                    break;
                }
                index++;
            }

            if (!hasPlaceholder && !hasConstant)
            {
                denominatorPrefix = null;
                denominatorPart = null;
                denominatorConstant = 0;
                denominatorSuffix = null;
                fractionSuffix = null;
                return false;
            }

            // The denominator starts here, keep the index
            var denominatorIndex = index;

            // Read placeholders or digits in sequence
            while (index < tokens.Count)
            {
                var token = tokens[index];
                if (hasPlaceholder && Token.IsPlaceholder(token))
                {
                    ; // OK
                }
                else
                if (hasConstant && (Token.IsDigit09(token)))
                {
                    constant.Append(token);
                }
                else
                {
                    break;
                }
                index++;
            }

            // 'index' is now at the first token after the denominator placeholders.
            // The remaining, if anything, is to be treated in one or two parts:
            // Any ultimately terminating literals are considered the "Fraction suffix".
            // Anything between the denominator and the fraction suffix is the "Denominator suffix".
            // Placeholders in the denominator suffix are treated as insignificant zeros.

            // Scan backwards to determine the fraction suffix
            int fractionSuffixIndex = tokens.Count;
            while (fractionSuffixIndex > index)
            {
                var token = tokens[fractionSuffixIndex - 1];
                if (Token.IsPlaceholder(token))
                {
                    break;
                }

                fractionSuffixIndex--;
            }

            // Finally extract the detected token ranges

            if (denominatorIndex > 0)
                denominatorPrefix = tokens.GetRange(0, denominatorIndex);
            else
                denominatorPrefix = null;

            if (hasConstant)
                denominatorConstant = int.Parse(constant.ToString());
            else
                denominatorConstant = 0;

            denominatorPart = tokens.GetRange(denominatorIndex, index - denominatorIndex);

            if (index < fractionSuffixIndex)
                denominatorSuffix = tokens.GetRange(index, fractionSuffixIndex - index);
            else
                denominatorSuffix = null;

            if (fractionSuffixIndex < tokens.Count)
                fractionSuffix = tokens.GetRange(fractionSuffixIndex, tokens.Count - fractionSuffixIndex);
            else
                fractionSuffix = null;

            return true;
        }
    }
    #endregion
}
