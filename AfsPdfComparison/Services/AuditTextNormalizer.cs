using System.Globalization;
using System.Text.RegularExpressions;

namespace AfsPdfComparison.Services;

public static class AuditTextNormalizer
{
    private static readonly Regex NamedDateRegex = new(
        @"(?<!\w)(?:(?<day>\d{1,2})(?:st|nd|rd|th)?\s+(?<month>January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?,?\s+(?<year>\d{4})|(?<monthFirst>January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?\s+(?<dayFirst>\d{1,2})(?:st|nd|rd|th)?[,]?\s+(?<yearFirst>\d{4}))(?!\w)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex CurrencyPrefixRegex = new(
        @"(?<![\p{L}\d])(?:ZAR|USD|EUR|GBP|AUD|CAD|JPY|CHF|R|[$€£¥₹])(?=\s*[\(\-−]?\s*\d)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex CurrencySuffixRegex = new(
        @"(?<=\d)\s*(?:ZAR|USD|EUR|GBP|AUD|CAD|JPY|CHF)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex NegativeParentheticalRegex = new(
        @"\(\s*(?<value>\d{1,3}(?:[\s,]\d{3})*(?:\.\d+)?)\s*\)",
        RegexOptions.Compiled);

    private static readonly Regex NegativePrefixRegex = new(
        @"(?<![\w\d])[-−]\s*(?<value>\d{1,3}(?:[\s,]\d{3})*(?:\.\d+)?)",
        RegexOptions.Compiled);

    private static readonly Regex IsoDateRegex = new(
        @"\b\d{4}-\d{2}-\d{2}\b",
        RegexOptions.Compiled);

    private static readonly Regex NumericTokenRegex = new(
        @"(?<![\p{L}\d])\d+(?:(?:[\s,]\d{3})+)?(?:[.,]\d+)?(?![\p{L}\d])",
        RegexOptions.Compiled);

    public static string NormalizeNamedDates(string text)
    {
        if (string.IsNullOrEmpty(text))
            return text ?? "";

        return NamedDateRegex.Replace(text, match =>
        {
            string day = match.Groups["day"].Success
                ? match.Groups["day"].Value
                : match.Groups["dayFirst"].Value;
            string month = match.Groups["month"].Success
                ? match.Groups["month"].Value
                : match.Groups["monthFirst"].Value;
            string year = match.Groups["year"].Success
                ? match.Groups["year"].Value
                : match.Groups["yearFirst"].Value;

            if (DateTime.TryParse(
                    $"{day} {month.TrimEnd('.')} {year}",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out DateTime date))
            {
                return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }

            return match.Value;
        });
    }

    public static bool HasDifferentMeaningfulMarkers(string first, string second)
    {
        first = NormalizeNamedDates(first);
        second = NormalizeNamedDates(second);

        var firstCurrencies = GetCurrencies(first);
        var secondCurrencies = GetCurrencies(second);
        if (firstCurrencies.Count > 0 && secondCurrencies.Count > 0 &&
            !firstCurrencies.SetEquals(secondCurrencies))
        {
            return true;
        }

        if (first.Contains('%') != second.Contains('%'))
            return true;

        return !GetNegativeValues(first).SequenceEqual(GetNegativeValues(second));
    }

    public static string NormalizeCurrencyFormatting(string text)
    {
        string normalized = CurrencyPrefixRegex.Replace(text ?? "", " ");
        return CurrencySuffixRegex.Replace(normalized, " ");
    }

    public static string NormalizeDecimalFormatting(string text)
    {
        string normalized = Regex.Replace(text ?? "", @"(?<=\d)\.0+(?=\D|$)", "");
        return Regex.Replace(normalized, @"(\d+\.\d*?[1-9])0+(?=\D|$)", "$1");
    }

    public static HashSet<string> ExtractNormalizedNumbers(string text)
    {
        string normalized = NormalizeNamedDates(text);
        var dates = IsoDateRegex.Matches(normalized)
            .Select(match => "date:" + match.Value)
            .ToList();
        normalized = IsoDateRegex.Replace(normalized, " ");
        normalized = NormalizeCurrencyFormatting(normalized);

        var numbers = NumericTokenRegex.Matches(normalized)
            .Select(match => NormalizeNumberToken(match.Value))
            .Where(value => value.Length > 0)
            .ToHashSet(StringComparer.Ordinal);

        numbers.UnionWith(dates);
        return numbers;
    }

    public static bool FinancialValuesEquivalent(string first, string second)
    {
        if (!ExtractNormalizedNumbers(first).SetEquals(ExtractNormalizedNumbers(second)))
            return false;

        if (first.Contains('%') != second.Contains('%'))
            return false;

        if (!GetNegativeValues(first).SequenceEqual(GetNegativeValues(second)))
            return false;

        var firstCurrencies = GetCurrencies(first);
        var secondCurrencies = GetCurrencies(second);
        return firstCurrencies.Count == 0 || secondCurrencies.Count == 0 ||
               firstCurrencies.SetEquals(secondCurrencies);
    }

    public static string NormalizeFinancialValue(string value)
    {
        var numbers = ExtractNormalizedNumbers(value);
        string currency = string.Join(",", GetCurrencies(value).OrderBy(item => item, StringComparer.Ordinal));
        string negative = GetNegativeValues(value).Count > 0 ? "-" : "";
        string percent = value.Contains('%') ? "%" : "";
        return $"{negative}{string.Join(",", numbers.OrderBy(item => item, StringComparer.Ordinal))}|{currency}|{percent}";
    }

    private static HashSet<string> GetCurrencies(string text)
    {
        var currencies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (Match match in CurrencyPrefixRegex.Matches(text))
            currencies.Add(NormalizeCurrency(match.Value));

        foreach (Match match in CurrencySuffixRegex.Matches(text))
            currencies.Add(NormalizeCurrency(match.Value));

        return currencies;
    }

    private static string NormalizeCurrency(string value)
    {
        string currency = value.Trim().ToUpperInvariant();
        return currency switch
        {
            "R" or "ZAR" => "ZAR",
            "$" or "USD" => "USD",
            "€" or "EUR" => "EUR",
            "£" or "GBP" => "GBP",
            "¥" or "JPY" => "JPY",
            _ => currency,
        };
    }

    private static List<string> GetNegativeValues(string text)
    {
        var values = new List<string>();
        values.AddRange(NegativeParentheticalRegex.Matches(text)
            .Select(match => "-" + NormalizeNumberToken(match.Groups["value"].Value)));
        values.AddRange(NegativePrefixRegex.Matches(text)
            .Select(match => "-" + NormalizeNumberToken(match.Groups["value"].Value)));
        values.Sort(StringComparer.Ordinal);
        return values;
    }

    private static string NormalizeNumberToken(string value)
    {
        string normalized = Regex.Replace(value, @"\s", "");
        int comma = normalized.LastIndexOf(',');
        int dot = normalized.LastIndexOf('.');
        if (comma >= 0 && dot >= 0)
        {
            if (comma > dot)
                normalized = normalized.Replace(".", "").Replace(',', '.');
            else
                normalized = normalized.Replace(",", "");
        }
        else if (comma >= 0)
        {
            int digitsAfterLastComma = normalized.Length - comma - 1;
            bool thousandsGrouping = digitsAfterLastComma == 3 &&
                                     Regex.IsMatch(normalized, @"^\d{1,3}(?:,\d{3})+$");
            if (thousandsGrouping)
            {
                normalized = normalized.Replace(",", "");
            }
            else if (normalized.IndexOf(',') != comma)
            {
                normalized = normalized[..comma].Replace(",", "") + "." + normalized[(comma + 1)..];
            }
            else
            {
                normalized = normalized.Replace(',', '.');
            }
        }

        if (normalized.Contains('.'))
            normalized = normalized.TrimEnd('0').TrimEnd('.');
        return normalized;
    }
}
