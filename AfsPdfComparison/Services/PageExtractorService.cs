using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using AfsPdfComparison.Models;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace AfsPdfComparison.Services;

public class PageExtractorService
{
    private const int MinReadableChars = 100;

    private static readonly Regex YearRegex = new(@"(?:19|20)\d{2}", RegexOptions.Compiled);
    private static readonly Regex ContextYearRegex = new(
        @"\b(?:for\s+the\s+year\s+ended|year\s+ended|year\s+ending|financial\s+year\s+ended|financial\s+statements\s+for\s+the\s+year\s+ended)\b.{0,40}?\b((?:19|20)\d{2})\b",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline);
    private static readonly Regex DateYearRegex = new(
        @"\b\d{1,2}\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*\s+((?:19|20)\d{2})\b",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex SpaceRegex = new(@"\s+", RegexOptions.Compiled);
    private static readonly Regex InvisibleRegex =
        new(@"[\p{C}\u00AD\u200B-\u200F\uFEFF\uFFF0-\uFFFD]", RegexOptions.Compiled);

    private static readonly string[] PopplerBinCandidates =
    {
        @"C:\Program Files\poppler\Library\bin",
        @"C:\Program Files\poppler-24.02.0\Library\bin",
        @"C:\Program Files (x86)\poppler\bin",
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            @"Microsoft\WinGet\Packages\oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe\poppler-25.07.0\Library\bin"),
    };

    private static readonly string[] TesseractCandidates =
    {
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            @"Programs\Tesseract-OCR\tesseract.exe"),
        @"C:\Program Files\Tesseract-OCR\tesseract.exe",
    };

    private readonly AiVisionOcrService _aiVisionOcr;

    public PageExtractorService(AiVisionOcrService aiVisionOcr)
    {
        _aiVisionOcr = aiVisionOcr;
    }

    public PdfReport Extract(string pdfPath, string filename) =>
        ExtractAsync(pdfPath, filename).GetAwaiter().GetResult();

    public async Task<PdfReport> ExtractAsync(
        string pdfPath,
        string filename,
        CancellationToken cancellationToken = default)
    {
        var pages = TryExtractWithPdfPig(pdfPath, out int pageCount);
        string docType = ReadableCharCount(pages) >= MinReadableChars ? "digital" : string.Empty;

        string? toolFriendlyPath = null;
        bool deleteToolFriendlyPath = false;
        bool repairedHeader = false;

        try
        {
            if (docType.Length == 0)
            {
                toolFriendlyPath = EnsureToolFriendlyPdfCopy(pdfPath, out deleteToolFriendlyPath, out repairedHeader);
                string sourcePath = toolFriendlyPath ?? pdfPath;

                var repairedPages = TryExtractWithPdfPig(sourcePath, out int repairedPageCount);
                if (repairedPageCount > pageCount)
                    pageCount = repairedPageCount;

                if (ReadableCharCount(repairedPages) > ReadableCharCount(pages))
                    pages = repairedPages;

                if (ReadableCharCount(pages) >= MinReadableChars)
                {
                    docType = repairedHeader ? "digital (header repaired)" : "digital (recovered)";
                }
                else
                {
                    if (pageCount == 0)
                        pageCount = TryGetPageCountWithPdfInfo(sourcePath) ?? 0;

                    var popplerPages = TryExtractWithPopplerText(sourcePath, pageCount, out string popplerErrors);
                    if (ReadableCharCount(popplerPages) >= MinReadableChars)
                    {
                        pages = popplerPages;
                        docType = repairedHeader ? "digital (header repaired)" : "digital (recovered)";
                    }
                    else
                    {
                        bool ocrAvailable = FindExecutable("tesseract.exe", TesseractCandidates) != null;
                        bool renderable = pageCount > 0 && CanRenderFirstPage(sourcePath);
                        var ocrPages = TryExtractWithOcr(sourcePath, pageCount);
                        if (ReadableCharCount(ocrPages) >= MinReadableChars)
                        {
                            pages = ocrPages;
                            docType = "scanned (OCR)";
                        }
                        else
                        {
                            var aiOcrAttempt = await _aiVisionOcr.TryExtractPagesAsync(
                                sourcePath,
                                pageCount,
                                cancellationToken);
                            var aiOcrPages = aiOcrAttempt.Pages;

                            if (ReadableCharCount(aiOcrPages) >= MinReadableChars)
                            {
                                pages = aiOcrPages;
                                docType = "scanned (AI OCR)";
                            }
                            else
                            {
                                bool invalidStreams = HasInvalidStreamErrors(popplerErrors);
                                docType =
                                    invalidStreams || LooksCorrupted(pdfPath)
                                        ? (renderable && !ocrAvailable && aiOcrAttempt.Status == "missing_key"
                                            ? "corrupted PDF (invalid streams, AI OCR key missing)"
                                            : renderable && !ocrAvailable && aiOcrAttempt.Status == "disabled"
                                                ? "corrupted PDF (invalid streams, AI OCR disabled)"
                                                : renderable && !ocrAvailable && aiOcrAttempt.Status == "error"
                                                    ? "corrupted PDF (invalid streams, AI OCR failed)"
                                                    : renderable && !ocrAvailable
                                                        ? "corrupted PDF (invalid streams, OCR unavailable)"
                                                        : "corrupted PDF (invalid streams)")
                                        : (renderable && !ocrAvailable && aiOcrAttempt.Status == "missing_key"
                                            ? "scanned (AI OCR key missing)"
                                            : renderable && !ocrAvailable && aiOcrAttempt.Status == "disabled"
                                                ? "scanned (AI OCR disabled)"
                                                : renderable && !ocrAvailable && aiOcrAttempt.Status == "error"
                                                    ? "scanned (AI OCR failed)"
                                                    : renderable && !ocrAvailable
                                                        ? "scanned (OCR unavailable)"
                                                        : "scanned (no OCR)");

                                // Prefer whichever fallback produced the most text, even if it
                                // was not enough to be comparable.
                                if (ReadableCharCount(popplerPages) > ReadableCharCount(pages))
                                    pages = popplerPages;
                                if (ReadableCharCount(ocrPages) > ReadableCharCount(pages))
                                    pages = ocrPages;
                                if (ReadableCharCount(aiOcrPages) > ReadableCharCount(pages))
                                    pages = aiOcrPages;
                            }
                        }
                    }
                }
            }
        }
        finally
        {
            if (deleteToolFriendlyPath && !string.IsNullOrWhiteSpace(toolFriendlyPath))
            {
                try { File.Delete(toolFriendlyPath); } catch { }
            }
        }

        if (pages.Count == 0 && pageCount > 0)
            pages = Enumerable.Repeat(string.Empty, pageCount).ToList();

        if (ReadableCharCount(pages) >= MinReadableChars &&
            (docType.StartsWith("digital", StringComparison.OrdinalIgnoreCase) ||
             docType.Contains("OCR", StringComparison.OrdinalIgnoreCase)))
        {
            pages = StripRunningHeadersAndFooters(pages);
        }

        string fullText = string.Join("\n", pages);
        var yearsFound = YearRegex.Matches(fullText)
            .Select(m => int.Parse(m.Value))
            .Distinct()
            .OrderBy(y => y)
            .ToList();
        int? primaryYear = DeterminePrimaryYear(fullText, pages, yearsFound);

        return new PdfReport
        {
            Filename = filename,
            DocType = docType,
            Pages = pages,
            FullText = fullText,
            PageCount = pages.Count > 0 ? pages.Count : pageCount,
            Years = yearsFound,
            PrimaryYear = primaryYear,
            PdfPath = pdfPath,
        };
    }

    private static List<string> TryExtractWithPdfPig(string pdfPath, out int pageCount)
    {
        pageCount = 0;
        var pages = new List<string>();

        try
        {
            using var document = PdfDocument.Open(pdfPath);
            foreach (var page in document.GetPages())
            {
                pages.Add(ExtractPageText(page));
                pageCount++;
            }
        }
        catch
        {
            // Fallback extractors below handle malformed files more gracefully.
        }

        return pages;
    }

    private static List<string> TryExtractWithPopplerText(string pdfPath, int expectedPageCount, out string errorOutput)
    {
        errorOutput = string.Empty;
        string? pdftotext = FindExecutable("pdftotext.exe", PopplerBinCandidates);
        if (pdftotext == null || !File.Exists(pdfPath))
            return new();

        string outPath = Path.Combine(Path.GetTempPath(), "afs_extract_" + Guid.NewGuid().ToString("N") + ".txt");
        try
        {
            var psi = new ProcessStartInfo(pdftotext)
            {
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            psi.ArgumentList.Add("-layout");
            psi.ArgumentList.Add(pdfPath);
            psi.ArgumentList.Add(outPath);

            using var proc = Process.Start(psi);
            if (proc == null)
                return new();

            if (!proc.WaitForExit(30_000))
            {
                try { proc.Kill(true); } catch { }
                return new();
            }

            errorOutput = proc.StandardError.ReadToEnd();

            if (!File.Exists(outPath))
                return new();

            string raw = File.ReadAllText(outPath);
            return SplitExtractedPages(raw, expectedPageCount);
        }
        catch
        {
            return new();
        }
        finally
        {
            try { if (File.Exists(outPath)) File.Delete(outPath); } catch { }
        }
    }

    private static List<string> TryExtractWithOcr(string pdfPath, int expectedPageCount)
    {
        string? pdftoppm = FindExecutable("pdftoppm.exe", PopplerBinCandidates);
        string? tesseract = FindExecutable("tesseract.exe", TesseractCandidates);
        if (pdftoppm == null || tesseract == null || expectedPageCount <= 0 || !File.Exists(pdfPath))
            return new();

        var pages = new List<string>(expectedPageCount);
        string tempDir = Path.Combine(Path.GetTempPath(), "afs_ocr_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            for (int pageNum = 1; pageNum <= expectedPageCount; pageNum++)
            {
                string prefix = Path.Combine(tempDir, "page");
                string imagePath = Path.Combine(tempDir, $"page-{pageNum:D2}.png");

                try
                {
                    var renderPsi = new ProcessStartInfo(pdftoppm)
                    {
                        RedirectStandardError = true,
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                    };
                    renderPsi.ArgumentList.Add("-png");
                    renderPsi.ArgumentList.Add("-r");
                    renderPsi.ArgumentList.Add("220");
                    renderPsi.ArgumentList.Add("-f");
                    renderPsi.ArgumentList.Add(pageNum.ToString());
                    renderPsi.ArgumentList.Add("-l");
                    renderPsi.ArgumentList.Add(pageNum.ToString());
                    renderPsi.ArgumentList.Add(pdfPath);
                    renderPsi.ArgumentList.Add(prefix);

                    using var renderProc = Process.Start(renderPsi);
                    if (renderProc == null)
                    {
                        pages.Add(string.Empty);
                        continue;
                    }

                    if (!renderProc.WaitForExit(30_000))
                    {
                        try { renderProc.Kill(true); } catch { }
                        pages.Add(string.Empty);
                        continue;
                    }

                    if (!File.Exists(imagePath))
                    {
                        pages.Add(string.Empty);
                        continue;
                    }

                    var ocrPsi = new ProcessStartInfo(tesseract)
                    {
                        RedirectStandardError = true,
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                    };
                    ocrPsi.ArgumentList.Add(imagePath);
                    ocrPsi.ArgumentList.Add("stdout");
                    ocrPsi.ArgumentList.Add("--psm");
                    ocrPsi.ArgumentList.Add("6");

                    using var ocrProc = Process.Start(ocrPsi);
                    if (ocrProc == null)
                    {
                        pages.Add(string.Empty);
                        continue;
                    }

                    string pageText = ocrProc.StandardOutput.ReadToEnd();
                    if (!ocrProc.WaitForExit(30_000))
                    {
                        try { ocrProc.Kill(true); } catch { }
                        pages.Add(string.Empty);
                        continue;
                    }

                    pages.Add(NormalizeOcrPage(pageText));
                }
                catch
                {
                    pages.Add(string.Empty);
                }
            }
        }
        finally
        {
            try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); } catch { }
        }

        return pages;
    }

    private static string NormalizeOcrPage(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        var sb = new StringBuilder();
        foreach (string rawLine in text.Replace("\r", "").Split('\n'))
        {
            string line = SpaceRegex.Replace(rawLine.Trim(), " ");
            if (line.Length > 0)
                sb.Append(line).Append('\n');
        }

        return sb.ToString();
    }

    private static List<string> SplitExtractedPages(string rawText, int expectedPageCount)
    {
        if (string.IsNullOrEmpty(rawText))
            return expectedPageCount > 0
                ? Enumerable.Repeat(string.Empty, expectedPageCount).ToList()
                : new();

        var pages = rawText
            .Replace("\r", "")
            .Split('\f')
            .ToList();

        while (pages.Count > 0 && pages[^1].Length == 0)
            pages.RemoveAt(pages.Count - 1);

        if (expectedPageCount > 0)
        {
            while (pages.Count < expectedPageCount)
                pages.Add(string.Empty);

            if (pages.Count > expectedPageCount)
                pages = pages.Take(expectedPageCount).ToList();
        }

        for (int i = 0; i < pages.Count; i++)
        {
            if (pages[i].Length > 0 && !pages[i].EndsWith('\n'))
                pages[i] += "\n";
        }

        return pages;
    }

    private static int ReadableCharCount(IEnumerable<string> pages) => pages.Sum(p => p?.Length ?? 0);

    private static int? TryGetPageCountWithPdfInfo(string pdfPath)
    {
        string? pdfinfo = FindExecutable("pdfinfo.exe", PopplerBinCandidates);
        if (pdfinfo == null || !File.Exists(pdfPath))
            return null;

        try
        {
            var psi = new ProcessStartInfo(pdfinfo)
            {
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            psi.ArgumentList.Add(pdfPath);

            using var proc = Process.Start(psi);
            if (proc == null)
                return null;

            if (!proc.WaitForExit(15_000))
            {
                try { proc.Kill(true); } catch { }
                return null;
            }

            string output = proc.StandardOutput.ReadToEnd();
            foreach (string line in output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (!line.StartsWith("Pages:", StringComparison.OrdinalIgnoreCase))
                    continue;

                string value = line["Pages:".Length..].Trim();
                if (int.TryParse(value, out int pages))
                    return pages;
            }
        }
        catch
        {
            // Best effort only.
        }

        return null;
    }

    private static bool CanRenderFirstPage(string pdfPath)
    {
        string? pdftoppm = FindExecutable("pdftoppm.exe", PopplerBinCandidates);
        if (pdftoppm == null || !File.Exists(pdfPath))
            return false;

        string tempDir = Path.Combine(Path.GetTempPath(), "afs_render_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            string prefix = Path.Combine(tempDir, "page");
            var psi = new ProcessStartInfo(pdftoppm)
            {
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            psi.ArgumentList.Add("-png");
            psi.ArgumentList.Add("-f");
            psi.ArgumentList.Add("1");
            psi.ArgumentList.Add("-l");
            psi.ArgumentList.Add("1");
            psi.ArgumentList.Add(pdfPath);
            psi.ArgumentList.Add(prefix);

            using var proc = Process.Start(psi);
            if (proc == null)
                return false;

            if (!proc.WaitForExit(15_000))
            {
                try { proc.Kill(true); } catch { }
                return false;
            }

            return Directory.GetFiles(tempDir, "page-*.png").Length > 0;
        }
        catch
        {
            return false;
        }
        finally
        {
            try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); } catch { }
        }
    }

    private static bool HasInvalidStreamErrors(string errorOutput) =>
        !string.IsNullOrWhiteSpace(errorOutput) &&
        (errorOutput.Contains("Bad FCHECK", StringComparison.OrdinalIgnoreCase) ||
         errorOutput.Contains("Syntax Error", StringComparison.OrdinalIgnoreCase));

    private static bool LooksCorrupted(string pdfPath)
    {
        try
        {
            byte[] bytes = File.ReadAllBytes(pdfPath);
            if (bytes.Length == 0)
                return false;

            string head = Encoding.ASCII.GetString(bytes, 0, Math.Min(bytes.Length, 16));
            bool missingPdfSigil = head.StartsWith("PDF-", StringComparison.Ordinal);

            int replacementTriples = 0;
            for (int i = 0; i < bytes.Length - 2; i++)
            {
                if (bytes[i] == 0xEF && bytes[i + 1] == 0xBF && bytes[i + 2] == 0xBD)
                    replacementTriples++;
            }

            return missingPdfSigil || replacementTriples > 100;
        }
        catch
        {
            return false;
        }
    }

    private static string? EnsureToolFriendlyPdfCopy(string pdfPath, out bool deleteWhenDone, out bool repairedHeader)
    {
        deleteWhenDone = false;
        repairedHeader = false;

        try
        {
            byte[] bytes = File.ReadAllBytes(pdfPath);
            if (bytes.Length >= 5 &&
                bytes[0] == (byte)'P' &&
                bytes[1] == (byte)'D' &&
                bytes[2] == (byte)'F' &&
                bytes[3] == (byte)'-')
            {
                string tempPath = Path.Combine(Path.GetTempPath(), "afs_pdf_" + Guid.NewGuid().ToString("N") + ".pdf");
                byte[] fixedBytes = new byte[bytes.Length + 1];
                fixedBytes[0] = 0x25; // %
                Array.Copy(bytes, 0, fixedBytes, 1, bytes.Length);
                File.WriteAllBytes(tempPath, fixedBytes);
                deleteWhenDone = true;
                repairedHeader = true;
                return tempPath;
            }
        }
        catch
        {
            // Best effort only.
        }

        return null;
    }

    private static string? FindExecutable(string fileName, IEnumerable<string> directCandidates)
    {
        string? fromPath = FindExecutableOnPath(fileName);
        if (!string.IsNullOrWhiteSpace(fromPath))
            return fromPath;

        foreach (string candidate in directCandidates)
        {
            if (string.IsNullOrWhiteSpace(candidate))
                continue;

            if (candidate.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
            {
                if (File.Exists(candidate))
                    return candidate;
                continue;
            }

            string fullPath = Path.Combine(candidate, fileName);
            if (File.Exists(fullPath))
                return fullPath;
        }

        return null;
    }

    private static string? FindExecutableOnPath(string fileName)
    {
        string? pathValue = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrWhiteSpace(pathValue))
            return null;

        foreach (string dir in pathValue.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            try
            {
                string fullPath = Path.Combine(dir.Trim(), fileName);
                if (File.Exists(fullPath))
                    return fullPath;
            }
            catch
            {
                // Ignore invalid PATH segments.
            }
        }

        return null;
    }

    private static int? DeterminePrimaryYear(string fullText, List<string> pages, List<int> yearsFound)
    {
        if (yearsFound.Count == 0)
            return null;

        var weighted = new Dictionary<int, double>();

        void AddScore(int year, double score)
        {
            if (!weighted.ContainsKey(year))
                weighted[year] = 0;

            weighted[year] += score;
        }

        foreach (Match match in YearRegex.Matches(fullText))
        {
            if (int.TryParse(match.Value, out int year))
                AddScore(year, 1.0);
        }

        foreach (Match match in ContextYearRegex.Matches(fullText))
        {
            if (int.TryParse(match.Groups[1].Value, out int year))
                AddScore(year, 8.0);
        }

        foreach (Match match in DateYearRegex.Matches(fullText))
        {
            if (int.TryParse(match.Groups[1].Value, out int year))
                AddScore(year, 4.0);
        }

        foreach (string page in pages.Take(3))
        {
            foreach (Match match in YearRegex.Matches(page))
            {
                if (int.TryParse(match.Value, out int year))
                    AddScore(year, 0.5);
            }
        }

        return weighted
            .OrderByDescending(kv => kv.Value)
            .ThenByDescending(kv => kv.Key)
            .Select(kv => (int?)kv.Key)
            .FirstOrDefault();
    }

    private static string StripInvisible(string s)
    {
        if (string.IsNullOrEmpty(s))
            return s;

        return InvisibleRegex.Replace(s.Normalize(NormalizationForm.FormKD), "");
    }

    private static string LineNorm(string s)
    {
        if (string.IsNullOrWhiteSpace(s))
            return string.Empty;

        return SpaceRegex.Replace(StripInvisible(s).Trim(), " ").ToLowerInvariant();
    }

    private static int WordCount(string s)
    {
        string norm = LineNorm(s);
        return norm.Length == 0 ? 0 : norm.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length;
    }

    private static bool IsBoundary(char c) => !char.IsLetterOrDigit(c);

    private static bool ContainsWholeLine(string haystack, string needle)
    {
        if (string.IsNullOrWhiteSpace(haystack) || string.IsNullOrWhiteSpace(needle))
            return false;

        int start = 0;
        while (start < haystack.Length)
        {
            int idx = haystack.IndexOf(needle, start, StringComparison.Ordinal);
            if (idx < 0)
                return false;

            int end = idx + needle.Length;
            bool leftOk = idx == 0 || IsBoundary(haystack[idx - 1]);
            bool rightOk = end >= haystack.Length || IsBoundary(haystack[end]);
            if (leftOk && rightOk)
                return true;

            start = idx + 1;
        }

        return false;
    }

    private static bool ShouldMergeLines(string prev, string curr)
    {
        if (string.IsNullOrWhiteSpace(prev) || string.IsNullOrWhiteSpace(curr))
            return false;

        string current = curr.TrimStart();
        char last = prev[^1];
        bool prevEndsSentence = last is '.' or '!' or '?' or ';' or ':';
        bool startsLower = current.Length > 0 && char.IsLower(current[0]);
        bool startsParenthetical = current.StartsWith("(") || current.StartsWith("[");

        if (prev.Length < 3)
            return false;

        if (startsParenthetical)
            return true;

        return !prevEndsSentence && startsLower && prev.Length >= 15;
    }

    private static bool IsRecurringEdgeCandidate(string line)
    {
        string norm = LineNorm(line);
        if (norm.Length < 8)
            return false;
        if (WordCount(norm) > 14)
            return false;

        int letters = norm.Count(char.IsLetter);
        int digits = norm.Count(char.IsDigit);
        if (letters < 4)
            return false;
        if (digits > letters)
            return false;

        return true;
    }

    private static List<string> StripRunningHeadersAndFooters(List<string> pages)
    {
        if (pages.Count < 4)
            return pages;

        var pageLines = pages
            .Select(p => p.Split('\n', StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim())
                .Where(l => l.Length > 0)
                .ToList())
            .ToList();

        int threshold = Math.Max(4, pages.Count / 8);
        var firstCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        var lastCounts = new Dictionary<string, int>(StringComparer.Ordinal);

        foreach (var lines in pageLines)
        {
            foreach (var line in lines.Take(4).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (!IsRecurringEdgeCandidate(line))
                    continue;

                string norm = LineNorm(line);
                firstCounts[norm] = firstCounts.TryGetValue(norm, out int count) ? count + 1 : 1;
            }

            foreach (var line in lines.Skip(Math.Max(0, lines.Count - 2)).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (!IsRecurringEdgeCandidate(line))
                    continue;

                string norm = LineNorm(line);
                lastCounts[norm] = lastCounts.TryGetValue(norm, out int count) ? count + 1 : 1;
            }
        }

        var recurringHeaders = firstCounts
            .Where(kv => kv.Value >= threshold)
            .Select(kv => kv.Key)
            .ToHashSet(StringComparer.Ordinal);

        var recurringFooters = lastCounts
            .Where(kv => kv.Value >= threshold)
            .Select(kv => kv.Key)
            .ToHashSet(StringComparer.Ordinal);

        if (recurringHeaders.Count == 0 && recurringFooters.Count == 0)
            return pages;

        var cleanedPages = new List<string>(pageLines.Count);
        foreach (var lines in pageLines)
        {
            var cleaned = new List<string>(lines);

            for (int i = Math.Min(3, cleaned.Count - 1); i >= 0; i--)
            {
                if (recurringHeaders.Contains(LineNorm(cleaned[i])))
                    cleaned.RemoveAt(i);
            }

            for (int i = cleaned.Count - 1; i >= Math.Max(0, cleaned.Count - 2); i--)
            {
                if (recurringFooters.Contains(LineNorm(cleaned[i])))
                    cleaned.RemoveAt(i);
            }

            cleanedPages.Add(cleaned.Count == 0 ? string.Empty : string.Join("\n", cleaned) + "\n");
        }

        return cleanedPages;
    }

    private static string ExtractPageText(Page page)
    {
        try
        {
            double pageHeight = page.Height;
            var lineGroups = page.GetWords()
                .GroupBy(w => (int)(Math.Round((pageHeight - w.BoundingBox.Top) / 5.0) * 5))
                .OrderBy(g => g.Key);

            var rawLines = new List<string>();
            foreach (var lineGroup in lineGroups)
            {
                var words = lineGroup
                    .OrderBy(w => w.BoundingBox.Left)
                    .Select(w => StripInvisible(w.Text ?? string.Empty).Trim())
                    .Where(w => w.Length > 0);

                string line = string.Join(" ", words).Trim();
                if (line.Length > 0)
                    rawLines.Add(line);
            }

            var mergedLines = new List<string>(rawLines.Count);
            foreach (string line in rawLines)
            {
                if (mergedLines.Count > 0 &&
                    string.Equals(mergedLines[^1], line, StringComparison.OrdinalIgnoreCase))
                    continue;

                bool merged = false;
                if (mergedLines.Count > 0)
                {
                    string prev = mergedLines[^1];
                    if (ShouldMergeLines(prev, line))
                    {
                        mergedLines[^1] = prev + " " + line;
                        merged = true;
                    }
                }

                if (!merged)
                    mergedLines.Add(line);
            }

            var mergedNorms = mergedLines.Select(LineNorm).ToList();
            var cleanedLines = new List<string>(mergedLines.Count);
            for (int i = 0; i < mergedLines.Count; i++)
            {
                string currentNorm = mergedNorms[i];
                if (currentNorm.Length == 0)
                    continue;

                bool isOrphanFragment = WordCount(currentNorm) <= 4 &&
                    mergedNorms.Where((other, idx) => idx != i && other.Length > currentNorm.Length)
                        .Any(other => ContainsWholeLine(other, currentNorm));

                if (!isOrphanFragment)
                    cleanedLines.Add(mergedLines[i]);
            }

            var sb = new StringBuilder();
            foreach (string line in cleanedLines)
                sb.Append(line).Append('\n');

            return sb.ToString();
        }
        catch
        {
            return string.Empty;
        }
    }
}
