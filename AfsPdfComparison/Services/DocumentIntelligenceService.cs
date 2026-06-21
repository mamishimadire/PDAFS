using System.Text;
using System.Text.RegularExpressions;
using AfsPdfComparison.Models;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace AfsPdfComparison.Services;

public class RichLine
{
    public string Text { get; set; } = "";
    public double FontSize { get; set; }
    public double XPosition { get; set; }
    public double YPosition { get; set; }
    public bool IsBold { get; set; }
    public bool IsAllCaps { get; set; }
    public int WordCount { get; set; }
    public bool HasNumbers { get; set; }
    public int PageNumber { get; set; }
}

public class DiscoveredHeading
{
    public string Text { get; set; } = "";
    public int PageNumber { get; set; }
    public double FontSize { get; set; }
    public double HeadingScore { get; set; }
}

public class DiscoveredSection
{
    public string Title { get; set; } = "";
    public int StartPage { get; set; }
    public int EndPage { get; set; }
    public int PageCount => EndPage - StartPage + 1;
    public List<string> PageTexts { get; set; } = new();
    public string FullText => string.Join("\n", PageTexts);
    public double HeadingFontSize { get; set; }
    public HashSet<string> ContentWords { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public List<string> Numbers { get; set; } = new();
}

public class SectionComparisonResult
{
    public DiscoveredSection? Section1 { get; set; }
    public DiscoveredSection? Section2 { get; set; }
    public string ChangeKind { get; set; } = "";
    public double ContentSim { get; set; }
    public double BoilerplateSim { get; set; }
    public string Description { get; set; } = "";
    public bool IsMaterial { get; set; }
    public List<LineDiffResult> LineDiffs { get; set; } = new();
}

public class DocumentComparisonReport
{
    public List<SectionComparisonResult> Sections { get; set; } = new();
    public string Summary { get; set; } = "";
    public int TotalPages1 { get; set; }
    public int TotalPages2 { get; set; }
}

public class RichPageExtractor
{
    private static readonly Regex InvisRe =
        new(@"[\p{C}\u00AD\u200B-\u200F\uFEFF\uFFF0-\uFFFD]", RegexOptions.Compiled);

    private static string Strip(string s) =>
        string.IsNullOrEmpty(s) ? s : InvisRe.Replace(s.Normalize(NormalizationForm.FormKD), "").Trim();

    public List<List<RichLine>> ExtractRichPages(string pdfPath)
    {
        var result = new List<List<RichLine>>();
        using var doc = PdfDocument.Open(pdfPath);
        foreach (var page in doc.GetPages())
            result.Add(ExtractRichPage(page));
        return result;
    }

    private static List<RichLine> ExtractRichPage(Page page)
    {
        double pageH = page.Height;
        var lines = new List<RichLine>();

        var groups = page.GetWords()
            .GroupBy(w => (int)(Math.Round((pageH - w.BoundingBox.Top) / 5.0) * 5))
            .OrderBy(g => g.Key);

        foreach (var grp in groups)
        {
            var words = grp.OrderBy(w => w.BoundingBox.Left).ToList();
            if (words.Count == 0)
                continue;

            string text = string.Join(" ", words.Select(w => Strip(w.Text ?? "")).Where(t => t.Length > 0));
            if (string.IsNullOrWhiteSpace(text))
                continue;

            var fontSizes = new List<double>();
            foreach (var word in words)
            {
                foreach (var letter in word.Letters)
                {
                    double fs = letter.FontSize > 0 ? letter.FontSize : 0;
                    if (fs > 0)
                        fontSizes.Add(fs);
                }
            }

            double medianSize = fontSizes.Count > 0 ? Median(fontSizes) : 10.0;
            bool isBold = words.Any(w =>
                w.Letters.Any(l => (l.FontName ?? "").IndexOf("Bold", StringComparison.OrdinalIgnoreCase) >= 0));
            bool isAllCaps = text.Any(char.IsLetter) && text.Where(char.IsLetter).All(char.IsUpper);
            bool hasNums = Regex.IsMatch(text, @"\d{3,}");

            lines.Add(new RichLine
            {
                Text = text,
                FontSize = medianSize,
                XPosition = words.Min(w => w.BoundingBox.Left),
                YPosition = grp.Key,
                IsBold = isBold,
                IsAllCaps = isAllCaps,
                WordCount = text.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length,
                HasNumbers = hasNums,
                PageNumber = page.Number
            });
        }

        return lines;
    }

    private static double Median(List<double> values)
    {
        var sorted = values.OrderBy(v => v).ToList();
        int mid = sorted.Count / 2;
        return sorted.Count % 2 == 0 ? (sorted[mid - 1] + sorted[mid]) / 2.0 : sorted[mid];
    }
}

public class HeadingDetector
{
    public List<DiscoveredHeading> Discover(List<List<RichLine>> richPages)
    {
        var allSizes = richPages
            .SelectMany(p => p)
            .Where(l => l.WordCount >= 3)
            .Select(l => Math.Round(l.FontSize, 1))
            .ToList();

        double bodySize = allSizes.Count > 0 ? Mode(allSizes) : 10.0;
        double leftMargin = richPages
            .SelectMany(p => p)
            .Where(l => l.WordCount >= 2)
            .Select(l => l.XPosition)
            .OrderBy(x => x)
            .Take(Math.Max(1, (int)(richPages.SelectMany(p => p).Count() * 0.1)))
            .DefaultIfEmpty(0)
            .Average();

        var headings = new List<DiscoveredHeading>();
        foreach (var page in richPages)
        {
            foreach (var line in page)
            {
                double score = HeadingScore(line, bodySize, leftMargin);
                if (score >= 0.45)
                {
                    headings.Add(new DiscoveredHeading
                    {
                        Text = line.Text,
                        PageNumber = line.PageNumber,
                        FontSize = line.FontSize,
                        HeadingScore = score
                    });
                }
            }
        }

        return headings;
    }

    private static double HeadingScore(RichLine line, double bodySize, double leftMargin)
    {
        double score = 0.0;

        if (line.HasNumbers && line.WordCount <= 4)
            return 0.0;
        if (line.WordCount > 16)
            return 0.0;
        if (line.Text.Trim().Length < 3)
            return 0.0;

        if (bodySize > 0 && line.FontSize > bodySize)
        {
            double ratio = (line.FontSize - bodySize) / bodySize;
            score += Math.Min(0.40, ratio * 1.5);
        }

        if (line.IsBold && line.WordCount <= 12)
            score += 0.25;
        if (line.IsAllCaps && line.WordCount <= 10)
            score += 0.20;
        if (line.XPosition <= leftMargin + 20)
            score += 0.10;
        if (line.WordCount <= 4)
            score += 0.15;
        else if (line.WordCount <= 8)
            score += 0.10;
        if (line.HasNumbers)
            score -= 0.25;

        return Math.Min(1.0, score);
    }

    private static double Mode(List<double> values) =>
        values.GroupBy(v => v).OrderByDescending(g => g.Count()).First().Key;
}

public class DocumentSectionBuilder
{
    private static readonly Regex NumberRe = new(@"\b\d{2,}\b", RegexOptions.Compiled);

    public List<DiscoveredSection> Build(
        PdfReport report,
        List<DiscoveredHeading> headings,
        double minHeadingScore = 0.55)
    {
        var boundaries = headings
            .Where(h => h.HeadingScore >= minHeadingScore)
            .GroupBy(h => h.PageNumber)
            .Select(g => g.OrderByDescending(h => h.HeadingScore).First())
            .OrderBy(h => h.PageNumber)
            .ToList();

        var sections = new List<DiscoveredSection>();
        if (boundaries.Count == 0)
        {
            sections.Add(BuildSection("Document", 1, report.Pages.Count, report.Pages, 10.0));
            return sections;
        }

        if (boundaries[0].PageNumber > 1)
        {
            int end = boundaries[0].PageNumber - 1;
            sections.Add(BuildSection(
                report.Pages[0].Split('\n').FirstOrDefault(l => l.Trim().Length > 5)?.Trim() ?? "Preamble",
                1,
                end,
                report.Pages.Take(end).ToList(),
                10.0));
        }

        for (int b = 0; b < boundaries.Count; b++)
        {
            int start = boundaries[b].PageNumber;
            int end = b + 1 < boundaries.Count ? boundaries[b + 1].PageNumber - 1 : report.Pages.Count;
            var pageTexts = report.Pages.Skip(start - 1).Take(end - start + 1).ToList();
            sections.Add(BuildSection(boundaries[b].Text, start, end, pageTexts, boundaries[b].FontSize));
        }

        return sections;
    }

    private static DiscoveredSection BuildSection(
        string title,
        int start,
        int end,
        List<string> pages,
        double headingFontSize)
    {
        string fullText = string.Join("\n", pages);
        var words = fullText
            .ToLowerInvariant()
            .Split(new[] { ' ', '\n', '\r', '\t', '.', ',', ';', ':' }, StringSplitOptions.RemoveEmptyEntries)
            .Where(w => w.Length >= 5 && !Regex.IsMatch(w, @"^\d+$") && w.All(char.IsLetter))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var numbers = NumberRe.Matches(fullText)
            .Select(m => m.Value)
            .Distinct()
            .OrderBy(n => n)
            .ToList();

        return new DiscoveredSection
        {
            Title = title.Trim(),
            StartPage = start,
            EndPage = end,
            PageTexts = pages,
            HeadingFontSize = headingFontSize,
            ContentWords = words,
            Numbers = numbers
        };
    }
}

public class BoilerplateDetector
{
    public HashSet<string> Discover(PdfReport report1, PdfReport report2)
    {
        var boilerplate = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int totalPages = report1.Pages.Count + report2.Pages.Count;
        int threshold = Math.Max(3, totalPages / 8);
        var lineCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        foreach (var page in report1.Pages.Concat(report2.Pages))
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var rawLine in page.Split('\n'))
            {
                string line = Normalise(rawLine);
                if (line.Length < 5 || seen.Contains(line))
                    continue;
                seen.Add(line);
                lineCounts[line] = lineCounts.GetValueOrDefault(line, 0) + 1;
            }
        }

        foreach (var (line, count) in lineCounts)
        {
            if (count >= threshold)
                boilerplate.Add(line);
        }

        return boilerplate;
    }

    public string StripBoilerplate(string text, HashSet<string> boilerplate)
    {
        var lines = text.Split('\n').Where(l => !boilerplate.Contains(Normalise(l))).ToList();
        return string.Join("\n", lines);
    }

    public double BoilerplateFraction(string pageText, HashSet<string> boilerplate)
    {
        var lines = pageText.Split('\n').Select(l => l.Trim()).Where(l => l.Length >= 5).ToList();
        if (lines.Count == 0)
            return 0.0;
        int boilerCount = lines.Count(l => boilerplate.Contains(Normalise(l)));
        return (double)boilerCount / lines.Count;
    }

    public static string Normalise(string s)
    {
        if (string.IsNullOrWhiteSpace(s))
            return "";
        string t = s.Normalize(NormalizationForm.FormKD).ToLowerInvariant().Trim();
        t = Regex.Replace(t, @"[\p{C}\u00AD\u200B-\u200F\uFEFF\uFFF0-\uFFFD]", "");
        t = Regex.Replace(t, @"\s+", " ").Trim();
        t = Regex.Replace(t, @"^[-–—\s]*\d{1,4}[-–—\s]*$", "").Trim();
        return t;
    }
}

public class SectionMatcher
{
    public List<(DiscoveredSection? S1, DiscoveredSection? S2)> Match(
        List<DiscoveredSection> sections1,
        List<DiscoveredSection> sections2,
        HashSet<string> boilerplate)
    {
        var result = new List<(DiscoveredSection?, DiscoveredSection?)>();
        var used2 = new HashSet<int>();

        for (int i = 0; i < sections1.Count; i++)
        {
            var s1 = sections1[i];
            int bestIdx = -1;
            double bestSc = 0.0;

            for (int j = 0; j < sections2.Count; j++)
            {
                if (used2.Contains(j))
                    continue;
                double sc = SectionSimilarity(s1, sections2[j], boilerplate);
                if (sc > bestSc)
                {
                    bestSc = sc;
                    bestIdx = j;
                }
            }

            if (bestIdx >= 0 && bestSc >= 0.20)
            {
                used2.Add(bestIdx);
                result.Add((s1, sections2[bestIdx]));
            }
            else
            {
                result.Add((s1, null));
            }
        }

        for (int j = 0; j < sections2.Count; j++)
        {
            if (!used2.Contains(j))
                result.Add((null, sections2[j]));
        }

        return result;
    }

    private static double SectionSimilarity(
        DiscoveredSection s1,
        DiscoveredSection s2,
        HashSet<string> boilerplate)
    {
        string title1 = BoilerplateDetector.Normalise(s1.Title);
        string title2 = BoilerplateDetector.Normalise(s2.Title);
        double titleSim = PageAlignmentService.SequenceSimilarity(title1, title2);

        string clean1 = string.Join("\n", s1.PageTexts.Where(p => !boilerplate.Contains(BoilerplateDetector.Normalise(p))));
        string clean2 = string.Join("\n", s2.PageTexts.Where(p => !boilerplate.Contains(BoilerplateDetector.Normalise(p))));

        var w1 = clean1
            .ToLowerInvariant()
            .Split(new[] { ' ', '\n', '\r', '\t', '.', ',', ';', ':' }, StringSplitOptions.RemoveEmptyEntries)
            .Where(w => w.Length >= 5 && w.Any(char.IsLetter))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var w2 = clean2
            .ToLowerInvariant()
            .Split(new[] { ' ', '\n', '\r', '\t', '.', ',', ';', ':' }, StringSplitOptions.RemoveEmptyEntries)
            .Where(w => w.Length >= 5 && w.Any(char.IsLetter))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        int inter = w1.Intersect(w2).Count();
        int union = w1.Union(w2).Count();
        double wordSim = union == 0 ? 0.0 : (double)inter / union;

        var n1 = new HashSet<string>(s1.Numbers);
        var n2 = new HashSet<string>(s2.Numbers);
        int ni = n1.Intersect(n2).Count();
        int nu = n1.Union(n2).Count();
        double numSim = nu == 0 ? 0.0 : (double)ni / nu;

        string open1 = s1.FullText[..Math.Min(300, s1.FullText.Length)].ToLowerInvariant();
        string open2 = s2.FullText[..Math.Min(300, s2.FullText.Length)].ToLowerInvariant();
        double openSim = PageAlignmentService.SequenceSimilarity(open1, open2);

        return titleSim * 0.30 + wordSim * 0.30 + numSim * 0.25 + openSim * 0.15;
    }
}

public class SectionDiffAnalyser
{
    private readonly LineComparatorService _lineComparator;

    public SectionDiffAnalyser(LineComparatorService lineComparator)
    {
        _lineComparator = lineComparator;
    }

    public SectionComparisonResult Analyse(
        DiscoveredSection? s1,
        DiscoveredSection? s2,
        HashSet<string> boilerplate)
    {
        if (s1 == null && s2 != null)
        {
            return new SectionComparisonResult
            {
                Section2 = s2,
                ChangeKind = "added",
                IsMaterial = true,
                Description = $"NEW SECTION in AFS2 - \"{s2.Title}\" (AFS2 pages {s2.StartPage}-{s2.EndPage}, {s2.PageCount} page(s)). This section does not exist in AFS1."
            };
        }

        if (s1 != null && s2 == null)
        {
            return new SectionComparisonResult
            {
                Section1 = s1,
                ChangeKind = "removed",
                IsMaterial = true,
                Description = $"SECTION REMOVED - \"{s1.Title}\" was on AFS1 pages {s1.StartPage}-{s1.EndPage} ({s1.PageCount} page(s)) but has no counterpart in AFS2."
            };
        }

        string clean1 = StripBoilerplateLines(s1!.FullText, boilerplate);
        string clean2 = StripBoilerplateLines(s2!.FullText, boilerplate);
        double contentSim = PageAlignmentService.SequenceSimilarity(clean1.ToLowerInvariant(), clean2.ToLowerInvariant());
        double bpFrac1 = BoilerplateFraction(s1.FullText, boilerplate);
        double bpFrac2 = BoilerplateFraction(s2.FullText, boilerplate);

        bool boilerplateOnly = bpFrac1 >= 0.70 && bpFrac2 >= 0.70 && contentSim < 0.40 && clean1.Length < 200 && clean2.Length < 200;
        if (boilerplateOnly)
        {
            var bpLineDiffs = _lineComparator.CompareLines(s1.FullText, s2.FullText);
            return new SectionComparisonResult
            {
                Section1 = s1,
                Section2 = s2,
                ChangeKind = "boilerplate-only",
                ContentSim = contentSim,
                BoilerplateSim = Math.Max(bpFrac1, bpFrac2),
                IsMaterial = true,
                LineDiffs = bpLineDiffs,
                Description = $"BOILERPLATE-ONLY MATCH: AFS1 p{s1.StartPage} and AFS2 p{s2.StartPage} were matched because they share recurring content. The actual unique content of these pages is very different (similarity {contentSim:P0})."
            };
        }

        double pageRatio = s1.PageCount > 0 ? (double)s2.PageCount / s1.PageCount : 1.0;
        bool significantlyLarger = s2.PageCount > s1.PageCount + 2 && pageRatio > 1.5;
        bool significantlySmaller = s1.PageCount > s2.PageCount + 2 && pageRatio < 0.67;
        bool pageShift = s1.StartPage != s2.StartPage && contentSim >= 0.85;

        List<LineDiffResult> lineDiffs;
        if (contentSim >= 0.96)
        {
            lineDiffs = new List<LineDiffResult>();
        }
        else
        {
            string t1 = s1.PageCount <= s2.PageCount ? s1.FullText : string.Join("\n", s1.PageTexts.Take(s2.PageCount));
            string t2 = s2.PageCount <= s1.PageCount ? s2.FullText : string.Join("\n", s2.PageTexts.Take(s1.PageCount));
            lineDiffs = _lineComparator.CompareLines(t1, t2);
        }

        string kind;
        string description;
        bool material;

        if (contentSim >= 0.96 && !significantlyLarger && !significantlySmaller)
        {
            kind = pageShift ? "page-shift" : "same";
            material = false;
            description = pageShift
                ? $"PAGE SHIFT: \"{s1.Title}\" content is on AFS1 p{s1.StartPage}-{s1.EndPage} and AFS2 p{s2.StartPage}-{s2.EndPage}."
                : $"UNCHANGED: \"{s1.Title}\" - identical in both drafts.";
        }
        else if (significantlyLarger)
        {
            kind = "expanded";
            material = true;
            description = $"EXPANDED: \"{s1.Title}\" grew from {s1.PageCount} page(s) to {s2.PageCount} page(s). Content similarity on overlapping portion: {contentSim:P0}.";
        }
        else if (significantlySmaller)
        {
            kind = "contracted";
            material = true;
            description = $"CONTRACTED: \"{s1.Title}\" shrank from {s1.PageCount} page(s) to {s2.PageCount} page(s). Content removed from this section in AFS2.";
        }
        else if (contentSim >= 0.85)
        {
            kind = pageShift ? "page-shift" : "same";
            material = false;
            description = pageShift
                ? $"PAGE SHIFT: \"{s1.Title}\" - content similarity {contentSim:P0}, AFS1 p{s1.StartPage} to AFS2 p{s2.StartPage}."
                : $"ESSENTIALLY UNCHANGED: \"{s1.Title}\" - similarity {contentSim:P0}.";
        }
        else
        {
            kind = "changed";
            material = true;
            description = $"CHANGED: \"{s1.Title}\" - AFS1 p{s1.StartPage}-{s1.EndPage} vs AFS2 p{s2.StartPage}-{s2.EndPage}. Content similarity: {contentSim:P0}. Page count: {s1.PageCount} to {s2.PageCount}.";
        }

        return new SectionComparisonResult
        {
            Section1 = s1,
            Section2 = s2,
            ChangeKind = kind,
            ContentSim = contentSim,
            BoilerplateSim = Math.Max(bpFrac1, bpFrac2),
            IsMaterial = material,
            LineDiffs = lineDiffs,
            Description = description
        };
    }

    private static string StripBoilerplateLines(string text, HashSet<string> boilerplate)
    {
        var lines = text.Split('\n').Where(l => !boilerplate.Contains(BoilerplateDetector.Normalise(l))).ToList();
        return string.Join("\n", lines);
    }

    private static double BoilerplateFraction(string text, HashSet<string> boilerplate)
    {
        var lines = text.Split('\n').Select(l => l.Trim()).Where(l => l.Length >= 5).ToList();
        if (lines.Count == 0)
            return 0.0;
        int bp = lines.Count(l => boilerplate.Contains(BoilerplateDetector.Normalise(l)));
        return (double)bp / lines.Count;
    }
}

public class DocumentIntelligenceService
{
    private readonly RichPageExtractor _extractor;
    private readonly HeadingDetector _headingDetector;
    private readonly DocumentSectionBuilder _sectionBuilder;
    private readonly BoilerplateDetector _boilerplateDetector;
    private readonly SectionMatcher _matcher;
    private readonly SectionDiffAnalyser _analyser;

    public DocumentIntelligenceService(LineComparatorService lineComparator)
    {
        _extractor = new RichPageExtractor();
        _headingDetector = new HeadingDetector();
        _sectionBuilder = new DocumentSectionBuilder();
        _boilerplateDetector = new BoilerplateDetector();
        _matcher = new SectionMatcher();
        _analyser = new SectionDiffAnalyser(lineComparator);
    }

    public DocumentComparisonReport Compare(PdfReport report1, PdfReport report2)
    {
        var richPages1 = _extractor.ExtractRichPages(report1.PdfPath);
        var richPages2 = _extractor.ExtractRichPages(report2.PdfPath);
        var headings1 = _headingDetector.Discover(richPages1);
        var headings2 = _headingDetector.Discover(richPages2);
        var sections1 = _sectionBuilder.Build(report1, headings1);
        var sections2 = _sectionBuilder.Build(report2, headings2);
        var boilerplate = _boilerplateDetector.Discover(report1, report2);
        var pairs = _matcher.Match(sections1, sections2, boilerplate);
        var sectionResults = pairs.Select(p => _analyser.Analyse(p.S1, p.S2, boilerplate)).ToList();

        int expanded = sectionResults.Count(r => r.ChangeKind == "expanded");
        int contracted = sectionResults.Count(r => r.ChangeKind == "contracted");
        int changed = sectionResults.Count(r => r.ChangeKind == "changed");
        int added = sectionResults.Count(r => r.ChangeKind == "added");
        int removed = sectionResults.Count(r => r.ChangeKind == "removed");
        int bpOnly = sectionResults.Count(r => r.ChangeKind == "boilerplate-only");
        int pageShift = sectionResults.Count(r => r.ChangeKind == "page-shift");

        string summary =
            $"AFS1: {sections1.Count} sections, {report1.PageCount} pages. " +
            $"AFS2: {sections2.Count} sections, {report2.PageCount} pages. " +
            $"Boilerplate lines discovered: {boilerplate.Count}. " +
            $"Section changes: {expanded} expanded, {contracted} contracted, {changed} changed, {added} added, {removed} removed" +
            (bpOnly > 0 ? $", {bpOnly} boilerplate-only match(es) - manual review needed" : "") +
            (pageShift > 0 ? $", {pageShift} page-shift(s) - same content, different page numbers" : "") +
            ".";

        return new DocumentComparisonReport
        {
            Sections = sectionResults,
            Summary = summary,
            TotalPages1 = report1.PageCount,
            TotalPages2 = report2.PageCount
        };
    }
}
