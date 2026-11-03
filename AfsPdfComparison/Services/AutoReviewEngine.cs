// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  SERVICE — AutoReviewEngine                                                  ║
// ║  Deterministic audit review engine — generates one finding per changed,     ║
// ║  added, or removed line detected by the diff engine. No AI required.        ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using System.Text.RegularExpressions;
using AfsPdfComparison.Models;

namespace AfsPdfComparison.Services;

public class AutoReviewEngine
{
    private readonly LabeledNumberExtractor _labelExtractor;

    public AutoReviewEngine(LabeledNumberExtractor labelExtractor)
    {
        _labelExtractor = labelExtractor;
    }

    public (List<AutoReviewFinding> Findings, List<LabeledNumberChange> NumberChanges)
        GenerateFindings(ComparisonResult comparison)
    {
        var findings      = new List<AutoReviewFinding>();
        var numberChanges = new List<LabeledNumberChange>();

        // #1 — Build section map upfront so every finding can get a SectionGroup
        var sectionMap1 = BuildSectionMap(comparison.Report1.Pages);
        var sectionMap2 = BuildSectionMap(comparison.Report2.Pages);

        try { findings.AddRange(FromChangedLines(comparison, sectionMap1, sectionMap2)); }   catch { }
        try { findings.AddRange(FromAddedLines(comparison, sectionMap2)); }                  catch { }
        try { findings.AddRange(FromRemovedLines(comparison, sectionMap1)); }                catch { }
        try { findings.AddRange(FindUnmatchedPages(comparison, sectionMap1, sectionMap2)); } catch { }
        // #4 — Comparative column validation (GRAP 3 / IAS 8)
        try { findings.AddRange(ValidateComparativeColumns(comparison)); }                   catch { }
        // #8 — Cross-reference integrity check
        try { findings.AddRange(CheckCrossReferences(comparison)); }                         catch { }

        try
        {
            var n1 = _labelExtractor.Extract(comparison.Report1.Pages);
            var n2 = _labelExtractor.Extract(comparison.Report2.Pages);
            numberChanges = _labelExtractor.Compare(n1, n2, comparison.PageDiffs);
        }
        catch { }

        // Deduplicate by title, then sort high → medium → low, then page
        var seen    = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var deduped = new List<AutoReviewFinding>(findings.Count);
        foreach (var f in findings)
            if (seen.Add(f.Title))
                deduped.Add(f);

        deduped.Sort((a, b) =>
        {
            int sa = SeverityRank(a.Severity), sb = SeverityRank(b.Severity);
            if (sa != sb) return sb - sa;
            // #7 — Higher confidence first within same severity
            int ca = a.Confidence, cb = b.Confidence;
            if (ca != cb) return cb - ca;
            return (a.PageAfs1 ?? a.PageAfs2 ?? 999) - (b.PageAfs1 ?? b.PageAfs2 ?? 999);
        });

        return (deduped, numberChanges);
    }

    // ── #1 Section Map ─────────────────────────────────────────────────────────
    // Maps each page index → the financial section heading active on that page.
    // Detects "Note X — Title", "Statement of ...", numbered section headings.

    private static readonly Regex _sectionHeadingRe = new(
        @"^(?:note\s+\d+(?:\.\d+)*\s*[-–—:]\s*[a-z].{0,80}|" +           // Note 12 — Provisions
        @"statement\s+of\s+[a-z ]{4,60}|" +                                // Statement of Financial Position
        @"\d+(?:\.\d+)+\s+[A-Z][a-z ]{3,60}|" +                           // 1.2  Revenue
        @"(?:appendix|schedule|section)\s+[a-z0-9 ]{2,60})$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static Dictionary<int, string> BuildSectionMap(List<string> pages)
    {
        var map     = new Dictionary<int, string>();
        string last = "General";
        for (int i = 0; i < pages.Count; i++)
        {
            foreach (var line in pages[i].Split('\n'))
            {
                string t = line.Trim();
                if (t.Length >= 6 && t.Length <= 100 && _sectionHeadingRe.IsMatch(t))
                {
                    last = t.Length > 60 ? t[..60] + "…" : t;
                    break;
                }
            }
            map[i] = last;
        }
        return map;
    }

    private static string GetSection(Dictionary<int, string> map, int? pageNum)
    {
        if (pageNum.HasValue && map.TryGetValue(pageNum.Value - 1, out string? s))
            return s;
        return "General";
    }

    // ── #7 Confidence scoring ──────────────────────────────────────────────────
    // Returns 0–100. High = genuine change, Low = likely layout artefact.

    private static int ComputeConfidence(
        LineDiffResult d, PageDiffResult pd, bool numbersDiffer, bool signFlip)
    {
        int score = 70;  // baseline

        // Confirmed number change: very reliable
        if (signFlip)       score += 25;
        else if (numbersDiffer) score += 15;

        // High-similarity page pair = mostly same, diff is probably real
        if (pd.PctSame < 0.90) score += 5;
        else if (pd.PctSame >= 0.99) score -= 15; // very similar page — diff is suspect

        // Word diff supports the change
        if (d.WordDiff.Count > 0 && d.WordDiff.Any(w => w.Tag != "same"))
            score += 5;

        // Very short lines are often header/footer artefacts
        if ((d.Line1?.Length ?? 0) < 10 || (d.Line2?.Length ?? 0) < 10)
            score -= 10;

        // Similarity close to threshold suggests a weak match
        if (d.Similarity < 0.70) score -= 10;

        return Math.Clamp(score, 10, 100);
    }

    // ── 1. One finding per CHANGED line ──────────────────────────────────────

    private static List<AutoReviewFinding> FromChangedLines(
        ComparisonResult cmp,
        Dictionary<int, string> sectionMap1,
        Dictionary<int, string> sectionMap2)
    {
        var findings = new List<AutoReviewFinding>();

        foreach (var pd in cmp.PageDiffs.Where(p => p.Changed > 0))
        {
            int ordinal = 0;
            foreach (var d in pd.Diffs.Where(d => d.Status == "changed"))
            {
                ordinal++;
                var nums1 = ExtractValues(d.Line1);
                var nums2 = ExtractValues(d.Line2);
                bool hasNumbers    = nums1.Count > 0 || nums2.Count > 0;
                bool numbersDiffer = hasNumbers && !new HashSet<string>(
                    ExtractRawTokens(d.Line1)).SetEquals(ExtractRawTokens(d.Line2));

                bool signFlip = false;
                decimal n1val = 0, n2val = 0;
                if (nums1.Count > 0 && nums2.Count > 0)
                {
                    n1val = nums1[0]; n2val = nums2[0];
                    signFlip = (n1val > 0 && n2val < 0) || (n1val < 0 && n2val > 0);
                }

                bool isPolicyLine    = PolicyRe.IsMatch(d.Line1 + d.Line2);
                bool isPersonnelLine = PersonnelRe.IsMatch(d.Line1 + d.Line2);

                string severity = (signFlip || isPolicyLine || isPersonnelLine) ? "high"
                                : numbersDiffer ? "medium"
                                : "low";
                string category = signFlip        ? "sign-change"
                                : isPolicyLine    ? "policy-change"
                                : isPersonnelLine ? "director-change"
                                : numbersDiffer   ? "numeric-change"
                                : "text-change";

                string label  = ShortLabel(d.Line1);
                string pctStr = "";
                double? pctChange = null;
                if (numbersDiffer && n1val != 0)
                {
                    double pct = (double)Math.Abs(n2val - n1val) / (double)Math.Abs(n1val);
                    pctStr    = pct >= 10 ? " (>1000%)" : $" ({pct:P1})";
                    pctChange = Math.Round(pct * 100, 1);
                }

                // #1 section group, #7 confidence
                string section    = GetSection(sectionMap1, pd.PageAfs1) is var s1 && s1 != "General"
                                    ? s1 : GetSection(sectionMap2, pd.PageAfs2);
                int    confidence = ComputeConfidence(d, pd, numbersDiffer, signFlip);

                findings.Add(new AutoReviewFinding
                {
                    Category     = category,
                    Title        = $"Changed — p.{pd.PageAfs1 ?? pd.PageAfs2} #{ordinal}: {label}{pctStr}",
                    Description  = BuildChangedDesc(d, n1val, n2val, signFlip, numbersDiffer, pctStr),
                    Severity     = severity,
                    PageAfs1     = pd.PageAfs1,
                    PageAfs2     = pd.PageAfs2,
                    PairIndex    = pd.PairIndex,
                    Before       = Trunc(d.Line1, 160),
                    After        = Trunc(d.Line2, 160),
                    PctChange    = pctChange,
                    AbsChange    = numbersDiffer ? Math.Abs(n2val - n1val) : null,
                    Confidence   = confidence,
                    SectionGroup = section,
                });
            }
        }

        return findings;
    }

    // ── 2. One finding per PAGE PAIR for ADDED lines ──────────────────────────

    private static List<AutoReviewFinding> FromAddedLines(
        ComparisonResult cmp, Dictionary<int, string> sectionMap2)
    {
        var findings = new List<AutoReviewFinding>();

        foreach (var pd in cmp.PageDiffs.Where(p => p.Added > 0))
        {
            var added = pd.Diffs.Where(d => d.Status == "added").ToList();
            if (added.Count == 0) continue;

            bool hasNumbers = added.Any(d => ExtractRawTokens(d.Line2).Count > 0);
            string severity = hasNumbers ? "medium" : "low";
            if (added.Any(d => PersonnelRe.IsMatch(d.Line2) || PolicyRe.IsMatch(d.Line2)))
                severity = "high";

            string preview = string.Join(" | ", added.Take(3).Select(d => Trunc(d.Line2, 60)));
            int page = pd.PageAfs2 ?? pd.PageAfs1 ?? 0;

            findings.Add(new AutoReviewFinding
            {
                Category     = "content-added",
                Title        = $"Added — p.{page}: {added.Count} new line(s) in AFS 2",
                Description  = $"{added.Count} line(s) present in AFS 2 (page {pd.PageAfs2}) but absent from AFS 1 (page {pd.PageAfs1}). Preview: \"{preview}\".",
                Severity     = severity,
                PageAfs1     = pd.PageAfs1,
                PageAfs2     = pd.PageAfs2,
                PairIndex    = pd.PairIndex,
                After        = preview,
                Confidence   = hasNumbers ? 75 : 60,
                SectionGroup = GetSection(sectionMap2, pd.PageAfs2),
            });
        }

        return findings;
    }

    // ── 3. One finding per PAGE PAIR for REMOVED lines ────────────────────────

    private static List<AutoReviewFinding> FromRemovedLines(
        ComparisonResult cmp, Dictionary<int, string> sectionMap1)
    {
        var findings = new List<AutoReviewFinding>();

        foreach (var pd in cmp.PageDiffs.Where(p => p.Removed > 0))
        {
            var removed = pd.Diffs.Where(d => d.Status == "removed").ToList();
            if (removed.Count == 0) continue;

            bool hasNumbers = removed.Any(d => ExtractRawTokens(d.Line1).Count > 0);
            string severity = hasNumbers ? "medium" : "low";
            if (removed.Any(d => PersonnelRe.IsMatch(d.Line1) || PolicyRe.IsMatch(d.Line1)))
                severity = "high";

            string preview = string.Join(" | ", removed.Take(3).Select(d => Trunc(d.Line1, 60)));
            int page = pd.PageAfs1 ?? pd.PageAfs2 ?? 0;

            findings.Add(new AutoReviewFinding
            {
                Category     = "content-removed",
                Title        = $"Removed — p.{page}: {removed.Count} line(s) removed from AFS 1",
                Description  = $"{removed.Count} line(s) present in AFS 1 (page {pd.PageAfs1}) but absent from AFS 2 (page {pd.PageAfs2}). Preview: \"{preview}\".",
                Severity     = severity,
                PageAfs1     = pd.PageAfs1,
                PageAfs2     = pd.PageAfs2,
                PairIndex    = pd.PairIndex,
                Before       = preview,
                Confidence   = hasNumbers ? 75 : 60,
                SectionGroup = GetSection(sectionMap1, pd.PageAfs1),
            });
        }

        return findings;
    }

    // ── 4. Unmatched pages ────────────────────────────────────────────────────

    private static List<AutoReviewFinding> FindUnmatchedPages(
        ComparisonResult cmp,
        Dictionary<int, string> sectionMap1,
        Dictionary<int, string> sectionMap2)
    {
        var findings = new List<AutoReviewFinding>();

        foreach (var pd in cmp.PageDiffs.Where(p => p.Unmatched))
        {
            bool isAdded = pd.PageAfs1 == null;
            string pageText = isAdded
                ? cmp.Report2.Pages.ElementAtOrDefault((pd.PageAfs2 ?? 1) - 1) ?? ""
                : cmp.Report1.Pages.ElementAtOrDefault((pd.PageAfs1 ?? 1) - 1) ?? "";
            string heading = PageHeading(pageText);
            string section = isAdded
                ? GetSection(sectionMap2, pd.PageAfs2)
                : GetSection(sectionMap1, pd.PageAfs1);

            findings.Add(new AutoReviewFinding
            {
                Category     = isAdded ? "section-added" : "section-removed",
                Title        = $"{(isAdded ? "Page added" : "Page removed")} — p.{(isAdded ? pd.PageAfs2 : pd.PageAfs1)}: {heading}",
                Description  = isAdded
                    ? $"Page {pd.PageAfs2} in AFS 2 has no matching page in AFS 1. This is an entirely new page. Heading: \"{heading}\"."
                    : $"Page {pd.PageAfs1} in AFS 1 has no matching page in AFS 2 — content was removed. Heading: \"{heading}\".",
                Severity     = "high",
                Confidence   = 90,
                PageAfs1     = pd.PageAfs1,
                PageAfs2     = pd.PageAfs2,
                PairIndex    = pd.PairIndex,
                SectionGroup = section,
            });
        }

        return findings;
    }

    // ── #4 Year-over-year comparative column validation ───────────────────────
    // AFS documents have two columns: current year (col A) and prior year comparative (col B).
    // When AFS 2 is the new report, AFS 2 col B (prior year) MUST equal AFS 1 col A (current year).
    // A mismatch is a GRAP 3 / IAS 8 restatement red flag.

    private static readonly Regex _twoColNumRe = new(
        @"^(.{3,60}?)\s{2,}([\d\s,()R-]+)\s{2,}([\d\s,()R-]+)\s*$",
        RegexOptions.Compiled);

    private static List<AutoReviewFinding> ValidateComparativeColumns(ComparisonResult cmp)
    {
        var findings = new List<AutoReviewFinding>();

        foreach (var pd in cmp.PageDiffs.Where(p => !p.Unmatched && p.PageAfs1.HasValue && p.PageAfs2.HasValue))
        {
            string page1 = cmp.Report1.Pages.ElementAtOrDefault(pd.PageAfs1!.Value - 1) ?? "";
            string page2 = cmp.Report2.Pages.ElementAtOrDefault(pd.PageAfs2!.Value - 1) ?? "";

            var rows1 = ParseTwoColRows(page1);
            var rows2 = ParseTwoColRows(page2);

            foreach (var kvp in rows1)
            {
                string label    = kvp.Key;
                string curYear1 = kvp.Value.col1;
                if (!rows2.TryGetValue(label, out var r2)) continue;
                string priorYear2 = r2.col2;  // col2 of AFS 2 = prior year comparative

                string n1 = NormNum(curYear1);
                string n2 = NormNum(priorYear2);
                if (n1.Length < 2 || n2.Length < 2 || n1 == n2) continue;
                if (!decimal.TryParse(n1, out decimal v1) || !decimal.TryParse(n2, out decimal v2)) continue;
                if (v1 == 0 && v2 == 0) continue;

                double pct = v1 != 0 ? Math.Abs((double)(v2 - v1) / (double)v1) * 100 : 100;
                if (pct < 0.5) continue;  // < 0.5% tolerance for rounding

                string sev = pct >= 5 ? "high" : "medium";
                findings.Add(new AutoReviewFinding
                {
                    Category     = "comparative-mismatch",
                    Title        = $"Comparative mismatch — {Trunc(label, 50)}: AFS1={curYear1} vs AFS2 prior={priorYear2}",
                    Description  = $"GRAP 3 / IAS 8: The prior-year comparative in AFS 2 ({priorYear2}) does not match the current-year amount in AFS 1 ({curYear1}) for \"{label}\". " +
                                   $"Difference: {pct:F1}%. This may indicate a restatement that should be disclosed.",
                    Severity     = sev,
                    Confidence   = 85,
                    PageAfs1     = pd.PageAfs1,
                    PageAfs2     = pd.PageAfs2,
                    PairIndex    = pd.PairIndex,
                    Before       = $"AFS 1 current year: {curYear1}",
                    After        = $"AFS 2 prior year:   {priorYear2}",
                    PctChange    = Math.Round(pct, 1),
                    SectionGroup = "Comparative Validation (GRAP 3)",
                });
            }
        }

        return findings;
    }

    private static Dictionary<string, (string col1, string col2)> ParseTwoColRows(string pageText)
    {
        var result = new Dictionary<string, (string, string)>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in pageText.Split('\n'))
        {
            var m = _twoColNumRe.Match(line.Trim());
            if (!m.Success) continue;
            string label = m.Groups[1].Value.Trim();
            string c1    = m.Groups[2].Value.Trim();
            string c2    = m.Groups[3].Value.Trim();
            if (label.Length >= 3 && c1.Length >= 1 && c2.Length >= 1)
                result.TryAdd(label, (c1, c2));
        }
        return result;
    }

    private static string NormNum(string s)
        => Regex.Replace(s.Replace("(", "-").Replace(")", "").Replace("R", "").Replace(",", "").Replace(" ", ""), @"[^\d.-]", "").Trim();

    // ── #8 Cross-reference integrity check ────────────────────────────────────
    // Finds "Note X", "see note X" references in AFS 2 and checks that the
    // referenced note heading actually exists in AFS 2. If it existed in AFS 1
    // but not AFS 2, it was likely removed — a disclosure gap.

    private static readonly Regex _noteRefRe = new(
        @"\b(?:see\s+)?note\s+(\d+(?:\.\d+)*)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static List<AutoReviewFinding> CheckCrossReferences(ComparisonResult cmp)
    {
        var findings = new List<AutoReviewFinding>();

        // Build note index for each document: note number → page it appears on
        var notes1 = BuildNoteIndex(cmp.Report1.Pages);
        var notes2 = BuildNoteIndex(cmp.Report2.Pages);

        // Scan AFS 2 for references to notes that no longer exist in AFS 2
        for (int pi = 0; pi < cmp.Report2.Pages.Count; pi++)
        {
            var page = cmp.Report2.Pages[pi];
            foreach (Match m in _noteRefRe.Matches(page))
            {
                string noteNum = m.Groups[1].Value;
                if (notes2.ContainsKey(noteNum)) continue;       // note exists — OK
                if (!notes1.ContainsKey(noteNum)) continue;      // not in AFS 1 either — likely a different type of ref

                // Note referenced in AFS 2 text but the note itself is gone in AFS 2
                int pageNum = pi + 1;
                string title = $"Stale cross-reference — Note {noteNum} referenced on p.{pageNum} but does not exist in AFS 2";
                if (findings.Any(f => f.Title == title)) continue;

                findings.Add(new AutoReviewFinding
                {
                    Category     = "cross-ref-broken",
                    Title        = title,
                    Description  = $"AFS 2 page {pageNum} references Note {noteNum}, but Note {noteNum} was present in AFS 1 (page {notes1[noteNum]}) and is no longer found in AFS 2. " +
                                   $"This may indicate a removed disclosure or a renumbered note that was not updated in all references.",
                    Severity     = "high",
                    Confidence   = 80,
                    PageAfs1     = notes1.TryGetValue(noteNum, out int p1) ? p1 : null,
                    PageAfs2     = pageNum,
                    SectionGroup = "Cross-Reference Integrity",
                });
            }
        }

        return findings;
    }

    private static Dictionary<string, int> BuildNoteIndex(List<string> pages)
    {
        var result  = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var headRe  = new Regex(@"^note\s+(\d+(?:\.\d+)*)\b", RegexOptions.IgnoreCase);
        for (int i = 0; i < pages.Count; i++)
            foreach (var line in pages[i].Split('\n'))
            {
                var m = headRe.Match(line.Trim());
                if (m.Success) result.TryAdd(m.Groups[1].Value, i + 1);
            }
        return result;
    }

    // ── patterns ──────────────────────────────────────────────────────────────

    private static readonly Regex PolicyRe = new(
        @"\b(accounting\s+polic|measurement\s+basis|recognition|going\s+concern|basis\s+of\s+preparation|ifrs|grap|ias\s+\d|ipsas|grap\s+\d)\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex PersonnelRe = new(
        @"\b(director|trustee|councillor|municipal\s+manager|cfo|chief\s+financial|ceo|chief\s+executive|accounting\s+officer|mayor|speaker|chairperson|chairm)\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex NumRe = new(
        @"(\()([\d][\d\s,\.]*[\d])\)|([\d][\d\s,\.]*[\d])",
        RegexOptions.Compiled);

    // ── helpers ──────────────────────────────────────────────────────────────

    private static List<decimal> ExtractValues(string line)
    {
        var result = new List<decimal>();
        foreach (Match m in NumRe.Matches(line))
        {
            bool neg = m.Groups[1].Success;
            string raw = neg ? m.Groups[2].Value : m.Groups[3].Value;
            string clean = Regex.Replace(raw, @"[\s,]", "");
            if (decimal.TryParse(clean, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out decimal val))
                result.Add(neg ? -val : val);
        }
        return result;
    }

    private static List<string> ExtractRawTokens(string line)
    {
        var result = new List<string>();
        foreach (Match m in NumRe.Matches(line))
        {
            bool neg = m.Groups[1].Success;
            string raw = neg ? m.Groups[2].Value : m.Groups[3].Value;
            string clean = Regex.Replace(raw, @"[\s,]", "");
            if (clean.Length >= 2) result.Add(clean);
        }
        return result;
    }

    private static string ShortLabel(string line)
    {
        // Take text before the first big whitespace gap (financial table label)
        var m = Regex.Match(line, @"^(.{3,50}?)\s{2,}[\d\(]");
        if (m.Success) return Trunc(m.Groups[1].Value.Trim(), 50);
        return Trunc(string.Join(" ", line.Split(' ').Take(6)), 50);
    }

    private static string BuildChangedDesc(
        LineDiffResult d, decimal n1, decimal n2,
        bool signFlip, bool numChanged, string pctStr)
    {
        var sb = new System.Text.StringBuilder();
        if (signFlip)
            sb.Append($"Sign reversal: {n1:N0} → {n2:N0}. ");
        else if (numChanged)
            sb.Append($"Amount changed: {n1:N0} → {n2:N0}{pctStr}. ");

        sb.Append($"AFS 1: \"{Trunc(d.Line1, 100)}\"");
        sb.Append($" → AFS 2: \"{Trunc(d.Line2, 100)}\".");

        if (d.WordDiff.Count > 0)
        {
            var removedWords = d.WordDiff.Where(w => w.Tag == "removed").Select(w => w.Word).Take(5).ToList();
            var addedWords   = d.WordDiff.Where(w => w.Tag == "added").Select(w => w.Word).Take(5).ToList();
            if (removedWords.Count > 0 || addedWords.Count > 0)
            {
                sb.Append(" Changed words:");
                if (removedWords.Count > 0) sb.Append($" −[{string.Join(", ", removedWords)}]");
                if (addedWords.Count > 0)   sb.Append($" +[{string.Join(", ", addedWords)}]");
                sb.Append('.');
            }
        }

        return sb.ToString();
    }

    private static string PageHeading(string text)
    {
        return text.Split('\n', StringSplitOptions.RemoveEmptyEntries)
                   .Select(l => l.Trim())
                   .Where(l => l.Length >= 4 && l.Any(char.IsLetter))
                   .Take(8)
                   .FirstOrDefault(l => l.Length <= 80) ?? "";
    }

    private static string Trunc(string s, int max)
        => (s?.Length ?? 0) <= max ? (s ?? "") : s![..max] + "…";

    private static int SeverityRank(string s)
        => s == "high" ? 2 : s == "medium" ? 1 : 0;
}
