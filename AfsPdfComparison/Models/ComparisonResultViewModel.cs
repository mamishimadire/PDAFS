// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — ViewModel                                   ║
// ║  ComparisonResultViewModel — used by ComparisonController / Result view.    ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using AfsPdfComparison.Services;

namespace AfsPdfComparison.Models
{
    /// <summary>
    /// View model for the ComparisonController.Compare action result view.
    /// Carries the full output of the new LineComparisonService pipeline.
    /// </summary>
    public class ComparisonResultViewModel
    {
        public string  Report1Filename   { get; set; } = "";
        public string  Report2Filename   { get; set; } = "";
        public int     Report1PageCount  { get; set; }
        public int     Report2PageCount  { get; set; }
        public int     Report1WordCount  { get; set; }
        public int     Report2WordCount  { get; set; }

        public ComparisonCounts  Counts           { get; set; } = new();
        public List<LineDiff>    Diffs            { get; set; } = new();

        public double       NumSimilarityPct { get; set; }
        public List<string> NumOnlyInAfs1    { get; set; } = new();
        public List<string> NumOnlyInAfs2    { get; set; } = new();

        public string Snapshot1Base64 { get; set; } = "";
        public string Snapshot2Base64 { get; set; } = "";
    }
}
