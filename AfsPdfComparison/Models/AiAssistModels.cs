namespace AfsPdfComparison.Models;

// ── AI Direct Content Reading ───────────────────────────────────────────────

/// <summary>Result from Gemini directly reading and comparing both AFS documents.</summary>
public class AiContentReadResult
{
    public bool Enabled { get; set; }
    public string Status { get; set; } = "disabled";    // disabled | missing_key | error | ready
    public string Message { get; set; } = string.Empty;
    public string OverallAssessment { get; set; } = string.Empty;
    public string RiskLevel { get; set; } = string.Empty;          // High | Medium | Low
    public List<AiSectionDiff> SectionDiffs { get; set; } = new();
    public List<AiLineItemChange> LineItemChanges { get; set; } = new();
    public List<string> NewDisclosures { get; set; } = new();
    public List<string> RemovedDisclosures { get; set; } = new();
    public List<string> PolicyChanges { get; set; } = new();
    public string AuditConclusion { get; set; } = string.Empty;
    public List<string> Limitations { get; set; } = new();
    public DateTime GeneratedAt { get; set; } = DateTime.UtcNow;
}

public class AiSectionDiff
{
    public string SectionName { get; set; } = string.Empty;
    public string ChangeType { get; set; } = string.Empty;        // numeric | narrative | structural | unchanged
    public string Severity { get; set; } = string.Empty;          // High | Medium | Low
    public string Afs1Summary { get; set; } = string.Empty;
    public string Afs2Summary { get; set; } = string.Empty;
    public string ChangeDescription { get; set; } = string.Empty;
    public List<string> ChangedValues { get; set; } = new();
    public int? PageAfs1 { get; set; }
    public int? PageAfs2 { get; set; }
    public int? PairIndex { get; set; }
}

public class AiLineItemChange
{
    public string LineItem { get; set; } = string.Empty;
    public string ValueInAfs1 { get; set; } = string.Empty;
    public string ValueInAfs2 { get; set; } = string.Empty;
    public string PctChange { get; set; } = string.Empty;
    public string MaterialityNote { get; set; } = string.Empty;
    public int? PageAfs1 { get; set; }
    public int? PageAfs2 { get; set; }
    public int? PairIndex { get; set; }
}

public class AiAssistResult
{
    public bool Enabled { get; set; }
    public string Provider { get; set; } = "Gemini";
    public string Model { get; set; } = string.Empty;
    public string Status { get; set; } = "disabled";
    public string Message { get; set; } = string.Empty;
    public string ExecutiveSummary { get; set; } = string.Empty;
    public string OverallRisk { get; set; } = string.Empty;
    public List<AiMaterialFinding> MaterialFindings { get; set; } = new();
    public List<AiPageFocus> PageFocus { get; set; } = new();
    public List<string> HighlightGuidance { get; set; } = new();
    public List<string> Limitations { get; set; } = new();
    public DateTime GeneratedAt { get; set; } = DateTime.UtcNow;
}

public class AiMaterialFinding
{
    public string Title { get; set; } = string.Empty;
    public string Severity { get; set; } = string.Empty;
    public int? PairIndex { get; set; }
    public int? PageAfs1 { get; set; }
    public int? PageAfs2 { get; set; }
    public List<int> FocusDiffOrdinals { get; set; } = new();
    public string AuditRationale { get; set; } = string.Empty;
    public string RecommendedProcedure { get; set; } = string.Empty;
    public List<string> ChangedNumbers { get; set; } = new();
}

public class AiPageFocus
{
    public int PairIndex { get; set; }
    public int? PageAfs1 { get; set; }
    public int? PageAfs2 { get; set; }
    public string Priority { get; set; } = string.Empty;
    public string Why { get; set; } = string.Empty;
    public List<int> FocusDiffOrdinals { get; set; } = new();
}
