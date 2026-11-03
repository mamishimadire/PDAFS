namespace AfsPdfComparison.Models;

public class AiBackendStatusSnapshot
{
    public bool BoostActive { get; set; }
    public string Summary { get; set; } = string.Empty;
    public List<AiBackendProviderStatus> Providers { get; set; } = new();
}

public class AiBackendProviderStatus
{
    public string Key { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Model { get; set; } = string.Empty;
    public bool Active { get; set; }
    public string Reason { get; set; } = string.Empty;
}
