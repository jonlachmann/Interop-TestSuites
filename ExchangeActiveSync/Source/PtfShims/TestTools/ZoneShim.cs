// Minimal stand-in for legacy System.Security.Policy.Zone and SecurityZone used by ActiveSyncClient.
// Keeps namespace identical so existing code compiles without changes.
namespace System.Security.Policy;

public enum SecurityZone
{
    MyComputer = 0,
    Intranet = 1,
    Trusted = 2,
    Internet = 3,
    Untrusted = 4,
    NoZone = -1
}

public sealed class Zone
{
    private Zone(SecurityZone zone)
    {
        SecurityZone = zone;
    }

    public SecurityZone SecurityZone { get; }

    public static Zone CreateFromUrl(string url)
    {
        // Basic heuristic: treat file/localhost as intranet, otherwise internet.
        if (!string.IsNullOrEmpty(url) && (url.StartsWith("file:", StringComparison.OrdinalIgnoreCase) || url.Contains("localhost")))
        {
            return new Zone(SecurityZone.Intranet);
        }

        return new Zone(SecurityZone.Internet);
    }
}