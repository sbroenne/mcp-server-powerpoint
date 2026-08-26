using System.Security.Cryptography;
using System.Text;

namespace Sbroenne.PowerPointMcp.CLI.Infrastructure;

internal static class DaemonPipeIdentity
{
    internal static string GetStableKey(string? pipeName)
    {
        var canonicalName = string.IsNullOrWhiteSpace(pipeName)
            ? DaemonAutoStart.GetPipeName()
            : pipeName.Trim();
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(canonicalName.ToUpperInvariant()));
        return Convert.ToHexString(bytes).ToLowerInvariant();
    }

    internal static string GetHash(string pipeName) => GetStableKey(pipeName);
}
