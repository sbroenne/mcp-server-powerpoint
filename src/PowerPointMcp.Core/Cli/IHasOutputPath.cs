namespace Sbroenne.PowerPointMcp.Core.Cli;

/// <summary>
/// Implemented by generated CLI settings classes that expose the shared <c>-o|--output</c>
/// option. Lets <c>ServiceCommandBase&lt;TSettings&gt;</c> read <c>OutputPath</c> without
/// reflection (<see cref="System.Type.GetProperty(string)"/>), which is required for the
/// standalone-exe release build to be trim-safe (see the release workflow's
/// <c>PublishTrimmed</c> exe publish).
/// </summary>
public interface IHasOutputPath
{
    /// <summary>Gets the file path to write command output to, or <see langword="null"/> to write to stdout.</summary>
    string? OutputPath { get; }
}
