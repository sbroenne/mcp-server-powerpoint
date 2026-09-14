namespace Sbroenne.PowerPointMcp.Core.Tests;

/// <summary>
/// Acquires the same named lock the production code uses to serialize PowerPoint's shared
/// Format Painter clipboard, so tests can hold it on behalf of another session or process.
/// </summary>
internal static class FormattingClipboardTestLock
{
    internal static Mutex Create() => new(
        "Sbroenne.PowerPointMcp.ShapeFormattingClipboard",
        new NamedWaitHandleOptions { CurrentUserOnly = true, CurrentSessionOnly = false });

    /// <summary>
    /// Holds the lock until <paramref name="release"/> is signalled. This runs on a raw background
    /// thread, where an escaping exception would terminate the whole test host instead of failing
    /// the test, so problems are reported back through <paramref name="failure"/>.
    /// </summary>
    internal static void Hold(
        ManualResetEventSlim acquired,
        ManualResetEventSlim release,
        ref Exception? failure)
    {
        bool held = false;
        Mutex? formattingClipboardMutex = null;
        try
        {
            formattingClipboardMutex = Create();
            held = formattingClipboardMutex.WaitOne(TimeSpan.FromSeconds(10));
            if (!held)
            {
                failure = new InvalidOperationException(
                    "The test could not take the formatting clipboard lock it needs to hold.");
            }
        }
        catch (Exception ex)
        {
            failure = ex;
        }
        finally
        {
            acquired.Set();
        }

        try
        {
            release.Wait();
            if (held)
            {
                formattingClipboardMutex!.ReleaseMutex();
            }
        }
        catch (Exception ex)
        {
            failure ??= ex;
        }
        finally
        {
            formattingClipboardMutex?.Dispose();
        }
    }
}
