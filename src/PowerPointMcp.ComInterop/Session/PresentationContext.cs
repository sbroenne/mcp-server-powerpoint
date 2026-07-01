using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.ComInterop.Session;

/// <summary>
/// Provides access to PowerPoint COM objects for operations.
/// Simplifies passing the PowerPoint application and presentation to operations.
/// </summary>
public sealed class PresentationContext
{
    /// <summary>Creates a new PresentationContext.</summary>
    /// <param name="presentationPath">Full path to the presentation.</param>
    /// <param name="app">PowerPoint.Application COM object.</param>
    /// <param name="presentation">PowerPoint.Presentation COM object.</param>
    public PresentationContext(string presentationPath, PowerPoint.Application app, PowerPoint.Presentation presentation)
    {
        PresentationPath = presentationPath ?? throw new ArgumentNullException(nameof(presentationPath));
        App = app ?? throw new ArgumentNullException(nameof(app));
        Presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
    }

    /// <summary>Gets the full path to the presentation.</summary>
    public string PresentationPath { get; }

    /// <summary>Gets the PowerPoint.Application COM object.</summary>
    public PowerPoint.Application App { get; }

    /// <summary>Gets the PowerPoint.Presentation COM object.</summary>
    public PowerPoint.Presentation Presentation { get; }
}
