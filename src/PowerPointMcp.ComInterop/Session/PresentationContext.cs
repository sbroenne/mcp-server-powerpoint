using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.ComInterop.Session;

/// <summary>
/// Provides access to PowerPoint COM objects for operations.
/// Simplifies passing the PowerPoint application and presentation to operations.
/// </summary>
public sealed class PresentationContext
{
    private readonly Dictionary<object, object> _ownedComResources =
        new(ReferenceEqualityComparer.Instance);
    private readonly HashSet<object> _ownedComOwners =
        new(ReferenceEqualityComparer.Instance);
    private bool _resourcesReleased;
    private bool _disposed;

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
    public string PresentationPath { get; private set; }

    /// <summary>Gets the PowerPoint.Application COM object.</summary>
    public PowerPoint.Application App { get; }

    /// <summary>Gets the PowerPoint.Presentation COM object.</summary>
    public PowerPoint.Presentation Presentation { get; }

    /// <summary>
    /// Returns one context-owned COM resource for an owner RCW.
    /// </summary>
    /// <remarks>
    /// PowerPoint reuses some child RCWs, including <see cref="PowerPoint.Tags"/>, across property
    /// calls. Command-level release disconnects those cached proxies, so the context retains and
    /// releases them immediately before the owning presentation is closed.
    /// </remarks>
    public T GetOrAddOwnedComResource<T>(object owner, Func<T> factory) where T : class
    {
        ObjectDisposedException.ThrowIf(_resourcesReleased || _disposed, this);
        ArgumentNullException.ThrowIfNull(owner);
        ArgumentNullException.ThrowIfNull(factory);

        if (_ownedComResources.TryGetValue(owner, out object? existing))
        {
            return (T)existing;
        }

        T resource = factory();
        _ownedComResources.Add(owner, resource);
        return resource;
    }

    /// <summary>
    /// Retains the first acquisition of an owner RCW for context teardown.
    /// </summary>
    /// <returns>
    /// <see langword="true"/> when this acquisition became context-owned; otherwise
    /// <see langword="false"/> so the caller can release its repeated acquisition.
    /// </returns>
    public bool RetainOwnedComOwner(object owner)
    {
        ObjectDisposedException.ThrowIf(_resourcesReleased || _disposed, this);
        ArgumentNullException.ThrowIfNull(owner);
        return _ownedComOwners.Add(owner);
    }

    /// <summary>Updates the path after Save As without discarding context-owned COM resources.</summary>
    internal void UpdatePresentationPath(string presentationPath)
    {
        ObjectDisposedException.ThrowIf(_resourcesReleased || _disposed, this);
        PresentationPath = presentationPath ?? throw new ArgumentNullException(nameof(presentationPath));
    }

    /// <summary>Releases context-owned child COM resources before the presentation is closed.</summary>
    internal void ReleaseOwnedComResources()
    {
        if (_resourcesReleased || _disposed)
        {
            return;
        }

        _resourcesReleased = true;

        foreach (object resourceValue in _ownedComResources.Values.Reverse())
        {
            object? resource = resourceValue;
            ComUtilities.Release(ref resource);
        }
    }

    /// <summary>Releases cached owner proxies after the presentation has been closed.</summary>
    internal void ReleaseOwnedComOwners()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        foreach (object ownerValue in _ownedComOwners.Reverse())
        {
            object? owner = ownerValue;
            ComUtilities.Release(ref owner);
        }

        _ownedComOwners.Clear();
        _ownedComResources.Clear();
    }
}
