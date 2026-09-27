using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.CustomShow;

/// <summary>
/// Named custom slide shows: curated, ordered subsets of a presentation's slides.
/// </summary>
[ServiceCategory("customshow", "CustomShow")]
[McpTool("customshow", Title = "Custom Shows", Destructive = true, Category = "content",
    Description = "Create, list, and delete named custom slide shows: curated, ordered subsets of a "
    + "presentation's slides used to reuse one deck for different audiences. A custom show may repeat "
    + "a slide and does not need to include every slide.")]
public interface ICustomShowCommands
{
    /// <summary>Lists all custom shows in the presentation, in collection order.</summary>
    CustomShowOperationResult List(IPresentationBatch batch);

    /// <summary>
    /// Creates a named custom show from an ordered list of 1-based slide indices. Slides may repeat
    /// and the show does not need to include every slide. Fails if a custom show with the same name
    /// already exists or if any slide index is out of range.
    /// </summary>
    CustomShowOperationResult Create(
        IPresentationBatch batch,
        string name,
        IReadOnlyList<int> slideIndices);

    /// <summary>Deletes the custom show with the given name.</summary>
    CustomShowOperationResult Delete(IPresentationBatch batch, string name);
}

