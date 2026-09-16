using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.CustomShow;

/// <inheritdoc cref="ICustomShowCommands"/>
public sealed class CustomShowCommands : ICustomShowCommands
{
    /// <inheritdoc/>
    public CustomShowOperationResult List(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.SlideShowSettings? settings = null;
            PowerPoint.NamedSlideShows? shows = null;
            PowerPoint.Slides? slides = null;
            try
            {
                settings = ctx.Presentation.SlideShowSettings;
                shows = settings.NamedSlideShows;
                slides = ctx.Presentation.Slides;
                var slideIndexById = BuildSlideIndexById(slides);

                int count = shows.Count;
                var entries = new List<CustomShowOperationResult.CustomShowEntry>(count);
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.NamedSlideShow? show = null;
                    try
                    {
                        show = shows[i];
                        entries.Add(new CustomShowOperationResult.CustomShowEntry
                        {
                            Index = i,
                            Name = show.Name,
                            SlideIndices = ResolveSlideIndices(slideIndexById, show)
                        });
                    }
                    finally
                    {
                        if (show != null) ComUtilities.Release(ref show);
                    }
                }

                return new CustomShowOperationResult { Success = true, Shows = entries };
            }
            finally
            {
                if (slides != null) ComUtilities.Release(ref slides);
                if (shows != null) ComUtilities.Release(ref shows);
                if (settings != null) ComUtilities.Release(ref settings);
            }
        });
    }

    /// <inheritdoc/>
    public CustomShowOperationResult Create(IPresentationBatch batch, string name, IReadOnlyList<int> slideIndices)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(slideIndices);

        if (string.IsNullOrWhiteSpace(name))
        {
            return new CustomShowOperationResult
            {
                Success = false,
                ErrorMessage = "A custom show name is required."
            };
        }

        if (slideIndices.Count == 0)
        {
            return new CustomShowOperationResult
            {
                Success = false,
                ErrorMessage = "slideIndices must contain at least one slide index."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.SlideShowSettings? settings = null;
            PowerPoint.NamedSlideShows? shows = null;
            PowerPoint.Slides? slides = null;
            PowerPoint.NamedSlideShow? created = null;
            try
            {
                slides = ctx.Presentation.Slides;
                int slideCount = slides.Count;
                foreach (int index in slideIndices)
                {
                    if (index < 1 || index > slideCount)
                    {
                        return new CustomShowOperationResult
                        {
                            Success = false,
                            ErrorMessage = $"Slide index {index} is out of range (presentation has {slideCount} slide(s))."
                        };
                    }
                }

                settings = ctx.Presentation.SlideShowSettings;
                shows = settings.NamedSlideShows;

                if (FindByName(shows, name) > 0)
                {
                    return new CustomShowOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"A custom show named '{name}' already exists."
                    };
                }

                var slideIds = new int[slideIndices.Count];
                for (int i = 0; i < slideIndices.Count; i++)
                {
                    PowerPoint.Slide? slide = null;
                    try
                    {
                        slide = slides[slideIndices[i]];
                        slideIds[i] = slide.SlideID;
                    }
                    finally
                    {
                        if (slide != null) ComUtilities.Release(ref slide);
                    }
                }

                created = shows.Add(name, slideIds);

                return new CustomShowOperationResult
                {
                    Success = true,
                    Name = name,
                    SlideIndices = slideIndices.ToArray()
                };
            }
            finally
            {
                if (created != null) ComUtilities.Release(ref created);
                if (shows != null) ComUtilities.Release(ref shows);
                if (settings != null) ComUtilities.Release(ref settings);
                if (slides != null) ComUtilities.Release(ref slides);
            }
        });
    }

    /// <inheritdoc/>
    public CustomShowOperationResult Delete(IPresentationBatch batch, string name)
    {
        ArgumentNullException.ThrowIfNull(batch);

        if (string.IsNullOrWhiteSpace(name))
        {
            return new CustomShowOperationResult
            {
                Success = false,
                ErrorMessage = "A custom show name is required."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.SlideShowSettings? settings = null;
            PowerPoint.NamedSlideShows? shows = null;
            PowerPoint.NamedSlideShow? match = null;
            try
            {
                settings = ctx.Presentation.SlideShowSettings;
                shows = settings.NamedSlideShows;

                int matchIndex = FindByName(shows, name);
                if (matchIndex < 0)
                {
                    return new CustomShowOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"No custom show named '{name}' was found."
                    };
                }

                match = shows[matchIndex];
                match.Delete();

                return new CustomShowOperationResult { Success = true, Name = name };
            }
            finally
            {
                if (match != null) ComUtilities.Release(ref match);
                if (shows != null) ComUtilities.Release(ref shows);
                if (settings != null) ComUtilities.Release(ref settings);
            }
        });
    }

    /// <summary>Returns the 1-based index of the show named <paramref name="name"/>, or -1 if not found.</summary>
    private static int FindByName(PowerPoint.NamedSlideShows shows, string name)
    {
        for (int i = 1; i <= shows.Count; i++)
        {
            PowerPoint.NamedSlideShow? candidate = null;
            try
            {
                candidate = shows[i];
                // PowerPoint's NamedSlideShows collection treats names case-insensitively: creating
                // "Demo" when "demo" already exists reaches Add() and fails at the COM layer instead
                // of the documented duplicate-name result, and delete-by-name would not find an
                // existing show that differs only in casing.
                if (string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }
            finally
            {
                if (candidate != null) ComUtilities.Release(ref candidate);
            }
        }

        return -1;
    }

    /// <summary>Maps every current slide's SlideID to its 1-based index via one pass over <paramref name="slides"/>.</summary>
    private static Dictionary<int, int> BuildSlideIndexById(PowerPoint.Slides slides)
    {
        int count = slides.Count;
        var map = new Dictionary<int, int>(count);
        for (int i = 1; i <= count; i++)
        {
            PowerPoint.Slide? slide = null;
            try
            {
                slide = slides[i];
                map[slide.SlideID] = i;
            }
            finally
            {
                if (slide != null) ComUtilities.Release(ref slide);
            }
        }

        return map;
    }

    private static List<int> ResolveSlideIndices(Dictionary<int, int> slideIndexById, PowerPoint.NamedSlideShow show)
    {
        int slideCount = show.Count;
        var indices = new List<int>(slideCount);

        // NamedSlideShow.SlideIDs is declared as System.Object on the typed PIA. At runtime it has
        // been observed to return a 0-based System.Object[] of length Count + 1 whose element 0 is
        // an unused placeholder, matching VBA's documented 1-based access to this array ("For i = 1
        // To UBound(idArray)"): a 2-slide show returned {0, id1, id2}. Rather than assuming that
        // exact shape, only the array's own bounds are trusted: the real IDs are the last
        // slideCount elements, whatever the array's lower bound turns out to be for a given
        // PowerPoint/interop marshaling variant.
        if (slideCount == 0)
        {
            return indices;
        }

        object slideIdsObject = show.SlideIDs;
        if (slideIdsObject is not Array slideIdArray || slideIdArray.Length < slideCount)
        {
            // A non-array or undersized SlideIDs value for a non-empty show is an unexpected COM
            // result, not the documented deleted-slide case (a stale ID is still present in the
            // array and is filtered out below via the dictionary lookup). Throwing here lets
            // batch.Execute() surface it as a failure instead of List() silently reporting
            // Success=true with empty or partial slide indices.
            throw new InvalidOperationException(
                $"Custom show '{show.Name}' reported {slideCount} slide(s) but PowerPoint returned an " +
                "unexpected SlideIDs value that was not an array of at least that length.");
        }

        int upperBound = slideIdArray.GetUpperBound(0);
        int firstRealIndex = upperBound - slideCount + 1;
        for (int i = firstRealIndex; i <= upperBound; i++)
        {
            int slideId = Convert.ToInt32(slideIdArray.GetValue(i), System.Globalization.CultureInfo.InvariantCulture);

            // The slide was deleted after the custom show was created; PowerPoint keeps the
            // stale ID in the show. Rather than asking PowerPoint to resolve it (FindBySlideID's
            // failure behavior for a missing ID has been observed to vary - both a null return
            // and a COMException - and is not a safe signal to distinguish from a genuine COM
            // failure), the current slide list is looked up once per List() call and checked
            // here with a plain, non-throwing dictionary lookup.
            if (slideIndexById.TryGetValue(slideId, out int slideIndex))
            {
                indices.Add(slideIndex);
            }
        }

        return indices;
    }
}
