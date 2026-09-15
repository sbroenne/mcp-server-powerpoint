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
                            SlideIndices = ResolveSlideIndices(slides, show.SlideIDs)
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
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ArgumentNullException.ThrowIfNull(slideIndices);

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
        ArgumentException.ThrowIfNullOrWhiteSpace(name);

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
                if (string.Equals(candidate.Name, name, StringComparison.Ordinal))
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

    private static IReadOnlyList<int> ResolveSlideIndices(PowerPoint.Slides slides, object slideIdsObject)
    {
        // NamedSlideShow.SlideIDs is declared as System.Object on the typed PIA (an untyped Variant
        // SAFEARRAY), so its elements are read here via boxed conversion instead of a typed cast.
        if (slideIdsObject is not Array slideIdArray)
        {
            return Array.Empty<int>();
        }

        var indices = new List<int>(slideIdArray.Length);
        foreach (var rawSlideId in slideIdArray)
        {
            int slideId = Convert.ToInt32(rawSlideId, System.Globalization.CultureInfo.InvariantCulture);
            PowerPoint.Slide? slide = null;
            try
            {
                slide = slides.FindBySlideID(slideId);
                indices.Add(slide.SlideIndex);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // The slide was deleted after the custom show was created; PowerPoint keeps the
                // stale ID in the show, but there is no slide left to resolve it to. Omit it.
            }
            finally
            {
                if (slide != null) ComUtilities.Release(ref slide);
            }
        }

        return indices;
    }
}
