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
                            SlideIndices = ResolveSlideIndices(slides, show)
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

    private static List<int> ResolveSlideIndices(PowerPoint.Slides slides, PowerPoint.NamedSlideShow show)
    {
        int slideCount = show.Count;
        var indices = new List<int>(slideCount);

        // NamedSlideShow.SlideIDs is declared as System.Object on the typed PIA. At runtime it
        // returns a 0-based System.Object[] whose element 0 is an unused placeholder, matching
        // VBA's documented 1-based access to this array ("For i = 1 To UBound(idArray)"); the
        // real IDs are elements 1..Count (confirmed empirically: a 2-slide show returned an
        // array of {0, id1, id2}).
        object slideIdsObject = show.SlideIDs;
        if (slideIdsObject is not Array slideIdArray || slideIdArray.Length <= slideCount)
        {
            return indices;
        }

        for (int i = 1; i <= slideCount; i++)
        {
            int slideId = Convert.ToInt32(slideIdArray.GetValue(i), System.Globalization.CultureInfo.InvariantCulture);

            PowerPoint.Slide? slide = null;
            try
            {
                // The slide was deleted after the custom show was created; PowerPoint keeps the
                // stale ID in the show. FindBySlideID has been observed to both return null and
                // throw a COMException for a stale ID with no matching slide (confirmed by a
                // real-COM regression test), so both are treated the same way: the slide is
                // omitted, since either way there is nothing left to resolve it to.
                slide = slides.FindBySlideID(slideId);
                if (slide is not null)
                {
                    indices.Add(slide.SlideIndex);
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
            }
            finally
            {
                if (slide != null) ComUtilities.Release(ref slide);
            }
        }

        return indices;
    }
}
