using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.PageSetup;

/// <inheritdoc cref="IPageSetupCommands"/>
public sealed class PageSetupCommands : IPageSetupCommands
{
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

    /// <inheritdoc/>
    public PageSetupOperationResult GetSettings(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) => ReadSettings(ctx.Presentation));
    }

    /// <inheritdoc/>
    public PageSetupOperationResult SetSize(IPresentationBatch batch, float width, float height)
    {
        ArgumentNullException.ThrowIfNull(batch);

        if (!float.IsFinite(width) || width <= 0)
        {
            return new PageSetupOperationResult
            {
                Success = false,
                ErrorMessage = $"Width {width} is invalid; width must be greater than zero."
            };
        }

        if (!float.IsFinite(height) || height <= 0)
        {
            return new PageSetupOperationResult
            {
                Success = false,
                ErrorMessage = $"Height {height} is invalid; height must be greater than zero."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.PageSetup? pageSetup = null;
            try
            {
                pageSetup = ctx.Presentation.PageSetup;
                pageSetup.SlideWidth = width;
                pageSetup.SlideHeight = height;

                return CreateSettingsResult(pageSetup);
            }
            finally
            {
                if (pageSetup != null)
                {
                    ComUtilities.Release(ref pageSetup);
                }
            }
        });
    }

    /// <inheritdoc/>
    public PageSetupOperationResult SetFirstSlideNumber(IPresentationBatch batch, int firstSlideNumber)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.PageSetup? pageSetup = null;
            try
            {
                pageSetup = ctx.Presentation.PageSetup;
                pageSetup.FirstSlideNumber = firstSlideNumber;

                return CreateSettingsResult(pageSetup);
            }
            finally
            {
                if (pageSetup != null)
                {
                    ComUtilities.Release(ref pageSetup);
                }
            }
        });
    }

    /// <inheritdoc/>
    public PageSetupOperationResult GetFooter(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) => ReadFooter(ctx.Presentation));
    }

    /// <inheritdoc/>
    public PageSetupOperationResult SetFooter(
        IPresentationBatch batch,
        string? footerText = null,
        bool? showFooter = null,
        bool? showSlideNumber = null,
        bool? showDateTime = null,
        string? dateTimeMode = null,
        string? fixedDateTimeText = null,
        bool? showOnTitleSlide = null)
    {
        ArgumentNullException.ThrowIfNull(batch);

        string? normalizedDateTimeMode = dateTimeMode?.Trim().ToLowerInvariant();
        if (normalizedDateTimeMode is not null and not "automatic" and not "fixed")
        {
            return new PageSetupOperationResult
            {
                Success = false,
                ErrorMessage = $"Date/time mode '{dateTimeMode}' is invalid; it must be 'automatic' or 'fixed'."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Master? master = null;
            PowerPoint.HeadersFooters? headersFooters = null;
            PowerPoint.HeaderFooter? footer = null;
            PowerPoint.HeaderFooter? slideNumber = null;
            PowerPoint.HeaderFooter? dateAndTime = null;
            try
            {
                master = GetSlideMaster(ctx.Presentation);
                headersFooters = master.HeadersFooters;
                footer = headersFooters.Footer;
                slideNumber = headersFooters.SlideNumber;
                dateAndTime = headersFooters.DateAndTime;

                if (footerText != null)
                {
                    footer.Text = footerText;
                    SetTriStateProperty(footer, "Visible", true);
                }

                if (showFooter.HasValue)
                {
                    SetTriStateProperty(footer, "Visible", showFooter.Value);
                }

                if (showSlideNumber.HasValue)
                {
                    SetTriStateProperty(slideNumber, "Visible", showSlideNumber.Value);
                }

                if (showDateTime.HasValue)
                {
                    SetTriStateProperty(dateAndTime, "Visible", showDateTime.Value);
                }

                if (normalizedDateTimeMode != null)
                {
                    SetTriStateProperty(dateAndTime, "UseFormat", normalizedDateTimeMode == "automatic");
                }

                if (fixedDateTimeText != null)
                {
                    dateAndTime.Text = fixedDateTimeText;
                }

                if (showOnTitleSlide.HasValue)
                {
                    SetTriStateProperty(headersFooters, "DisplayOnTitleSlide", showOnTitleSlide.Value);
                }

                return CreateFooterResult(headersFooters, footer, slideNumber, dateAndTime);
            }
            finally
            {
                if (dateAndTime != null) ComUtilities.Release(ref dateAndTime);
                if (slideNumber != null) ComUtilities.Release(ref slideNumber);
                if (footer != null) ComUtilities.Release(ref footer);
                if (headersFooters != null) ComUtilities.Release(ref headersFooters);
                if (master != null) ComUtilities.Release(ref master);
            }
        });
    }

    private static PageSetupOperationResult ReadSettings(PowerPoint.Presentation presentation)
    {
        PowerPoint.PageSetup? pageSetup = null;
        try
        {
            pageSetup = presentation.PageSetup;
            return CreateSettingsResult(pageSetup);
        }
        finally
        {
            if (pageSetup != null)
            {
                ComUtilities.Release(ref pageSetup);
            }
        }
    }

    private static PageSetupOperationResult CreateSettingsResult(PowerPoint.PageSetup pageSetup)
    {
        float width = pageSetup.SlideWidth;
        float height = pageSetup.SlideHeight;

        return new PageSetupOperationResult
        {
            Success = true,
            Width = width,
            Height = height,
            Orientation = width > height ? "landscape" : width < height ? "portrait" : "square",
            FirstSlideNumber = pageSetup.FirstSlideNumber
        };
    }

    private static PageSetupOperationResult ReadFooter(PowerPoint.Presentation presentation)
    {
        PowerPoint.Master? master = null;
        PowerPoint.HeadersFooters? headersFooters = null;
        PowerPoint.HeaderFooter? footer = null;
        PowerPoint.HeaderFooter? slideNumber = null;
        PowerPoint.HeaderFooter? dateAndTime = null;
        try
        {
            master = GetSlideMaster(presentation);
            headersFooters = master.HeadersFooters;
            footer = headersFooters.Footer;
            slideNumber = headersFooters.SlideNumber;
            dateAndTime = headersFooters.DateAndTime;

            return CreateFooterResult(headersFooters, footer, slideNumber, dateAndTime);
        }
        finally
        {
            if (dateAndTime != null) ComUtilities.Release(ref dateAndTime);
            if (slideNumber != null) ComUtilities.Release(ref slideNumber);
            if (footer != null) ComUtilities.Release(ref footer);
            if (headersFooters != null) ComUtilities.Release(ref headersFooters);
            if (master != null) ComUtilities.Release(ref master);
        }
    }

    private static PageSetupOperationResult CreateFooterResult(
        PowerPoint.HeadersFooters headersFooters,
        PowerPoint.HeaderFooter footer,
        PowerPoint.HeaderFooter slideNumber,
        PowerPoint.HeaderFooter dateAndTime)
    {
        bool useAutomaticDateTime = GetTriStateProperty(dateAndTime, "UseFormat");

        return new PageSetupOperationResult
        {
            Success = true,
            FooterText = footer.Text,
            ShowFooter = GetTriStateProperty(footer, "Visible"),
            ShowSlideNumber = GetTriStateProperty(slideNumber, "Visible"),
            ShowDateTime = GetTriStateProperty(dateAndTime, "Visible"),
            DateTimeMode = useAutomaticDateTime ? "automatic" : "fixed",
            FixedDateTimeText = useAutomaticDateTime ? null : dateAndTime.Text,
            ShowOnTitleSlide = GetTriStateProperty(headersFooters, "DisplayOnTitleSlide")
        };
    }

    private static PowerPoint.Master GetSlideMaster(PowerPoint.Presentation presentation)
    {
        // The embedded NoPIA Presentation.SlideMaster getter hangs in live PowerPoint. Follow the
        // established MasterCommands pattern and late-bind only this getter.
        dynamic presentationDispatch = presentation;
        return (PowerPoint.Master)presentationDispatch.SlideMaster;
    }

    private static bool GetTriStateProperty(object owner, string propertyName)
    {
        // These setters/getters are Office.MsoTriState-typed and office.dll is intentionally absent.
        dynamic dispatch = owner;
        return propertyName switch
        {
            "Visible" => (int)dispatch.Visible == MsoTrue,
            "UseFormat" => (int)dispatch.UseFormat == MsoTrue,
            "DisplayOnTitleSlide" => (int)dispatch.DisplayOnTitleSlide == MsoTrue,
            _ => throw new ArgumentOutOfRangeException(nameof(propertyName))
        };
    }

    private static void SetTriStateProperty(object owner, string propertyName, bool value)
    {
        // These setters are Office.MsoTriState-typed and office.dll is intentionally absent.
        dynamic dispatch = owner;
        int triState = value ? MsoTrue : MsoFalse;
        switch (propertyName)
        {
            case "Visible":
                dispatch.Visible = triState;
                break;
            case "UseFormat":
                dispatch.UseFormat = triState;
                break;
            case "DisplayOnTitleSlide":
                dispatch.DisplayOnTitleSlide = triState;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(propertyName));
        }
    }
}
