using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Notes;

/// <inheritdoc cref="INotesCommands"/>
public sealed class NotesCommands : INotesCommands
{
    /// <inheritdoc/>
    public NotesOperationResult SetNotesText(IPresentationBatch batch, int slideIndex, string text)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(text);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateSlideIndex(ctx, slideIndex);
            if (validation is not null) return validation;

            dynamic notesTextShape = FindNotesTextShape(ctx, slideIndex);
            notesTextShape.TextFrame.TextRange.Text = text;

            return new NotesOperationResult { Success = true, NotesText = text };
        });
    }

    /// <inheritdoc/>
    public NotesOperationResult GetNotesText(IPresentationBatch batch, int slideIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateSlideIndex(ctx, slideIndex);
            if (validation is not null) return validation;

            dynamic notesTextShape = FindNotesTextShape(ctx, slideIndex);
            string text = (string)notesTextShape.TextFrame.TextRange.Text;

            return new NotesOperationResult { Success = true, NotesText = text };
        });
    }

    /// <summary>
    /// Finds the speaker-notes body text placeholder shape on a slide's notes page.
    /// </summary>
    /// <remarks>
    /// The notes page conventionally has exactly two placeholders: index 1 is the slide
    /// image thumbnail, index 2 is the notes body text — this is the standard PowerPoint
    /// notes master layout. We defensively scan for the shape whose PlaceholderFormat.Type
    /// is ppPlaceholderBody rather than hard-coding index 2, in case a custom notes master
    /// reorders placeholders.
    /// </remarks>
    private static dynamic FindNotesTextShape(PresentationContext ctx, int slideIndex)
    {
        dynamic notesPage = ctx.Presentation.Slides[slideIndex].NotesPage;
        dynamic shapes = notesPage.Shapes;
        int count = (int)shapes.Count;

        for (int i = 1; i <= count; i++)
        {
            dynamic shape = shapes[i];
            bool isPlaceholder = (int)shape.Type == 14; // msoPlaceholder
            if (!isPlaceholder) continue;

            int placeholderType = (int)shape.PlaceholderFormat.Type;
            if (placeholderType == 2) // PpPlaceholderType.ppPlaceholderBody
            {
                return shape;
            }
        }

        throw new InvalidOperationException(
            $"Could not find the notes body text placeholder on slide {slideIndex}'s notes page.");
    }

    private static NotesOperationResult? ValidateSlideIndex(PresentationContext ctx, int slideIndex)
    {
        int slideCount = ctx.Presentation.Slides.Count;
        if (slideIndex < 1 || slideIndex > slideCount)
        {
            return new NotesOperationResult
            {
                Success = false,
                ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
            };
        }
        return null;
    }
}
