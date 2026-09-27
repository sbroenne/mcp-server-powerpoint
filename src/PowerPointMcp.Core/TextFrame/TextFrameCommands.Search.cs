extern alias OfficeInterop;

using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.TextFrame;

public sealed partial class TextFrameCommands
{
    /// <inheritdoc/>
    public TextFrameOperationResult ReplaceText(IPresentationBatch batch, int slideIndex, int shapeIndex,
        string findWhat, string replaceWhat, bool matchCase = false, bool wholeWords = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(replaceWhat);

        return SearchText(batch, slideIndex, shapeIndex, findWhat, replaceWhat, matchCase, wholeWords);
    }

    /// <inheritdoc/>
    public TextFrameOperationResult FindText(IPresentationBatch batch, int slideIndex, int shapeIndex,
        string findWhat, bool matchCase = false, bool wholeWords = false)
    {
        ArgumentNullException.ThrowIfNull(batch);
        return SearchText(batch, slideIndex, shapeIndex, findWhat, null, matchCase, wholeWords);
    }

    private static TextFrameOperationResult SearchText(IPresentationBatch batch, int slideIndex, int shapeIndex,
        string findWhat, string? replaceWhat, bool matchCase, bool wholeWords)
    {
        if (string.IsNullOrEmpty(findWhat))
        {
            return new TextFrameOperationResult { ErrorMessage = "findWhat must not be empty." };
        }

        return batch.Execute((ctx, ct) =>
        {
            PowerPoint.Slides? slides = null;
            PowerPoint.Slide? slide = null;
            PowerPoint.Shapes? shapes = null;
            PowerPoint.Shape? shape = null;
            PowerPoint.TextFrame? frame = null;
            PowerPoint.TextRange? range = null;
            try
            {
                slides = ctx.Presentation.Slides;
                if (slideIndex < 1 || slideIndex > slides.Count)
                {
                    return new TextFrameOperationResult { ErrorMessage = $"Slide index {slideIndex} is out of range (1-{slides.Count})." };
                }

                slide = slides[slideIndex];
                shapes = slide.Shapes;
                if (shapeIndex < 1 || shapeIndex > shapes.Count)
                {
                    return new TextFrameOperationResult { ErrorMessage = $"Shape index {shapeIndex} is out of range (1-{shapes.Count})." };
                }

                shape = shapes[shapeIndex];
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    return new TextFrameOperationResult { ErrorMessage = "The selected shape has no text frame." };
                }

                frame = shape.TextFrame;
                range = frame.TextRange;
                var matches = CollectMatches(range, findWhat, matchCase, wholeWords, ct);
                if (replaceWhat is null)
                {
                    return new TextFrameOperationResult { Success = true, MatchCount = matches.Count, Matches = matches };
                }

                for (int matchIndex = matches.Count - 1; matchIndex >= 0; matchIndex--)
                {
                    ct.ThrowIfCancellationRequested();
                    PowerPoint.TextRange? match = null;
                    try
                    {
                        var position = matches[matchIndex];
                        match = range.Characters(position.Start, position.Length);
                        match.Text = replaceWhat;
                    }
                    finally
                    {
                        ComUtilities.Release(ref match!);
                    }
                }

                return new TextFrameOperationResult { Success = true, ReplacementCount = matches.Count };
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref frame!);
                ComUtilities.Release(ref shape!);
                ComUtilities.Release(ref shapes!);
                ComUtilities.Release(ref slide!);
                ComUtilities.Release(ref slides!);
            }
        });
    }

    private static List<TextMatch> CollectMatches(PowerPoint.TextRange range,
        string findWhat, bool matchCase, bool wholeWords, CancellationToken cancellationToken)
    {
        var matches = new List<TextMatch>();
        int after = 0;
        int length = range.Length;
        while (after < length)
        {
            cancellationToken.ThrowIfCancellationRequested();
            PowerPoint.TextRange? match = null;
            try
            {
                match = range.Find(findWhat, after,
                    matchCase ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse,
                    wholeWords ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse);
                if (match is null) break;

                int start = match.Start;
                int matchLength = match.Length;
                if (start <= after || matchLength <= 0)
                {
                    throw new InvalidOperationException("PowerPoint returned a non-advancing text match.");
                }

                matches.Add(new TextMatch { Start = start, Length = matchLength, Text = match.Text });
                after = start + matchLength - 1;
            }
            finally
            {
                ComUtilities.Release(ref match!);
            }
        }

        return matches;
    }
}