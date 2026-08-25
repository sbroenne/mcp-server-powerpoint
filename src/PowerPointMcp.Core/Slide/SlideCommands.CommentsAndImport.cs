using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Slide;

public sealed partial class SlideCommands
{
    /// <inheritdoc/>
    public SlideOperationResult ListComments(IPresentationBatch batch, int slideIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (validation is not null) return validation;

            PowerPoint.Comments? comments = null;
            try
            {
                comments = ctx.Presentation.Slides[slideIndex].Comments;
                var items = new List<SlideCommentInfo>(comments.Count);
                for (int index = 1; index <= comments.Count; index++)
                {
                    PowerPoint.Comment? comment = null;
                    try
                    {
                        comment = comments[index];
                        items.Add(new SlideCommentInfo
                        {
                            CommentIndex = index,
                            Author = comment.Author,
                            Initials = comment.AuthorInitials,
                            Text = comment.Text,
                            Left = comment.Left,
                            Top = comment.Top
                        });
                    }
                    finally
                    {
                        if (comment is not null)
                        {
                            ComUtilities.Release(ref comment);
                        }
                    }
                }

                return new SlideOperationResult
                {
                    Success = true,
                    SlideIndex = slideIndex,
                    Comments = items,
                    CommentCount = items.Count
                };
            }
            finally
            {
                if (comments is not null)
                {
                    ComUtilities.Release(ref comments);
                }
            }
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult AddComment(
        IPresentationBatch batch,
        int slideIndex,
        string author,
        string initials,
        string text,
        float left = 0f,
        float top = 0f)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(author);
        ArgumentNullException.ThrowIfNull(initials);
        ArgumentNullException.ThrowIfNull(text);

        if (string.IsNullOrWhiteSpace(author) ||
            string.IsNullOrWhiteSpace(initials) ||
            string.IsNullOrWhiteSpace(text))
        {
            return new SlideOperationResult
            {
                ErrorMessage = "Comment author, initials, and text must not be empty."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (validation is not null) return validation;

            PowerPoint.Comments? comments = null;
            PowerPoint.Comment? comment = null;
            try
            {
                comments = ctx.Presentation.Slides[slideIndex].Comments;
                comment = comments.Add(left, top, author, initials, text);
                return new SlideOperationResult
                {
                    Success = true,
                    SlideIndex = slideIndex,
                    CommentCount = comments.Count
                };
            }
            finally
            {
                if (comment is not null)
                {
                    ComUtilities.Release(ref comment);
                }
                if (comments is not null)
                {
                    ComUtilities.Release(ref comments);
                }
            }
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult DeleteComment(
        IPresentationBatch batch,
        int slideIndex,
        int commentIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (validation is not null) return validation;

            PowerPoint.Comments? comments = null;
            PowerPoint.Comment? comment = null;
            try
            {
                comments = ctx.Presentation.Slides[slideIndex].Comments;
                if (commentIndex < 1 || commentIndex > comments.Count)
                {
                    return new SlideOperationResult
                    {
                        ErrorMessage = $"Comment index {commentIndex} is out of range. The slide has {comments.Count} comment(s).",
                        SlideIndex = slideIndex,
                        CommentCount = comments.Count
                    };
                }

                comment = comments[commentIndex];
                comment.Delete();
                return new SlideOperationResult
                {
                    Success = true,
                    SlideIndex = slideIndex,
                    CommentCount = comments.Count
                };
            }
            finally
            {
                if (comment is not null)
                {
                    ComUtilities.Release(ref comment);
                }
                if (comments is not null)
                {
                    ComUtilities.Release(ref comments);
                }
            }
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult ClearComments(IPresentationBatch batch, int slideIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (validation is not null) return validation;

            PowerPoint.Comments? comments = null;
            try
            {
                comments = ctx.Presentation.Slides[slideIndex].Comments;
                for (int index = comments.Count; index >= 1; index--)
                {
                    PowerPoint.Comment? comment = null;
                    try
                    {
                        comment = comments[index];
                        comment.Delete();
                    }
                    finally
                    {
                        if (comment is not null)
                        {
                            ComUtilities.Release(ref comment);
                        }
                    }
                }

                return new SlideOperationResult
                {
                    Success = true,
                    SlideIndex = slideIndex,
                    CommentCount = 0
                };
            }
            finally
            {
                if (comments is not null)
                {
                    ComUtilities.Release(ref comments);
                }
            }
        });
    }

    /// <inheritdoc/>
    public SlideOperationResult ImportFromFile(
        IPresentationBatch batch,
        string sourceFilePath,
        int destinationSlideIndex,
        int sourceStartSlide = 1,
        int? sourceEndSlide = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(sourceFilePath);

        if (string.IsNullOrWhiteSpace(sourceFilePath))
        {
            return new SlideOperationResult { ErrorMessage = "Source file path must not be empty." };
        }

        string fullSourcePath;
        try
        {
            fullSourcePath = Path.GetFullPath(sourceFilePath);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return new SlideOperationResult { ErrorMessage = $"Invalid source file path: {ex.Message}" };
        }

        if (!File.Exists(fullSourcePath))
        {
            return new SlideOperationResult { ErrorMessage = $"Source presentation not found: {fullSourcePath}." };
        }

        if (sourceStartSlide < 1 ||
            (sourceEndSlide.HasValue && sourceEndSlide.Value < sourceStartSlide))
        {
            return new SlideOperationResult
            {
                ErrorMessage = "Source slide range must be 1-based and end at or after its start."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            int destinationCount = ctx.Presentation.Slides.Count;
            var validation = ValidateSlideIndex(destinationCount, destinationSlideIndex);
            if (validation is not null) return validation;

            int sourceSlideCount;
            PowerPoint.Presentations? presentations = null;
            PowerPoint.Presentation? sourcePresentation = null;
            try
            {
                presentations = ctx.Presentation.Application.Presentations;
                // PIA gap: Presentations.Open uses Office.MsoTriState, unavailable without office.dll.
                sourcePresentation = ((dynamic)presentations).Open(
                    fullSourcePath,
                    MsoTrue,
                    MsoFalse,
                    MsoFalse);
                sourceSlideCount = sourcePresentation.Slides.Count;
            }
            finally
            {
                if (sourcePresentation is not null)
                {
                    sourcePresentation.Close();
                    ComUtilities.Release(ref sourcePresentation);
                }
                if (presentations is not null)
                {
                    ComUtilities.Release(ref presentations);
                }
            }

            int sourceEnd = sourceEndSlide ?? sourceSlideCount;
            if (sourceStartSlide > sourceSlideCount || sourceEnd > sourceSlideCount)
            {
                return new SlideOperationResult
                {
                    ErrorMessage = $"Source slide range {sourceStartSlide}-{sourceEnd} is out of range. The source presentation has {sourceSlideCount} slide(s).",
                    SlideCount = destinationCount
                };
            }

            int importedCount = ctx.Presentation.Slides.InsertFromFile(
                fullSourcePath,
                destinationSlideIndex,
                sourceStartSlide,
                sourceEnd);
            int[] importedIndexes = Enumerable.Range(destinationSlideIndex + 1, importedCount).ToArray();

            return new SlideOperationResult
            {
                Success = true,
                SlideCount = ctx.Presentation.Slides.Count,
                ImportedSlideCount = importedCount,
                ImportedSlideIndexes = importedIndexes
            };
        });
    }

    private static SlideOperationResult? ValidateSlideIndex(int slideCount, int slideIndex)
    {
        if (slideIndex < 1 || slideIndex > slideCount)
        {
            return new SlideOperationResult
            {
                ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount}).",
                SlideCount = slideCount
            };
        }

        return null;
    }
}
