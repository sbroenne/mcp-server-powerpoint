using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Tags;

namespace Sbroenne.PowerPointMcp.Core.Presentation;

public sealed partial class PresentationCommands
{
    /// <inheritdoc/>
    public PresentationOperationResult SetTag(IPresentationBatch batch, string tagName, string tagValue)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(tagValue);

        return batch.Execute((ctx, ct) =>
            ToPresentationResult(
                TagCollectionHelper.Set(
                    ctx,
                    ctx.Presentation,
                    () => ctx.Presentation.Tags,
                    tagName,
                    tagValue),
                batch.PresentationPath));
    }

    /// <inheritdoc/>
    public PresentationOperationResult GetTag(IPresentationBatch batch, string tagName)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
            ToPresentationResult(
                TagCollectionHelper.Get(
                    ctx,
                    ctx.Presentation,
                    () => ctx.Presentation.Tags,
                    tagName),
                batch.PresentationPath));
    }

    /// <inheritdoc/>
    public PresentationOperationResult ListTags(IPresentationBatch batch)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
            ToPresentationResult(
                TagCollectionHelper.List(
                    ctx,
                    ctx.Presentation,
                    () => ctx.Presentation.Tags),
                batch.PresentationPath));
    }

    /// <inheritdoc/>
    public PresentationOperationResult DeleteTag(IPresentationBatch batch, string tagName)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
            ToPresentationResult(
                TagCollectionHelper.Delete(
                    ctx,
                    ctx.Presentation,
                    () => ctx.Presentation.Tags,
                    tagName),
                batch.PresentationPath));
    }

    private static PresentationOperationResult ToPresentationResult(
        TagCollectionResult result,
        string presentationPath) => new()
        {
            Success = result.Success,
            ErrorMessage = result.ErrorMessage,
            PresentationPath = presentationPath,
            TagName = result.Name,
            TagValue = result.Value,
            TagIndex = result.Index,
            TagCount = result.Count,
            Tags = result.Tags
        };
}
