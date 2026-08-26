using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Tags;

internal static class TagCollectionHelper
{
    public static TagCollectionResult Set(
        PresentationContext context,
        object owner,
        Func<PowerPoint.Tags> acquireTags,
        string tagName,
        string value)
    {
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(owner);
        ArgumentNullException.ThrowIfNull(acquireTags);
        ArgumentNullException.ThrowIfNull(value);

        if (!TryNormalizeName(tagName, out string normalizedName, out TagCollectionResult? error))
        {
            return error;
        }

        return WithTags(context, owner, acquireTags, tags =>
        {
            tags.Add(normalizedName, value);
            int index = FindIndex(tags, normalizedName);

            return new TagCollectionResult
            {
                Success = true,
                Name = normalizedName,
                Value = tags[normalizedName],
                Index = index,
                Count = tags.Count
            };
        });
    }

    public static TagCollectionResult Get(
        PresentationContext context,
        object owner,
        Func<PowerPoint.Tags> acquireTags,
        string tagName)
    {
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(owner);
        ArgumentNullException.ThrowIfNull(acquireTags);

        if (!TryNormalizeName(tagName, out string normalizedName, out TagCollectionResult? error))
        {
            return error;
        }

        return WithTags(context, owner, acquireTags, tags =>
        {
            int index = FindIndex(tags, normalizedName);
            if (index == 0)
            {
                return MissingTag(normalizedName, tags.Count);
            }

            string storedName = tags.Name(index);
            return new TagCollectionResult
            {
                Success = true,
                Name = normalizedName,
                Value = tags[storedName],
                Index = index,
                Count = tags.Count
            };
        });
    }

    public static TagCollectionResult List(
        PresentationContext context,
        object owner,
        Func<PowerPoint.Tags> acquireTags)
    {
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(owner);
        ArgumentNullException.ThrowIfNull(acquireTags);

        return WithTags(context, owner, acquireTags, tags =>
        {
            var results = new List<TagInfo>(tags.Count);
            for (int index = 1; index <= tags.Count; index++)
            {
                results.Add(new TagInfo
                {
                    TagIndex = index,
                    Name = tags.Name(index),
                    Value = tags.Value(index)
                });
            }

            return new TagCollectionResult
            {
                Success = true,
                Count = tags.Count,
                Tags = results
            };
        });
    }

    public static TagCollectionResult Delete(
        PresentationContext context,
        object owner,
        Func<PowerPoint.Tags> acquireTags,
        string tagName)
    {
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(owner);
        ArgumentNullException.ThrowIfNull(acquireTags);

        if (!TryNormalizeName(tagName, out string normalizedName, out TagCollectionResult? error))
        {
            return error;
        }

        return WithTags(context, owner, acquireTags, tags =>
        {
            int index = FindIndex(tags, normalizedName);
            if (index == 0)
            {
                return MissingTag(normalizedName, tags.Count);
            }

            tags.Delete(tags.Name(index));
            return new TagCollectionResult
            {
                Success = true,
                Name = normalizedName,
                Count = tags.Count
            };
        });
    }

    private static T WithTags<T>(
        PresentationContext context,
        object owner,
        Func<PowerPoint.Tags> acquireTags,
        Func<PowerPoint.Tags, T> operation)
    {
        PowerPoint.Tags tags = context.GetOrAddOwnedComResource(owner, acquireTags);
        return operation(tags);
    }

    private static int FindIndex(PowerPoint.Tags tags, string normalizedName)
    {
        for (int index = 1; index <= tags.Count; index++)
        {
            if (string.Equals(tags.Name(index), normalizedName, StringComparison.OrdinalIgnoreCase))
            {
                return index;
            }
        }

        return 0;
    }

    private static bool TryNormalizeName(
        string tagName,
        out string normalizedName,
        out TagCollectionResult error)
    {
        if (!string.IsNullOrWhiteSpace(tagName))
        {
            normalizedName = tagName.ToUpperInvariant();
            error = null!;
            return true;
        }

        normalizedName = string.Empty;
        error = new TagCollectionResult
        {
            ErrorMessage = "A tag name is required."
        };
        return false;
    }

    private static TagCollectionResult MissingTag(string normalizedName, int count) => new()
    {
        ErrorMessage = $"No string tag named '{normalizedName}' was found.",
        Name = normalizedName,
        Count = count
    };
}
