using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.SmartArt;

/// <inheritdoc cref="ISmartArtCommands"/>
public sealed class SmartArtCommands : ISmartArtCommands
{
    private const int MsoTrue = -1;

    /// <inheritdoc/>
    public SmartArtOperationResult AddSmartArt(
        IPresentationBatch batch,
        int slideIndex,
        string layoutName,
        float left,
        float top,
        float width,
        float height)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(layoutName);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            // ctx.App is the shared, session-owned Application COM object - never release it here.
            dynamic app = ctx.App;
            dynamic? layout = null;
            dynamic? dynShapes = null;
            dynamic? newShape = null;
            try
            {
                layout = RetryOnTransientAccessDenied(() => FindLayoutByName(app.SmartArtLayouts, layoutName));
                if (layout is null)
                {
                    return new SmartArtOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"'{layoutName}' is not a recognized SmartArt layout name. Layout names come from PowerPoint's SmartArt gallery (e.g. 'Basic Process', 'Organization Chart', 'Basic Cycle', 'Basic Pyramid'). See smart-art.md for the full list."
                    };
                }

                PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
                dynShapes = slide.Shapes;
                dynShapes.AddSmartArt(layout, left, top, width, height);
                // Same NoPIA late-binding quirk as ShapeCommands.AddRectangle — avoid shape.Index, use Count.
                int newIndex = slide.Shapes.Count;
                newShape = slide.Shapes[newIndex];
                int nodeCount = (int)newShape.SmartArt.AllNodes.Count;

                return new SmartArtOperationResult
                {
                    Success = true,
                    ShapeIndex = newIndex,
                    ShapeCount = slide.Shapes.Count,
                    LayoutName = layoutName,
                    NodeCount = nodeCount
                };
            }
            finally
            {
                if (newShape != null)
                {
                    ComUtilities.Release(ref newShape!);
                }
                if (dynShapes != null)
                {
                    ComUtilities.Release(ref dynShapes!);
                }
                if (layout != null)
                {
                    ComUtilities.Release(ref layout!);
                }
                // app is an alias of ctx.App (shared, session-owned) - must NOT be released.
            }
        });
    }

    /// <inheritdoc/>
    public SmartArtOperationResult AddNode(IPresentationBatch batch, int slideIndex, int shapeIndex, string text)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(text);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateSmartArtShape(ctx, slideIndex, shapeIndex, out var smartArt);
            if (validation is not null) return validation;

            dynamic? nodes = null;
            dynamic? newNode = null;
            try
            {
                nodes = smartArt!.Nodes;
                newNode = nodes.Add();
                newNode.TextFrame2.TextRange.Text = text;

                // SmartArtNodes.Add() always appends "to the bottom of the data model at the top
                // most level" (verified live + per Microsoft's SmartArtNodes.Add documentation) — a
                // brand-new top-level node has no children yet and sits after every other top-level
                // branch, so it is always the LAST entry in the flattened AllNodes collection. No
                // search needed (SmartArtNode exposes no reliable, late-bindable unique identifier —
                // see FindNodeIndexByLevelAndText's remarks for why).
                int nodeIndex = (int)smartArt.AllNodes.Count;

                return new SmartArtOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    NodeIndex = nodeIndex,
                    NodeCount = nodeIndex,
                    NodeText = text
                };
            }
            finally
            {
                if (newNode != null)
                {
                    ComUtilities.Release(ref newNode!);
                }
                if (nodes != null)
                {
                    ComUtilities.Release(ref nodes!);
                }
                if (smartArt != null)
                {
                    ComUtilities.Release(ref smartArt!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public SmartArtOperationResult AddChildNode(IPresentationBatch batch, int slideIndex, int shapeIndex, int parentNodeIndex, string text)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(text);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateSmartArtShape(ctx, slideIndex, shapeIndex, out var smartArt);
            if (validation is not null) return validation;

            dynamic? allNodes = null;
            dynamic? parentNode = null;
            dynamic? newNode = null;
            try
            {
                allNodes = smartArt!.AllNodes;
                var nodeValidation = ValidateNodeIndex((int)allNodes.Count, parentNodeIndex);
                if (nodeValidation is not null) return nodeValidation;

                parentNode = allNodes.Item(parentNodeIndex);
                int parentLevel = (int)parentNode.Level;
                // SmartArtNode.AddNode(Position, Type) is the documented way to add a node relative
                // to an existing one (see Office.SmartArtNode.AddNode). msoSmartArtNodeBelow (5) with
                // msoSmartArtNodeTypeDefault (1) makes the new node a child one level below the
                // target — verified live against a real Organization Chart diagram. (The
                // SmartArtNodes.Add(pTargetNode:=) overload some samples reference does not resolve
                // via C#'s dynamic IDispatch binder — DISP_E_UNKNOWNNAME — so it is not used here.)
                newNode = parentNode.AddNode(5 /* msoSmartArtNodeBelow */, 1 /* msoSmartArtNodeTypeDefault */);
                newNode.TextFrame2.TextRange.Text = text;

                int nodeIndex = FindNodeIndexByLevelAndText(allNodes, parentLevel + 1, text);

                return new SmartArtOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    NodeIndex = nodeIndex,
                    NodeCount = (int)allNodes.Count,
                    NodeText = text
                };
            }
            finally
            {
                if (newNode != null)
                {
                    ComUtilities.Release(ref newNode!);
                }
                if (parentNode != null)
                {
                    ComUtilities.Release(ref parentNode!);
                }
                if (allNodes != null)
                {
                    ComUtilities.Release(ref allNodes!);
                }
                if (smartArt != null)
                {
                    ComUtilities.Release(ref smartArt!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public SmartArtOperationResult SetNodeText(IPresentationBatch batch, int slideIndex, int shapeIndex, int nodeIndex, string text)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(text);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateSmartArtShape(ctx, slideIndex, shapeIndex, out var smartArt);
            if (validation is not null) return validation;

            dynamic? allNodes = null;
            dynamic? node = null;
            try
            {
                allNodes = smartArt!.AllNodes;
                var nodeValidation = ValidateNodeIndex((int)allNodes.Count, nodeIndex);
                if (nodeValidation is not null) return nodeValidation;

                node = allNodes.Item(nodeIndex);
                node.TextFrame2.TextRange.Text = text;

                return new SmartArtOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    NodeIndex = nodeIndex,
                    NodeCount = (int)allNodes.Count,
                    NodeText = text
                };
            }
            finally
            {
                if (node != null)
                {
                    ComUtilities.Release(ref node!);
                }
                if (allNodes != null)
                {
                    ComUtilities.Release(ref allNodes!);
                }
                if (smartArt != null)
                {
                    ComUtilities.Release(ref smartArt!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public SmartArtOperationResult GetNodeText(IPresentationBatch batch, int slideIndex, int shapeIndex, int nodeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateSmartArtShape(ctx, slideIndex, shapeIndex, out var smartArt);
            if (validation is not null) return validation;

            dynamic? allNodes = null;
            dynamic? node = null;
            try
            {
                allNodes = smartArt!.AllNodes;
                var nodeValidation = ValidateNodeIndex((int)allNodes.Count, nodeIndex);
                if (nodeValidation is not null) return nodeValidation;

                node = allNodes.Item(nodeIndex);
                string text = (string)node.TextFrame2.TextRange.Text;

                return new SmartArtOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    NodeIndex = nodeIndex,
                    NodeCount = (int)allNodes.Count,
                    NodeText = text
                };
            }
            finally
            {
                if (node != null)
                {
                    ComUtilities.Release(ref node!);
                }
                if (allNodes != null)
                {
                    ComUtilities.Release(ref allNodes!);
                }
                if (smartArt != null)
                {
                    ComUtilities.Release(ref smartArt!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public SmartArtOperationResult DeleteNode(IPresentationBatch batch, int slideIndex, int shapeIndex, int nodeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateSmartArtShape(ctx, slideIndex, shapeIndex, out var smartArt);
            if (validation is not null) return validation;

            dynamic? allNodes = null;
            dynamic? node = null;
            try
            {
                allNodes = smartArt!.AllNodes;
                var nodeValidation = ValidateNodeIndex((int)allNodes.Count, nodeIndex);
                if (nodeValidation is not null) return nodeValidation;

                node = allNodes.Item(nodeIndex);
                node.Delete();

                return new SmartArtOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    NodeIndex = nodeIndex,
                    NodeCount = (int)allNodes.Count
                };
            }
            finally
            {
                if (node != null)
                {
                    ComUtilities.Release(ref node!);
                }
                if (allNodes != null)
                {
                    ComUtilities.Release(ref allNodes!);
                }
                if (smartArt != null)
                {
                    ComUtilities.Release(ref smartArt!);
                }
            }
        });
    }

    /// <inheritdoc/>
    public SmartArtOperationResult GetNodeCount(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var validation = ValidateSmartArtShape(ctx, slideIndex, shapeIndex, out var smartArt);
            if (validation is not null) return validation;

            try
            {
                return new SmartArtOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    NodeCount = (int)smartArt!.AllNodes.Count
                };
            }
            finally
            {
                if (smartArt != null)
                {
                    ComUtilities.Release(ref smartArt!);
                }
            }
        });
    }

    /// <summary>
    /// Retries <paramref name="action"/> a few times (short linear backoff) when it throws
    /// <see cref="UnauthorizedAccessException"/> with HResult <c>0x80070005</c> (E_ACCESSDENIED).
    /// </summary>
    /// <remarks>
    /// Verified live: <c>Application.SmartArtLayouts</c> can occasionally throw E_ACCESSDENIED
    /// for a short window right after Office finishes loading the SmartArt gallery, recovering on
    /// its own within a few hundred milliseconds — not a caller-correctable validation failure, so
    /// it is retried here rather than turned into a <c>Success = false</c> result (Rule 1b: only
    /// expected, caller-correctable failures are validated up front; this is an environmental COM
    /// transient). Note: this does NOT paper over the distinct, non-transient E_ACCESSDENIED that
    /// occurs when querying <c>SmartArtLayouts</c> immediately after closing/reopening a
    /// presentation via <c>PresentationBatch.ReopenPresentation</c> — that failure mode persists
    /// well beyond any retry budget and must be avoided by callers (isolate via a fresh slide in
    /// the same presentation instead of reopening it). If retries are exhausted here, the
    /// exception propagates naturally.
    /// </remarks>
    private static T RetryOnTransientAccessDenied<T>(Func<T> action)
    {
        const int accessDeniedHResult = unchecked((int)0x80070005);
        const int maxAttempts = 5;

        for (int attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                return action();
            }
            catch (UnauthorizedAccessException ex) when (ex.HResult == accessDeniedHResult && attempt < maxAttempts)
            {
                Thread.Sleep(TimeSpan.FromMilliseconds(200));
            }
        }

        return action();
    }

    /// <summary>
    /// Finds the SmartArt layout in <paramref name="layouts"/> (an <c>Office.SmartArtLayouts</c>
    /// collection, accessed dynamically to avoid an office.dll reference) whose <c>.Name</c>
    /// matches <paramref name="layoutName"/> (case-insensitive), or null if none matches.
    /// </summary>
    private static dynamic? FindLayoutByName(dynamic layouts, string layoutName)
    {
        int count = (int)layouts.Count;
        for (int i = 1; i <= count; i++)
        {
            dynamic? candidate = null;
            bool matched = false;
            try
            {
                candidate = layouts.Item(i);
                string name = (string)candidate.Name;
                if (string.Equals(name, layoutName, StringComparison.OrdinalIgnoreCase))
                {
                    matched = true;
                    return candidate;
                }
            }
            finally
            {
                if (candidate != null && !matched)
                {
                    ComUtilities.Release(ref candidate!);
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Finds the 1-based position, within the flattened <paramref name="allNodes"/> collection,
    /// of the most recently added node at hierarchy <paramref name="level"/> whose text equals
    /// <paramref name="text"/> (scanning from the end backwards, so the newest match wins on a
    /// text collision).
    /// </summary>
    /// <remarks>
    /// <c>SmartArtNode</c> exposes no unique identifier that resolves via C#'s dynamic IDispatch
    /// binder (its documented member set is Application/Creator/Hidden/Level/Nodes/OrgChartLayout
    /// /Parent/ParentNode/Shapes/TextFrame2/Type — no <c>Id</c>), and comparing the COM identity
    /// of two <c>dynamic</c> references via <c>Marshal.GetIUnknownForObject</c> was verified live
    /// to NOT reliably match for this object type (each late-bound property access appears to
    /// hand back a distinct proxy). Since every caller of this helper sets the node's text
    /// immediately after creation, matching on (level, text) is a practical, good-enough
    /// resolution strategy — same trade-off class as <c>ShapeCommands</c>'s Count-based index
    /// workaround for the NoPIA late-binding quirk.
    /// </remarks>
    private static int FindNodeIndexByLevelAndText(dynamic allNodes, int level, string text)
    {
        int count = (int)allNodes.Count;
        for (int i = count; i >= 1; i--)
        {
            dynamic? candidate = null;
            try
            {
                candidate = allNodes.Item(i);
                if ((int)candidate.Level == level && string.Equals((string)candidate.TextFrame2.TextRange.Text, text, StringComparison.Ordinal))
                {
                    return i;
                }
            }
            finally
            {
                if (candidate != null)
                {
                    ComUtilities.Release(ref candidate!);
                }
            }
        }

        // Should be unreachable — the node was just added and its text just set.
        return count;
    }

    /// <summary>
    /// Validates <paramref name="slideIndex"/>/<paramref name="shapeIndex"/> and that the shape
    /// at that position is a SmartArt diagram (<c>Shape.HasSmartArt</c>). On success, returns
    /// null and sets <paramref name="smartArt"/> to the shape's <c>SmartArt</c> object.
    /// </summary>
    private static SmartArtOperationResult? ValidateSmartArtShape(
        PresentationContext ctx,
        int slideIndex,
        int shapeIndex,
        out dynamic? smartArt)
    {
        smartArt = null;

        var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
        if (slideValidation is not null) return slideValidation;

        PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
        int shapeCount = slide.Shapes.Count;
        var shapeValidation = ValidateShapeIndex(shapeCount, shapeIndex);
        if (shapeValidation is not null) return shapeValidation;

        PowerPoint.Shape shape = slide.Shapes[shapeIndex];
        // Reason: HasSmartArt/SmartArt are Office.Core-backed members not on the strongly-typed
        // PIA Shape interface, so they are read via dynamic late binding.
        bool hasSmartArt = (int)((dynamic)shape).HasSmartArt == MsoTrue;
        if (!hasSmartArt)
        {
            return new SmartArtOperationResult
            {
                Success = false,
                ErrorMessage = $"Shape {shapeIndex} on slide {slideIndex} is not a SmartArt diagram."
            };
        }

        // Reason: SmartArt is an Office.Core-backed member not on the strongly-typed PIA Shape
        // interface, so it is read via dynamic late binding.
        smartArt = ((dynamic)shape).SmartArt;
        return null;
    }

    private static SmartArtOperationResult? ValidateSlideIndex(int slideCount, int slideIndex)
    {
        if (slideIndex < 1 || slideIndex > slideCount)
        {
            return new SmartArtOperationResult
            {
                Success = false,
                ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
            };
        }
        return null;
    }

    private static SmartArtOperationResult? ValidateShapeIndex(int shapeCount, int shapeIndex)
    {
        if (shapeIndex < 1 || shapeIndex > shapeCount)
        {
            return new SmartArtOperationResult
            {
                Success = false,
                ErrorMessage = $"Shape index {shapeIndex} is out of range. The slide has {shapeCount} shape(s) (valid range: 1-{shapeCount})."
            };
        }
        return null;
    }

    private static SmartArtOperationResult? ValidateNodeIndex(int nodeCount, int nodeIndex)
    {
        if (nodeIndex < 1 || nodeIndex > nodeCount)
        {
            return new SmartArtOperationResult
            {
                Success = false,
                ErrorMessage = $"Node index {nodeIndex} is out of range. The diagram has {nodeCount} node(s) (valid range: 1-{nodeCount})."
            };
        }
        return null;
    }
}
