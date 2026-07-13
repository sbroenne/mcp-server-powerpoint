using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;

namespace Sbroenne.PowerPointMcp.Core.Shape;

/// <inheritdoc cref="IShapeCommands"/>
public sealed class ShapeCommands : IShapeCommands
{
    // Shapes.AddShape's Type parameter and Shapes.AddTextbox's Orientation parameter are both
    // typed as Microsoft.Office.Core enums (MsoAutoShapeType / MsoTextOrientation) — i.e. office.dll
    // types. Per the project's established pattern (see PresentationBatch.cs), we avoid any
    // office.dll reference by calling these late-bound via dynamic, passing the raw enum int
    // values instead of the typed enum.
    private const int MsoShapeRectangle = 1; // MsoAutoShapeType.msoShapeRectangle
    private const int MsoTextOrientationHorizontal = 1; // MsoTextOrientation.msoTextOrientationHorizontal
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

    // PowerPoint.PpActionSetting.ppMouseClick / Office.MsoPresentationTarget-related action
    // constants — office.dll types, so used as raw ints per the project's dynamic-COM
    // convention (see MsoShapeRectangle above). Verified live via ActionSettings(ppMouseClick).
    private const int MsoMouseClick = 1; // PpActionSetting.ppMouseClick
    private const int PpActionNone = 0; // PpActionType.ppActionNone
    private const int PpActionHyperlink = 7; // PpActionType.ppActionHyperlink

    // MsoShadowStyle constant for SetShadow — verified live via ShapeEffectsDiagTests (a
    // temporary diagnostic spike, since removed): shape.Shadow.Type = 20 ("offset" style shadow)
    // is the closest match to GongRzhe/Office-PowerPoint-MCP-Server's parameterized drop-shadow
    // feature (color/transparency/blur/offset are all settable on this shadow style).
    private const int MsoShadow20 = 20;

    // MsoReflectionType constants for SetReflection/GetReflection — verified live via
    // ShapeEffectsDiagTests: 0 = none, 9 = "full reflection, touching" (the common default look).
    private const int MsoReflectionTypeNone = 0;
    private const int MsoReflectionType9 = 9;

    // MsoAutoShapeType member name -> value, for AddAutoShape. A curated subset of the full
    // enum covering the shapes authors most commonly need beyond a plain rectangle (arrows,
    // ovals, basic flowchart/callout shapes) — verified against the published MsoAutoShapeType
    // enumeration. Extend this table (never guess a value) if more shapes are needed later.
    private static readonly Dictionary<string, int> AutoShapeTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["msoShapeRectangle"] = 1,
        ["msoShapeParallelogram"] = 2,
        ["msoShapeTrapezoid"] = 3,
        ["msoShapeDiamond"] = 4,
        ["msoShapeRoundedRectangle"] = 5,
        ["msoShapeOctagon"] = 6,
        ["msoShapeIsoscelesTriangle"] = 7,
        ["msoShapeRightTriangle"] = 8,
        ["msoShapeOval"] = 9,
        ["msoShapeHexagon"] = 10,
        ["msoShapeCross"] = 11,
        ["msoShapeRegularPentagon"] = 12,
        ["msoShapeCan"] = 13,
        ["msoShapeCube"] = 14,
        ["msoShapeBevel"] = 15,
        ["msoShapeFoldedCorner"] = 16,
        ["msoShapeSmileyFace"] = 17,
        ["msoShapeDonut"] = 18,
        ["msoShapeNoSymbol"] = 19,
        ["msoShapeBlockArc"] = 20,
        ["msoShapeHeart"] = 21,
        ["msoShapeLightningBolt"] = 22,
        ["msoShapeSun"] = 23,
        ["msoShapeMoon"] = 24,
        ["msoShapeArc"] = 25,
        ["msoShapePlaque"] = 26,
        ["msoShapeLeftBracket"] = 29,
        ["msoShapeRightBracket"] = 30,
        ["msoShapeLeftBrace"] = 31,
        ["msoShapeRightBrace"] = 32,
        ["msoShapeRightArrow"] = 33,
        ["msoShapeLeftArrow"] = 34,
        ["msoShapeUpArrow"] = 35,
        ["msoShapeDownArrow"] = 36,
        ["msoShapeLeftRightArrow"] = 37,
        ["msoShapeUpDownArrow"] = 38,
    };

    // MsoConnectorType member name -> value, for AddConnector.
    private static readonly Dictionary<string, int> ConnectorTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["msoConnectorStraight"] = 1,
        ["msoConnectorElbow"] = 2,
        ["msoConnectorCurve"] = 3,
    };

    // MsoLineDashStyle member name -> value, for SetLine (learn.microsoft.com/office/vba/api/office.msolinedashstyle).
    private static readonly Dictionary<string, int> LineDashStyles = new(StringComparer.OrdinalIgnoreCase)
    {
        ["msoLineSolid"] = 1,
        ["msoLineSquareDot"] = 2,
        ["msoLineRoundDot"] = 3,
        ["msoLineDash"] = 4,
        ["msoLineDashDot"] = 5,
        ["msoLineDashDotDot"] = 6,
        ["msoLineLongDash"] = 7,
        ["msoLineLongDashDot"] = 8,
    };

    private static readonly Dictionary<int, string> LineDashStylesByValue =
        LineDashStyles.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    // MsoBevelType member name -> value, for SetBevel/GetBevel (learn.microsoft.com/office/vba/api/office.msobeveltype).
    private static readonly Dictionary<string, int> BevelTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        ["msoBevelNone"] = 6,
        ["msoBevelRelaxedInset"] = 2,
        ["msoBevelCircle"] = 3,
        ["msoBevelSlope"] = 8,
        ["msoBevelCross"] = 5,
        ["msoBevelAngle"] = 4,
        ["msoBevelSoftRound"] = 1,
        ["msoBevelConvex"] = 7,
        ["msoBevelCoolSlant"] = 12,
        ["msoBevelDivot"] = 9,
        ["msoBevelRiblet"] = 10,
        ["msoBevelHardEdge"] = 11,
        ["msoBevelArtDeco"] = 13,
    };

    private static readonly Dictionary<int, string> BevelTypesByValue =
        BevelTypes.ToDictionary(kvp => kvp.Value, kvp => kvp.Key);

    // MsoFlipCmd member -> value (learn.microsoft.com/office/vba/api/office.msoflipcmd).
    private static readonly Dictionary<string, int> FlipDirections = new(StringComparer.OrdinalIgnoreCase)
    {
        ["horizontal"] = 0, // msoFlipHorizontal
        ["vertical"] = 1,   // msoFlipVertical
    };

    // MsoZOrderCmd member -> value (learn.microsoft.com/office/vba/api/office.msozordercmd).
    // Word-only members (msoBringInFrontOfText, msoSendBehindText) are intentionally omitted.
    private static readonly Dictionary<string, int> ZOrderCommands = new(StringComparer.OrdinalIgnoreCase)
    {
        ["bring-to-front"] = 0, // msoBringToFront
        ["send-to-back"] = 1,   // msoSendToBack
        ["bring-forward"] = 2,  // msoBringForward
        ["send-backward"] = 3,  // msoSendBackward
    };

    /// <inheritdoc/>
    public ShapeOperationResult AddRectangle(IPresentationBatch batch, int slideIndex, float left, float top, float width, float height)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                slide.Shapes.AddShape(MsoShapeRectangle, left, top, width, height);
                // NOTE: discovered via integration test — accessing the newly-added shape's
                // .Index property dynamically threw a RuntimeBinderException ("'System.__ComObject'
                // does not contain a definition for 'Index'"), a NoPIA/late-binding quirk on the
                // COM object returned from AddShape. Sidestepped entirely: since shapes are always
                // appended, the new shape's 1-based index is simply the new Shapes.Count.
                int newIndex = (int)slide.Shapes.Count;

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = newIndex,
                    ShapeCount = (int)slide.Shapes.Count
                };
            }
            finally
            {
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult AddTextBox(IPresentationBatch batch, int slideIndex, float left, float top, float width, float height, string text)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(text);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                shape = slide.Shapes.AddTextbox(MsoTextOrientationHorizontal, left, top, width, height);
                shape.TextFrame.TextRange.Text = text;
                // Same NoPIA late-binding quirk as AddRectangle — avoid shape.Index, use Count.
                int newIndex = (int)slide.Shapes.Count;

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = newIndex,
                    ShapeCount = (int)slide.Shapes.Count
                };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult AddAutoShape(IPresentationBatch batch, int slideIndex, string shapeType, float left, float top, float width, float height)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(shapeType);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            if (!AutoShapeTypes.TryGetValue(shapeType, out var typeValue))
            {
                return new ShapeOperationResult
                {
                    Success = false,
                    ErrorMessage = $"'{shapeType}' is not a recognized MsoAutoShapeType name (e.g. 'msoShapeOval', 'msoShapeRightArrow', 'msoShapeDiamond')."
                };
            }

            dynamic? slide = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                slide.Shapes.AddShape(typeValue, left, top, width, height);
                // Same NoPIA late-binding quirk as AddRectangle — avoid shape.Index, use Count.
                int newIndex = (int)slide.Shapes.Count;

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = newIndex,
                    ShapeCount = (int)slide.Shapes.Count,
                    ShapeTypeName = shapeType
                };
            }
            finally
            {
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult AddLine(IPresentationBatch batch, int slideIndex, float beginX, float beginY, float endX, float endY)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                slide.Shapes.AddLine(beginX, beginY, endX, endY);
                // Same NoPIA late-binding quirk as AddRectangle — avoid shape.Index, use Count.
                int newIndex = (int)slide.Shapes.Count;

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = newIndex,
                    ShapeCount = (int)slide.Shapes.Count,
                    BeginX = beginX,
                    BeginY = beginY,
                    EndX = endX,
                    EndY = endY
                };
            }
            finally
            {
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult AddConnector(IPresentationBatch batch, int slideIndex, string connectorType, float beginX, float beginY, float endX, float endY)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(connectorType);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            if (!ConnectorTypes.TryGetValue(connectorType, out var typeValue))
            {
                return new ShapeOperationResult
                {
                    Success = false,
                    ErrorMessage = $"'{connectorType}' is not a recognized MsoConnectorType name (must be 'msoConnectorStraight', 'msoConnectorElbow', or 'msoConnectorCurve')."
                };
            }

            dynamic? slide = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                slide.Shapes.AddConnector(typeValue, beginX, beginY, endX, endY);
                // Same NoPIA late-binding quirk as AddRectangle — avoid shape.Index, use Count.
                int newIndex = (int)slide.Shapes.Count;

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = newIndex,
                    ShapeCount = (int)slide.Shapes.Count,
                    ConnectorTypeName = connectorType,
                    BeginX = beginX,
                    BeginY = beginY,
                    EndX = endX,
                    EndY = endY
                };
            }
            finally
            {
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult GetCount(IPresentationBatch batch, int slideIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeCount = (int)slide.Shapes.Count
                };
            }
            finally
            {
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult Delete(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                shape.Delete();

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    ShapeCount = (int)slide.Shapes.Count
                };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetPosition(IPresentationBatch batch, int slideIndex, int shapeIndex, float left, float top)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                shape.Left = left;
                shape.Top = top;

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    Left = (float)shape.Left,
                    Top = (float)shape.Top
                };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetSize(IPresentationBatch batch, int slideIndex, int shapeIndex, float width, float height)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                shape.Width = width;
                shape.Height = height;

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    Width = (float)shape.Width,
                    Height = (float)shape.Height
                };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetFill(IPresentationBatch batch, int slideIndex, int shapeIndex, byte red, byte green, byte blue)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                int rgb = red + (green << 8) + (blue << 16);
                shape.Fill.Solid();
                shape.Fill.ForeColor.RGB = rgb;

                return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, ColorRgb = rgb };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult GetFill(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                int rgb = (int)shape.Fill.ForeColor.RGB;

                return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, ColorRgb = rgb };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetLine(
        IPresentationBatch batch,
        int slideIndex,
        int shapeIndex,
        byte? red = null,
        byte? green = null,
        byte? blue = null,
        float? weight = null,
        string? dashStyle = null,
        bool? visible = null)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                int? dashStyleValue = null;
                if (dashStyle is not null)
                {
                    if (!LineDashStyles.TryGetValue(dashStyle, out var resolvedDashStyle))
                    {
                        return new ShapeOperationResult
                        {
                            Success = false,
                            ErrorMessage = $"'{dashStyle}' is not a recognized MsoLineDashStyle name (e.g. 'msoLineSolid', 'msoLineDash', 'msoLineDashDot')."
                        };
                    }
                    dashStyleValue = resolvedDashStyle;
                }

                shape = slide.Shapes[shapeIndex];

                if (red is not null || green is not null || blue is not null)
                {
                    int rgb = (red ?? 0) + ((green ?? 0) << 8) + ((blue ?? 0) << 16);
                    shape.Line.ForeColor.RGB = rgb;
                }

                if (weight is not null)
                {
                    shape.Line.Weight = weight.Value;
                }

                if (dashStyleValue is not null)
                {
                    shape.Line.DashStyle = dashStyleValue.Value;
                }

                if (visible is not null)
                {
                    shape.Line.Visible = visible.Value ? MsoTrue : MsoFalse;
                }

                return ReadLine(shape, shapeIndex);
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult GetLine(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                return ReadLine(shape, shapeIndex);
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetRotation(IPresentationBatch batch, int slideIndex, int shapeIndex, float degrees)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                shape.Rotation = degrees;

                return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, Rotation = (float)shape.Rotation };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult GetRotation(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];

                return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, Rotation = (float)shape.Rotation };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult Flip(IPresentationBatch batch, int slideIndex, int shapeIndex, string direction)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(direction);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                if (!FlipDirections.TryGetValue(direction, out var directionValue))
                {
                    return new ShapeOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"'{direction}' is not a recognized flip direction (must be 'horizontal' or 'vertical')."
                    };
                }

                shape = slide.Shapes[shapeIndex];
                shape.Flip(directionValue);

                return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, FlipDirection = direction };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetZOrder(IPresentationBatch batch, int slideIndex, int shapeIndex, string zOrderCommand)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(zOrderCommand);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                if (!ZOrderCommands.TryGetValue(zOrderCommand, out var commandValue))
                {
                    return new ShapeOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"'{zOrderCommand}' is not a recognized z-order command (must be 'bring-to-front', 'send-to-back', 'bring-forward', or 'send-backward')."
                    };
                }

                shape = slide.Shapes[shapeIndex];
                shape.ZOrder(commandValue);

                return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, ZOrderCommand = zOrderCommand };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetShadow(
        IPresentationBatch batch, int slideIndex, int shapeIndex, bool visible,
        byte red = 0, byte green = 0, byte blue = 0,
        float transparency = 0.5f, float blur = 4f, float offsetX = 3f, float offsetY = 3f)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            dynamic? shadow = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                shadow = shape.Shadow;
                shadow.Visible = visible ? MsoTrue : MsoFalse;

                if (!visible)
                {
                    return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, Visible = false };
                }

                int rgb = red + (green << 8) + (blue << 16);
                shadow.Type = MsoShadow20; // offset shadow — verified live via ShapeEffectsDiagTests
                shadow.ForeColor.RGB = rgb;
                shadow.Transparency = transparency;
                shadow.Blur = blur;
                shadow.OffsetX = offsetX;
                shadow.OffsetY = offsetY;

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    Visible = true,
                    ColorRgb = rgb,
                    Transparency = transparency,
                    Blur = blur,
                    OffsetX = offsetX,
                    OffsetY = offsetY,
                };
            }
            finally
            {
                if (shadow != null) ComUtilities.Release(ref shadow!);
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult GetShadow(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            dynamic? shadow = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                shadow = shape.Shadow;
                bool visible = (int)shadow.Visible == MsoTrue;

                if (!visible)
                {
                    return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, Visible = false };
                }

                int rgb = (int)shadow.ForeColor.RGB;
                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    Visible = true,
                    ColorRgb = rgb,
                    Transparency = (float)shadow.Transparency,
                    Blur = (float)shadow.Blur,
                    OffsetX = (float)shadow.OffsetX,
                    OffsetY = (float)shadow.OffsetY,
                };
            }
            finally
            {
                if (shadow != null) ComUtilities.Release(ref shadow!);
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetGlow(IPresentationBatch batch, int slideIndex, int shapeIndex, byte red, byte green, byte blue, float radius, float transparency = 0f)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            dynamic? glow = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                glow = shape.Glow;
                int rgb = red + (green << 8) + (blue << 16);
                glow.Radius = radius;
                glow.Color.RGB = rgb;
                glow.Transparency = transparency;

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    ColorRgb = rgb,
                    GlowRadius = radius,
                    Transparency = transparency,
                };
            }
            finally
            {
                if (glow != null) ComUtilities.Release(ref glow!);
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult GetGlow(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            dynamic? glow = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                glow = shape.Glow;

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    ColorRgb = (int)glow.Color.RGB,
                    GlowRadius = (float)glow.Radius,
                    Transparency = (float)glow.Transparency,
                };
            }
            finally
            {
                if (glow != null) ComUtilities.Release(ref glow!);
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetReflection(IPresentationBatch batch, int slideIndex, int shapeIndex, bool visible, float transparency = 0.5f, float size = 50f, float blur = 3f)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            dynamic? reflection = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                reflection = shape.Reflection;

                if (!visible)
                {
                    reflection.Type = MsoReflectionTypeNone;
                    return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, Visible = false };
                }

                reflection.Type = MsoReflectionType9; // full touching — verified live via ShapeEffectsDiagTests
                reflection.Transparency = transparency;
                reflection.Size = size;
                reflection.Blur = blur;

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    Visible = true,
                    Transparency = transparency,
                    ReflectionSize = size,
                    Blur = blur,
                };
            }
            finally
            {
                if (reflection != null) ComUtilities.Release(ref reflection!);
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult GetReflection(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            dynamic? reflection = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                reflection = shape.Reflection;
                bool visible = (int)reflection.Type != MsoReflectionTypeNone;

                if (!visible)
                {
                    return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, Visible = false };
                }

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    Visible = true,
                    Transparency = (float)reflection.Transparency,
                    ReflectionSize = (float)reflection.Size,
                    Blur = (float)reflection.Blur,
                };
            }
            finally
            {
                if (reflection != null) ComUtilities.Release(ref reflection!);
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetSoftEdge(IPresentationBatch batch, int slideIndex, int shapeIndex, float radius)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                shape.SoftEdge.Radius = radius;

                return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, SoftEdgeRadius = radius };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult GetSoftEdge(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                float radius = (float)shape.SoftEdge.Radius;

                return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, SoftEdgeRadius = radius };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetBevel(IPresentationBatch batch, int slideIndex, int shapeIndex, string bevelType, float depth = 6f, float inset = 6f)
    {
        ArgumentNullException.ThrowIfNull(batch);

        if (!BevelTypes.TryGetValue(bevelType, out var typeValue))
        {
            return new ShapeOperationResult
            {
                Success = false,
                ErrorMessage = $"'{bevelType}' is not a recognized MsoBevelType member name (e.g. 'msoBevelCircle', 'msoBevelSoftRound', 'msoBevelNone')."
            };
        }

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            dynamic? threeD = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                threeD = shape.ThreeD;
                threeD.BevelTopType = typeValue;
                threeD.BevelTopDepth = depth;
                threeD.BevelTopInset = inset;

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    BevelTypeName = bevelType,
                    BevelDepth = depth,
                    BevelInset = inset,
                };
            }
            finally
            {
                if (threeD != null) ComUtilities.Release(ref threeD!);
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult GetBevel(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            dynamic? threeD = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                threeD = shape.ThreeD;
                int typeValue = (int)threeD.BevelTopType;
                string typeName = BevelTypesByValue.TryGetValue(typeValue, out var name) ? name : $"unknown({typeValue})";

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    BevelTypeName = typeName,
                    BevelDepth = (float)threeD.BevelTopDepth,
                    BevelInset = (float)threeD.BevelTopInset,
                };
            }
            finally
            {
                if (threeD != null) ComUtilities.Release(ref threeD!);
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult Group(IPresentationBatch batch, int slideIndex, IReadOnlyList<int> shapeIndexes)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(shapeIndexes);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? range = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                int shapeCount = (int)slide.Shapes.Count;

                if (shapeIndexes.Count < 2)
                {
                    return new ShapeOperationResult
                    {
                        Success = false,
                        ErrorMessage = $"At least 2 shape indexes are required to group (got {shapeIndexes.Count})."
                    };
                }

                foreach (var index in shapeIndexes)
                {
                    var validation = ValidateShapeIndex(shapeCount, index);
                    if (validation is not null) return validation;
                }

                object[] indexArray = shapeIndexes.Select(i => (object)i).ToArray();
                range = slide.Shapes.Range(indexArray);
                range.Group();

                return new ShapeOperationResult { Success = true, ShapeCount = (int)slide.Shapes.Count };
            }
            finally
            {
                if (range != null) ComUtilities.Release(ref range!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult Ungroup(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            dynamic? ungrouped = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                ungrouped = shape.Ungroup();
                int ungroupedCount = (int)ungrouped.Count;

                return new ShapeOperationResult
                {
                    Success = true,
                    UngroupedShapeCount = ungroupedCount,
                    ShapeCount = (int)slide.Shapes.Count
                };
            }
            finally
            {
                if (ungrouped != null) ComUtilities.Release(ref ungrouped!);
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetName(IPresentationBatch batch, int slideIndex, int shapeIndex, string name)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(name);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                shape.Name = name;

                return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, Name = (string)shape.Name };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult GetName(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];

                return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, Name = (string)shape.Name };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetAltText(IPresentationBatch batch, int slideIndex, int shapeIndex, string altText)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(altText);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                shape.AlternativeText = altText;

                return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, AltText = (string)shape.AlternativeText };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult GetAltText(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];

                return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, AltText = (string)shape.AlternativeText };
            }
            finally
            {
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult SetHyperlink(IPresentationBatch batch, int slideIndex, int shapeIndex, string address, string? screenTip = null)
    {
        ArgumentNullException.ThrowIfNull(batch);
        ArgumentNullException.ThrowIfNull(address);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            dynamic? actionSetting = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                // ActionSettings(ppMouseClick).Hyperlink.Address — verified live: setting Address
                // automatically flips the action setting's Action to ppActionHyperlink; no separate
                // "enable hyperlink" step is needed.
                actionSetting = shape.ActionSettings[MsoMouseClick];
                actionSetting.Hyperlink.Address = address;
                if (screenTip is not null)
                {
                    actionSetting.Hyperlink.ScreenTip = screenTip;
                }

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    HasHyperlink = true,
                    HyperlinkAddress = (string)actionSetting.Hyperlink.Address,
                    HyperlinkScreenTip = (string)actionSetting.Hyperlink.ScreenTip
                };
            }
            finally
            {
                if (actionSetting != null) ComUtilities.Release(ref actionSetting!);
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult GetHyperlink(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            dynamic? actionSetting = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                actionSetting = shape.ActionSettings[MsoMouseClick];
                int action = (int)actionSetting.Action;
                bool hasHyperlink = action == PpActionHyperlink;

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    HasHyperlink = hasHyperlink,
                    HyperlinkAddress = hasHyperlink ? (string)actionSetting.Hyperlink.Address : null,
                    HyperlinkScreenTip = hasHyperlink ? (string)actionSetting.Hyperlink.ScreenTip : null
                };
            }
            finally
            {
                if (actionSetting != null) ComUtilities.Release(ref actionSetting!);
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    /// <inheritdoc/>
    public ShapeOperationResult RemoveHyperlink(IPresentationBatch batch, int slideIndex, int shapeIndex)
    {
        ArgumentNullException.ThrowIfNull(batch);

        return batch.Execute((ctx, ct) =>
        {
            var slideValidation = ValidateSlideIndex(ctx.Presentation.Slides.Count, slideIndex);
            if (slideValidation is not null) return slideValidation;

            dynamic? slide = null;
            dynamic? shape = null;
            dynamic? actionSetting = null;
            try
            {
                slide = ctx.Presentation.Slides[slideIndex];
                var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
                if (shapeValidation is not null) return shapeValidation;

                shape = slide.Shapes[shapeIndex];
                actionSetting = shape.ActionSettings[MsoMouseClick];
                actionSetting.Action = PpActionNone;

                return new ShapeOperationResult
                {
                    Success = true,
                    ShapeIndex = shapeIndex,
                    HasHyperlink = false
                };
            }
            finally
            {
                if (actionSetting != null) ComUtilities.Release(ref actionSetting!);
                if (shape != null) ComUtilities.Release(ref shape!);
                if (slide != null) ComUtilities.Release(ref slide!);
            }
        });
    }

    private static ShapeOperationResult ReadLine(dynamic shape, int shapeIndex)
    {
        int rgb = (int)shape.Line.ForeColor.RGB;
        float weight = (float)shape.Line.Weight;
        int dashStyleValue = (int)shape.Line.DashStyle;
        string dashStyleName = LineDashStylesByValue.TryGetValue(dashStyleValue, out var name)
            ? name
            : $"unknown ({dashStyleValue})";
        bool visible = (int)shape.Line.Visible == MsoTrue;

        return new ShapeOperationResult
        {
            Success = true,
            ShapeIndex = shapeIndex,
            ColorRgb = rgb,
            LineWeight = weight,
            DashStyleName = dashStyleName,
            Visible = visible
        };
    }

    private static ShapeOperationResult? ValidateSlideIndex(int slideCount, int slideIndex)
    {
        if (slideIndex < 1 || slideIndex > slideCount)
        {
            return new ShapeOperationResult
            {
                Success = false,
                ErrorMessage = $"Slide index {slideIndex} is out of range. The presentation has {slideCount} slide(s) (valid range: 1-{slideCount})."
            };
        }
        return null;
    }

    private static ShapeOperationResult? ValidateShapeIndex(int shapeCount, int shapeIndex)
    {
        if (shapeIndex < 1 || shapeIndex > shapeCount)
        {
            return new ShapeOperationResult
            {
                Success = false,
                ErrorMessage = $"Shape index {shapeIndex} is out of range. The slide has {shapeCount} shape(s) (valid range: 1-{shapeCount})."
            };
        }
        return null;
    }
}
