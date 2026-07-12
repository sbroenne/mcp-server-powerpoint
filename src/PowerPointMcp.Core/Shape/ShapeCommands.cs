using Sbroenne.PowerPointMcp.ComInterop.Session;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

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
    private const PowerPoint.PpMouseActivation MouseClickActivation = PowerPoint.PpMouseActivation.ppMouseClick;
    private const PowerPoint.PpActionType PpActionNone = PowerPoint.PpActionType.ppActionNone;
    private const PowerPoint.PpActionType PpActionHyperlink = PowerPoint.PpActionType.ppActionHyperlink;

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

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            dynamic dynShapes = slide.Shapes;
            dynShapes.AddShape(MsoShapeRectangle, left, top, width, height);
            // NOTE: discovered via integration test — accessing the newly-added shape's
            // .Index property dynamically threw a RuntimeBinderException ("'System.__ComObject'
            // does not contain a definition for 'Index'"), a NoPIA/late-binding quirk on the
            // COM object returned from AddShape. Sidestepped entirely: since shapes are always
            // appended, the new shape's 1-based index is simply the new Shapes.Count.
            int newIndex = slide.Shapes.Count;

            return new ShapeOperationResult
            {
                Success = true,
                ShapeIndex = newIndex,
                ShapeCount = slide.Shapes.Count
            };
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

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            dynamic dynShapes = slide.Shapes;
            PowerPoint.Shape shape = dynShapes.AddTextbox(MsoTextOrientationHorizontal, left, top, width, height);
            shape.TextFrame.TextRange.Text = text;
            // Same NoPIA late-binding quirk as AddRectangle — avoid shape.Index, use Count.
            int newIndex = slide.Shapes.Count;

            return new ShapeOperationResult
            {
                Success = true,
                ShapeIndex = newIndex,
                ShapeCount = slide.Shapes.Count
            };
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

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            dynamic dynShapes = slide.Shapes;
            dynShapes.AddShape(typeValue, left, top, width, height);
            // Same NoPIA late-binding quirk as AddRectangle — avoid shape.Index, use Count.
            int newIndex = slide.Shapes.Count;

            return new ShapeOperationResult
            {
                Success = true,
                ShapeIndex = newIndex,
                ShapeCount = slide.Shapes.Count,
                ShapeTypeName = shapeType
            };
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

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            slide.Shapes.AddLine(beginX, beginY, endX, endY);
            // Same NoPIA late-binding quirk as AddRectangle — avoid shape.Index, use Count.
            int newIndex = slide.Shapes.Count;

            return new ShapeOperationResult
            {
                Success = true,
                ShapeIndex = newIndex,
                ShapeCount = slide.Shapes.Count,
                BeginX = beginX,
                BeginY = beginY,
                EndX = endX,
                EndY = endY
            };
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

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            dynamic dynShapes = slide.Shapes;
            dynShapes.AddConnector(typeValue, beginX, beginY, endX, endY);
            // Same NoPIA late-binding quirk as AddRectangle — avoid shape.Index, use Count.
            int newIndex = slide.Shapes.Count;

            return new ShapeOperationResult
            {
                Success = true,
                ShapeIndex = newIndex,
                ShapeCount = slide.Shapes.Count,
                ConnectorTypeName = connectorType,
                BeginX = beginX,
                BeginY = beginY,
                EndX = endX,
                EndY = endY
            };
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

            return new ShapeOperationResult
            {
                Success = true,
                ShapeCount = ctx.Presentation.Slides[slideIndex].Shapes.Count
            };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            slide.Shapes[shapeIndex].Delete();

            return new ShapeOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                ShapeCount = slide.Shapes.Count
            };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            var shape = slide.Shapes[shapeIndex];
            shape.Left = left;
            shape.Top = top;

            return new ShapeOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                Left = shape.Left,
                Top = shape.Top
            };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            var shape = slide.Shapes[shapeIndex];
            shape.Width = width;
            shape.Height = height;

            return new ShapeOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                Width = shape.Width,
                Height = shape.Height
            };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            int rgb = red + (green << 8) + (blue << 16);
            shape.Fill.Solid();
            shape.Fill.ForeColor.RGB = rgb;

            return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, ColorRgb = rgb };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            int rgb = shape.Fill.ForeColor.RGB;

            return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, ColorRgb = rgb };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
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

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            dynamic line = shape.Line;

            if (red is not null || green is not null || blue is not null)
            {
                int rgb = (red ?? 0) + ((green ?? 0) << 8) + ((blue ?? 0) << 16);
                line.ForeColor.RGB = rgb;
            }

            if (weight is not null)
            {
                line.Weight = weight.Value;
            }

            if (dashStyleValue is not null)
            {
                line.DashStyle = dashStyleValue.Value;
            }

            if (visible is not null)
            {
                line.Visible = visible.Value ? MsoTrue : MsoFalse;
            }

            return ReadLine(shape, shapeIndex);
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            return ReadLine(shape, shapeIndex);
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            var shape = slide.Shapes[shapeIndex];
            shape.Rotation = degrees;

            return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, Rotation = shape.Rotation };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            var shape = slide.Shapes[shapeIndex];

            return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, Rotation = shape.Rotation };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            if (!FlipDirections.TryGetValue(direction, out var directionValue))
            {
                return new ShapeOperationResult
                {
                    Success = false,
                    ErrorMessage = $"'{direction}' is not a recognized flip direction (must be 'horizontal' or 'vertical')."
                };
            }

            dynamic shape = slide.Shapes[shapeIndex];
            shape.Flip(directionValue);

            return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, FlipDirection = direction };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            if (!ZOrderCommands.TryGetValue(zOrderCommand, out var commandValue))
            {
                return new ShapeOperationResult
                {
                    Success = false,
                    ErrorMessage = $"'{zOrderCommand}' is not a recognized z-order command (must be 'bring-to-front', 'send-to-back', 'bring-forward', or 'send-backward')."
                };
            }

            dynamic shape = slide.Shapes[shapeIndex];
            shape.ZOrder(commandValue);

            return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, ZOrderCommand = zOrderCommand };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            dynamic shape = slide.Shapes[shapeIndex];
            dynamic shadow = shape.Shadow;
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            dynamic shape = slide.Shapes[shapeIndex];
            dynamic shadow = shape.Shadow;
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            dynamic shape = slide.Shapes[shapeIndex];
            dynamic glow = shape.Glow;
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            dynamic shape = slide.Shapes[shapeIndex];
            dynamic glow = shape.Glow;

            return new ShapeOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                ColorRgb = (int)glow.Color.RGB,
                GlowRadius = (float)glow.Radius,
                Transparency = (float)glow.Transparency,
            };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            dynamic shape = slide.Shapes[shapeIndex];
            dynamic reflection = shape.Reflection;

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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            dynamic shape = slide.Shapes[shapeIndex];
            dynamic reflection = shape.Reflection;
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            dynamic shape = slide.Shapes[shapeIndex];
            shape.SoftEdge.Radius = radius;

            return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, SoftEdgeRadius = radius };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            dynamic shape = slide.Shapes[shapeIndex];
            float radius = (float)shape.SoftEdge.Radius;

            return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, SoftEdgeRadius = radius };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            dynamic shape = slide.Shapes[shapeIndex];
            dynamic threeD = shape.ThreeD;
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            dynamic shape = slide.Shapes[shapeIndex];
            dynamic threeD = shape.ThreeD;
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

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            int shapeCount = slide.Shapes.Count;

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
            PowerPoint.ShapeRange range = slide.Shapes.Range(indexArray);
            range.Group();

            return new ShapeOperationResult { Success = true, ShapeCount = slide.Shapes.Count };
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

            PowerPoint.Slide slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            PowerPoint.ShapeRange ungrouped = shape.Ungroup();
            int ungroupedCount = ungrouped.Count;

            return new ShapeOperationResult
            {
                Success = true,
                UngroupedShapeCount = ungroupedCount,
                ShapeCount = slide.Shapes.Count
            };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            var shape = slide.Shapes[shapeIndex];
            shape.Name = name;

            return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, Name = shape.Name };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            var shape = slide.Shapes[shapeIndex];

            return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, Name = shape.Name };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            shape.AlternativeText = altText;

            return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, AltText = shape.AlternativeText };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];

            return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, AltText = shape.AlternativeText };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            // ActionSettings(ppMouseClick).Hyperlink.Address — verified live: setting Address
            // automatically flips the action setting's Action to ppActionHyperlink; no separate
            // "enable hyperlink" step is needed.
            PowerPoint.ActionSetting actionSetting = shape.ActionSettings[MouseClickActivation];
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
                HyperlinkAddress = actionSetting.Hyperlink.Address,
                HyperlinkScreenTip = actionSetting.Hyperlink.ScreenTip
            };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            PowerPoint.ActionSetting actionSetting = shape.ActionSettings[MouseClickActivation];
            PowerPoint.PpActionType action = actionSetting.Action;
            bool hasHyperlink = action == PpActionHyperlink;

            return new ShapeOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                HasHyperlink = hasHyperlink,
                HyperlinkAddress = hasHyperlink ? actionSetting.Hyperlink.Address : null,
                HyperlinkScreenTip = hasHyperlink ? actionSetting.Hyperlink.ScreenTip : null
            };
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

            var slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex(slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            PowerPoint.Shape shape = slide.Shapes[shapeIndex];
            PowerPoint.ActionSetting actionSetting = shape.ActionSettings[MouseClickActivation];
            actionSetting.Action = PpActionNone;

            return new ShapeOperationResult
            {
                Success = true,
                ShapeIndex = shapeIndex,
                HasHyperlink = false
            };
        });
    }

    private static ShapeOperationResult ReadLine(PowerPoint.Shape shape, int shapeIndex)
    {
        dynamic line = shape.Line;
        int rgb = (int)line.ForeColor.RGB;
        float weight = (float)line.Weight;
        int dashStyleValue = (int)line.DashStyle;
        string dashStyleName = LineDashStylesByValue.TryGetValue(dashStyleValue, out var name)
            ? name
            : $"unknown ({dashStyleValue})";
        bool visible = (int)line.Visible == MsoTrue;

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
