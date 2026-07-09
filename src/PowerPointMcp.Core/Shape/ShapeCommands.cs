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

            dynamic slide = ctx.Presentation.Slides[slideIndex];
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

            dynamic slide = ctx.Presentation.Slides[slideIndex];
            dynamic shape = slide.Shapes.AddTextbox(MsoTextOrientationHorizontal, left, top, width, height);
            shape.TextFrame.TextRange.Text = text;
            // Same NoPIA late-binding quirk as AddRectangle — avoid shape.Index, use Count.
            int newIndex = (int)slide.Shapes.Count;

            return new ShapeOperationResult
            {
                Success = true,
                ShapeIndex = newIndex,
                ShapeCount = (int)slide.Shapes.Count
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

            dynamic slide = ctx.Presentation.Slides[slideIndex];
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

            dynamic slide = ctx.Presentation.Slides[slideIndex];
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

            dynamic slide = ctx.Presentation.Slides[slideIndex];
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

            dynamic shape = slide.Shapes[shapeIndex];
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

            dynamic shape = slide.Shapes[shapeIndex];
            int rgb = (int)shape.Fill.ForeColor.RGB;

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

            dynamic shape = slide.Shapes[shapeIndex];

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

            dynamic shape = slide.Shapes[shapeIndex];
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
    public ShapeOperationResult SetShadow(IPresentationBatch batch, int slideIndex, int shapeIndex, bool visible)
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
            shape.Shadow.Visible = visible ? MsoTrue : MsoFalse;

            return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, Visible = visible };
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
            bool visible = (int)shape.Shadow.Visible == MsoTrue;

            return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, Visible = visible };
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

            dynamic slide = ctx.Presentation.Slides[slideIndex];
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
            dynamic range = slide.Shapes.Range(indexArray);
            range.Group();

            return new ShapeOperationResult { Success = true, ShapeCount = (int)slide.Shapes.Count };
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

            dynamic slide = ctx.Presentation.Slides[slideIndex];
            var shapeValidation = ValidateShapeIndex((int)slide.Shapes.Count, shapeIndex);
            if (shapeValidation is not null) return shapeValidation;

            dynamic shape = slide.Shapes[shapeIndex];
            dynamic ungrouped = shape.Ungroup();
            int ungroupedCount = (int)ungrouped.Count;

            return new ShapeOperationResult
            {
                Success = true,
                UngroupedShapeCount = ungroupedCount,
                ShapeCount = (int)slide.Shapes.Count
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

            dynamic shape = slide.Shapes[shapeIndex];
            shape.AlternativeText = altText;

            return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, AltText = (string)shape.AlternativeText };
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

            dynamic shape = slide.Shapes[shapeIndex];

            return new ShapeOperationResult { Success = true, ShapeIndex = shapeIndex, AltText = (string)shape.AlternativeText };
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
