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
