using Aspose.Slides.Foss;
using Aspose.Slides.Foss.Export;
using FluentAssertions;

namespace Aspose.Slides.Foss.Tests;

/// <summary>
/// Tests for PVIObject, ISlideComponent, and IPresentationComponent hierarchy.
/// test_shapes (shape component hierarchy), and test_slides (slide layout access).
/// </summary>
public sealed class PVIObjectTests
{
    private sealed class BaseSlideStub : IBaseSlide
    {
        public IPresentation? Presentation { get; set; }
        public IShapeCollection? Shapes => null;
        public string Name { get; set; } = string.Empty;
        public int SlideId => 0;
    }

    private sealed class PresentationStub : IPresentation
    {
        public override DateTime CurrentDateTime { get; set; }
        public override ISlideCollection Slides => null!;
        public override INotesSize NotesSize => null!;
        public override IGlobalLayoutSlideCollection LayoutSlides => null!;
        public override IMasterSlideCollection Masters => null!;
        public override ICommentAuthorCollection CommentAuthors => null!;
        public override IDocumentProperties DocumentProperties => null!;
        public override IImageCollection Images => null!;
        public override SourceFormat SourceFormat => SourceFormat.Pptx;
        public override int FirstSlideNumber { get; set; }
        public override ISectionCollection Sections => null!;
        public override IPresentationComponent AsIPresentationComponent => this;
        public override void Save(string fname, SaveFormat format) => throw new InvalidOperationException("Not used in tests.");
        public override void Save(Stream stream, SaveFormat format) => throw new InvalidOperationException("Not used in tests.");
        public override void Save(string fname, SaveFormat format, ISaveOptions options) => throw new InvalidOperationException("Not used in tests.");
        public override void Save(Stream stream, SaveFormat format, ISaveOptions options) => throw new InvalidOperationException("Not used in tests.");
        public override void Save(string fname, int[] slides, SaveFormat format) => throw new InvalidOperationException("Not used in tests.");
        public override void Save(string fname, int[] slides, SaveFormat format, ISaveOptions options) => throw new InvalidOperationException("Not used in tests.");
        public override void Save(Stream stream, int[] slides, SaveFormat format) => throw new InvalidOperationException("Not used in tests.");
        public override void Save(Stream stream, int[] slides, SaveFormat format, ISaveOptions options) => throw new InvalidOperationException("Not used in tests.");
        public override void Save(ISaveOptions options) => throw new InvalidOperationException("Not used in tests.");
    }

    private sealed class TestablePVIObject : PVIObject
    {
        public void SetParentSlide(IBaseSlide? slide) => _parentSlide = slide;
    }

    [Fact]
    public void Slide_IsNullByDefault()
    {
        var obj = new TestablePVIObject();

        obj.Slide.Should().BeNull();
    }

    [Fact]
    public void Slide_ReturnsAssignedSlide()
    {
        // From test_notes_parent_slide: notes slide references its parent slide
        var obj = new TestablePVIObject();
        var slide = new BaseSlideStub();

        obj.SetParentSlide(slide);

        obj.Slide.Should().BeSameAs(slide);
    }

    [Fact]
    public void Presentation_IsNullWhenNoSlide()
    {
        var obj = new TestablePVIObject();

        obj.Presentation.Should().BeNull();
    }

    [Fact]
    public void Presentation_DelegatesToSlide()
    {
        var pres = new PresentationStub();
        var slide = new BaseSlideStub { Presentation = pres };
        var obj = new TestablePVIObject();

        obj.SetParentSlide(slide);

        obj.Presentation.Should().BeSameAs(pres);
    }

    [Fact]
    public void AsIPresentationComponent_ReturnsSelf()
    {
        var obj = new TestablePVIObject();

        obj.AsIPresentationComponent.Should().BeSameAs(obj);
    }

    [Fact]
    public void IsInstanceOf_ISlideComponent()
    {
        var obj = new TestablePVIObject();

        obj.Should().BeAssignableTo<ISlideComponent>();
    }

    [Fact]
    public void IsInstanceOf_IPresentationComponent()
    {
        var obj = new TestablePVIObject();

        obj.Should().BeAssignableTo<IPresentationComponent>();
    }

    [Fact]
    public void Presentation_IsNullWhenSlideHasNoPresentation()
    {
        var slide = new BaseSlideStub { Presentation = null };
        var obj = new TestablePVIObject();

        obj.SetParentSlide(slide);

        obj.Presentation.Should().BeNull();
    }
}
