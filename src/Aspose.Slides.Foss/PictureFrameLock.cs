using System.Xml.Linq;
using Aspose.Slides.Foss.Internal;

namespace Aspose.Slides.Foss;

/// <summary>
/// Determines which operations are disabled on the parent picture frame.
/// </summary>
public sealed class PictureFrameLock : BaseShapeLock, IPictureFrameLock
{
    private XElement? _picLocks;
    private SlidePart? _slidePart;

    /// <summary>
    /// Initializes internal state with the picLocks XML element and slide part.
    /// </summary>
    /// <param name="picLocks">The a:picLocks element, or <c>null</c> if absent.</param>
    /// <param name="slidePart">The <see cref="SlidePart"/> for saving changes.</param>
    internal void InitInternal(XElement? picLocks, SlidePart? slidePart)
    {
        _picLocks = picLocks;
        _slidePart = slidePart;
    }

    /// <inheritdoc/>
    public bool GroupingLocked
    {
        get => GetLock("noGrp");
        set => SetLock("noGrp", value);
    }

    /// <inheritdoc/>
    public bool SelectLocked
    {
        get => GetLock("noSelect");
        set => SetLock("noSelect", value);
    }

    /// <inheritdoc/>
    public bool RotationLocked
    {
        get => GetLock("noRot");
        set => SetLock("noRot", value);
    }

    /// <inheritdoc/>
    public bool AspectRatioLocked
    {
        get => GetLock("noChangeAspect");
        set => SetLock("noChangeAspect", value);
    }

    /// <inheritdoc/>
    public bool PositionLocked
    {
        get => GetLock("noMove");
        set => SetLock("noMove", value);
    }

    /// <inheritdoc/>
    public bool SizeLocked
    {
        get => GetLock("noResize");
        set => SetLock("noResize", value);
    }

    /// <inheritdoc/>
    public bool EditPointsLocked
    {
        get => GetLock("noEditPoints");
        set => SetLock("noEditPoints", value);
    }

    /// <inheritdoc/>
    public bool AdjustHandlesLocked
    {
        get => GetLock("noAdjustHandles");
        set => SetLock("noAdjustHandles", value);
    }

    /// <inheritdoc/>
    public bool ArrowheadsLocked
    {
        get => GetLock("noChangeArrowheads");
        set => SetLock("noChangeArrowheads", value);
    }

    /// <inheritdoc/>
    public bool ShapeTypeLocked
    {
        get => GetLock("noChangeShapeType");
        set => SetLock("noChangeShapeType", value);
    }

    /// <inheritdoc/>
    public bool CropLocked
    {
        get => GetLock("noCrop");
        set => SetLock("noCrop", value);
    }

    /// <inheritdoc/>
    public bool NoLocks
    {
        get
        {
            if (_picLocks is null)
                return true;

            ReadOnlySpan<string> lockAttrs =
            [
                "noGrp", "noSelect", "noRot", "noChangeAspect", "noMove",
                "noResize", "noEditPoints", "noAdjustHandles", "noChangeArrowheads",
                "noChangeShapeType", "noCrop"
            ];

            foreach (var attr in lockAttrs)
            {
                if (_picLocks.Attribute(attr)?.Value is "1")
                    return false;
            }

            return true;
        }
    }

    /// <summary>
    /// Reads a lock attribute from the picLocks element.
    /// </summary>
    /// <param name="attrName">The XML attribute name.</param>
    /// <returns><c>true</c> if the attribute value is "1"; otherwise <c>false</c>.</returns>
    internal bool GetLock(string attrName)
    {
        if (_picLocks is null)
            return false;
        return _picLocks.Attribute(attrName)?.Value is "1";
    }

    /// <summary>
    /// Writes a lock attribute to the picLocks element.
    /// </summary>
    /// <param name="attrName">The XML attribute name.</param>
    /// <param name="value">The lock value to set.</param>
    internal void SetLock(string attrName, bool value)
    {
        if (_picLocks is null)
            return;

        if (value)
            _picLocks.SetAttributeValue(attrName, "1");
        else
            _picLocks.Attribute(attrName)?.Remove();
    }
}
