namespace Aspose.Slides.Foss;

/// <summary>
/// Determines which editing operations are disabled on a picture frame.
/// </summary>
public interface IPictureFrameLock
{
    /// <summary>
    /// Gets or sets whether grouping is locked.
    /// </summary>
    bool GroupingLocked { get; set; }

    /// <summary>
    /// Gets or sets whether selection is locked.
    /// </summary>
    bool SelectLocked { get; set; }

    /// <summary>
    /// Gets or sets whether rotation is locked.
    /// </summary>
    bool RotationLocked { get; set; }

    /// <summary>
    /// Gets or sets whether aspect ratio is locked.
    /// </summary>
    bool AspectRatioLocked { get; set; }

    /// <summary>
    /// Gets or sets whether position is locked.
    /// </summary>
    bool PositionLocked { get; set; }

    /// <summary>
    /// Gets or sets whether size is locked.
    /// </summary>
    bool SizeLocked { get; set; }

    /// <summary>
    /// Gets or sets whether edit points are locked.
    /// </summary>
    bool EditPointsLocked { get; set; }

    /// <summary>
    /// Gets or sets whether adjust handles are locked.
    /// </summary>
    bool AdjustHandlesLocked { get; set; }

    /// <summary>
    /// Gets or sets whether arrowheads are locked.
    /// </summary>
    bool ArrowheadsLocked { get; set; }

    /// <summary>
    /// Gets or sets whether shape type is locked.
    /// </summary>
    bool ShapeTypeLocked { get; set; }

    /// <summary>
    /// Gets or sets whether cropping is locked.
    /// </summary>
    bool CropLocked { get; set; }

    /// <summary>
    /// Gets whether all locks are disabled (no lock attributes are set).
    /// </summary>
    bool NoLocks { get; }
}
