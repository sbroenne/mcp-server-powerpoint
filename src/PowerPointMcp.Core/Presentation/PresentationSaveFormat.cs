using System.Text.Json.Serialization;

namespace Sbroenne.PowerPointMcp.Core.Presentation;

/// <summary>Supported presentation formats for Save As.</summary>
public enum PresentationSaveFormat
{
    /// <summary>Infer the format from the destination extension.</summary>
    [JsonStringEnumMemberName("auto")]
    Auto,

    /// <summary>Open XML presentation (.pptx).</summary>
    [JsonStringEnumMemberName("pptx")]
    Pptx,

    /// <summary>Macro-enabled Open XML presentation (.pptm).</summary>
    [JsonStringEnumMemberName("pptm")]
    Pptm,

    /// <summary>Legacy binary presentation (.ppt).</summary>
    [JsonStringEnumMemberName("ppt")]
    Ppt
}
