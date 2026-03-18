using System.Collections.Frozen;

namespace Aspose.Slides.Foss.Internal;

/// <summary>
/// Provides bidirectional mapping between OOXML table style GUIDs and <see cref="TableStylePreset"/> values.
/// </summary>
internal static class TableStyleMapping
{
    private static readonly FrozenDictionary<string, TableStylePreset> GuidToPreset = new Dictionary<string, TableStylePreset>
    {
        ["{2D5ABB26-0587-4C30-8999-92F81FD0307C}"] = TableStylePreset.NoStyleNoGrid,
        ["{5940675A-B579-460E-94D1-54222C63F5DA}"] = TableStylePreset.NoStyleTableGrid,
        ["{3C2FFA5D-87B4-456A-9821-1D502468CF0F}"] = TableStylePreset.ThemedStyle1Accent1,
        ["{284085F0-18D9-4B04-A16A-4E4A7013F568}"] = TableStylePreset.ThemedStyle1Accent2,
        ["{69C7853C-536D-4A76-A0AE-DD22B4D04AA4}"] = TableStylePreset.ThemedStyle1Accent3,
        ["{775DCB02-9BB8-47FD-8907-85C794F793BA}"] = TableStylePreset.ThemedStyle1Accent4,
        ["{35758FB7-9AC5-4552-8A53-C91805E547FA}"] = TableStylePreset.ThemedStyle1Accent5,
        ["{08FB837D-C827-4EFA-A057-4D05807E0F7C}"] = TableStylePreset.ThemedStyle1Accent6,
        ["{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}"] = TableStylePreset.ThemedStyle2Accent1,
        ["{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}"] = TableStylePreset.ThemedStyle2Accent2,
        ["{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}"] = TableStylePreset.ThemedStyle2Accent3,
        ["{E269D01E-BC32-4049-B463-5C60D7B0CCD2}"] = TableStylePreset.ThemedStyle2Accent4,
        ["{327F97BB-C833-4FB7-BDE5-3F7075034690}"] = TableStylePreset.ThemedStyle2Accent5,
        ["{638B1855-1B75-4FBE-930C-398BA8C253C6}"] = TableStylePreset.ThemedStyle2Accent6,
        ["{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}"] = TableStylePreset.LightStyle1,
        ["{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}"] = TableStylePreset.LightStyle1Accent1,
        ["{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}"] = TableStylePreset.LightStyle1Accent2,
        ["{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}"] = TableStylePreset.LightStyle1Accent3,
        ["{D27102A9-8310-4765-A935-A1911B00CA55}"] = TableStylePreset.LightStyle1Accent4,
        ["{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}"] = TableStylePreset.LightStyle1Accent5,
        ["{68D230F3-CF80-4859-8CE7-A43EE81993B5}"] = TableStylePreset.LightStyle1Accent6,
        ["{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}"] = TableStylePreset.LightStyle2,
        ["{69012ECD-51FC-41F1-AA8D-1B2483CD663E}"] = TableStylePreset.LightStyle2Accent1,
        ["{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}"] = TableStylePreset.LightStyle2Accent2,
        ["{F2DE63D5-997A-4646-A377-4702673A728D}"] = TableStylePreset.LightStyle2Accent3,
        ["{17292A2E-F333-43FB-9621-5CBBE7FDCDCD}"] = TableStylePreset.LightStyle2Accent4,
        ["{5A111915-BE36-4E01-A7E5-04B1672EAD32}"] = TableStylePreset.LightStyle2Accent5,
        ["{912C8C85-51F0-491E-9774-3900AFEF0FD7}"] = TableStylePreset.LightStyle2Accent6,
        ["{616DA210-FB22-4F2B-B426-AB7765F5A3C5}"] = TableStylePreset.LightStyle3,
        ["{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}"] = TableStylePreset.LightStyle3Accent1,
        ["{5DA37D80-6434-44D0-A028-1B22A696006F}"] = TableStylePreset.LightStyle3Accent2,
        ["{8799B23B-EC83-4686-B30A-512413B5E67A}"] = TableStylePreset.LightStyle3Accent3,
        ["{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}"] = TableStylePreset.LightStyle3Accent4,
        ["{BDBED569-4797-4DF6-BD98-5491D3F75D02}"] = TableStylePreset.LightStyle3Accent5,
        ["{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}"] = TableStylePreset.LightStyle3Accent6,
        ["{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}"] = TableStylePreset.MediumStyle1,
        ["{B301B821-A1FF-4177-AEE7-76D212191A09}"] = TableStylePreset.MediumStyle1Accent1,
        ["{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}"] = TableStylePreset.MediumStyle1Accent2,
        ["{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}"] = TableStylePreset.MediumStyle1Accent3,
        ["{1E171933-4619-4E11-9A3F-F7608DF75F80}"] = TableStylePreset.MediumStyle1Accent4,
        ["{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}"] = TableStylePreset.MediumStyle1Accent5,
        ["{10A1B5F5-9B99-4C35-A422-299274C87571}"] = TableStylePreset.MediumStyle1Accent6,
        ["{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}"] = TableStylePreset.MediumStyle2,
        ["{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"] = TableStylePreset.MediumStyle2Accent1,
        ["{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}"] = TableStylePreset.MediumStyle2Accent2,
        ["{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}"] = TableStylePreset.MediumStyle2Accent3,
        ["{00A15C55-8517-42AA-B614-E9B94910E393}"] = TableStylePreset.MediumStyle2Accent4,
        ["{7DF18680-E054-41AD-8BC1-D1AEF088D02A}"] = TableStylePreset.MediumStyle2Accent5,
        ["{93296810-A885-4BE3-A3E7-6D5BEEA58F35}"] = TableStylePreset.MediumStyle2Accent6,
        ["{8EC20E35-A176-4012-BC5E-935CFFF8708E}"] = TableStylePreset.MediumStyle3,
        ["{6E25E649-3F16-4E02-A733-19D2CDBF48F0}"] = TableStylePreset.MediumStyle3Accent1,
        ["{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}"] = TableStylePreset.MediumStyle3Accent2,
        ["{EB344D84-9AFB-497E-A393-DC336BA19D2E}"] = TableStylePreset.MediumStyle3Accent3,
        ["{EB9631B5-78F2-41C9-869B-9F39066F8104}"] = TableStylePreset.MediumStyle3Accent4,
        ["{C8F8D0E7-142A-4E7B-B5E7-6C5F2E3C5D8A}"] = TableStylePreset.MediumStyle3Accent5,
        ["{CF6FCB3F-D1E5-47BA-9D2E-6C82A4F0D4D7}"] = TableStylePreset.MediumStyle3Accent6,
        ["{D7AC3CCA-C797-4891-BE02-D94E43425B78}"] = TableStylePreset.MediumStyle4,
        ["{69CF1AB2-1976-4502-BF36-3FF5EA218861}"] = TableStylePreset.MediumStyle4Accent1,
        ["{8A107856-5554-42FB-B03E-39F5DBC370BA}"] = TableStylePreset.MediumStyle4Accent2,
        ["{0505E3EF-67EA-436B-97B2-0124C06EBD24}"] = TableStylePreset.MediumStyle4Accent3,
        ["{C4B1156A-380E-4F78-BDF5-A137F74CB7F7}"] = TableStylePreset.MediumStyle4Accent4,
        ["{D37F0D3E-1E42-4205-97A0-B0D6F0B2F22B}"] = TableStylePreset.MediumStyle4Accent5,
        ["{2A488322-F2BA-4B5B-9748-0D474271808F}"] = TableStylePreset.MediumStyle4Accent6,
        ["{E8034E78-7F5D-4C2E-B375-FC64B27BC917}"] = TableStylePreset.DarkStyle1,
        ["{125E5076-3810-47DD-B79F-674D7AD40C01}"] = TableStylePreset.DarkStyle1Accent1,
        ["{37CE84F3-28C3-443E-9E96-99CF82512B78}"] = TableStylePreset.DarkStyle1Accent2,
        ["{D03447BB-5D67-496B-8275-5EBE8813A66C}"] = TableStylePreset.DarkStyle1Accent3,
        ["{E929F9F4-4A8F-4326-A1B4-22849713DDAB}"] = TableStylePreset.DarkStyle1Accent4,
        ["{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}"] = TableStylePreset.DarkStyle1Accent5,
        ["{AF606853-7671-496A-8E4F-DF71F8EC918B}"] = TableStylePreset.DarkStyle1Accent6,
        ["{5202B0CA-FC54-4496-8BCA-5EF66A818D29}"] = TableStylePreset.DarkStyle2,
        ["{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}"] = TableStylePreset.DarkStyle2Accent1AndAccent2,
        ["{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}"] = TableStylePreset.DarkStyle2Accent3AndAccent4,
        ["{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}"] = TableStylePreset.DarkStyle2Accent5AndAccent6,
    }.ToFrozenDictionary(StringComparer.OrdinalIgnoreCase);

    private static readonly FrozenDictionary<TableStylePreset, string> PresetToGuid;

    static TableStyleMapping()
    {
        var dict = new Dictionary<TableStylePreset, string>(GuidToPreset.Count);
        foreach (var (guid, preset) in GuidToPreset)
            dict.TryAdd(preset, guid);
        PresetToGuid = dict.ToFrozenDictionary();
    }

    /// <summary>
    /// Resolves a table style GUID to its <see cref="TableStylePreset"/> value.
    /// Returns <see cref="TableStylePreset.Custom"/> if the GUID is unrecognized.
    /// </summary>
    internal static TableStylePreset FromGuid(string? guid)
    {
        if (string.IsNullOrWhiteSpace(guid))
            return TableStylePreset.None;
        return GuidToPreset.TryGetValue(guid, out var preset) ? preset : TableStylePreset.Custom;
    }

    /// <summary>
    /// Resolves a <see cref="TableStylePreset"/> to its OOXML GUID string.
    /// Returns <c>null</c> if the preset has no GUID mapping.
    /// </summary>
    internal static string? ToGuid(TableStylePreset preset)
    {
        return PresetToGuid.TryGetValue(preset, out var guid) ? guid : null;
    }
}
