using System.Reflection;

namespace Aspose.Slides.Foss.Internal.Pptx;

/// <summary>
/// PPTX template loading.
/// Loads the Template.pptx file for new presentations.
/// </summary>
public static class Template
{
    /// <summary>
    /// Path to the Template.pptx file (in the same directory as the executing assembly).
    /// </summary>
    public static readonly string TemplatePath = Path.Combine(
        Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? AppContext.BaseDirectory,
        "Template.pptx");

    /// <summary>
    /// Gets the path to the Template.pptx file.
    /// </summary>
    /// <returns>Absolute path to Template.pptx.</returns>
    /// <exception cref="FileNotFoundException">If the template file doesn't exist.</exception>
    public static string GetTemplatePath()
    {
        if (!File.Exists(TemplatePath))
        {
            throw new FileNotFoundException(
                $"Template.pptx not found at {TemplatePath}. Please ensure the template file exists.",
                TemplatePath);
        }

        return TemplatePath;
    }

    /// <summary>
    /// Loads the Template.pptx into the given package.
    /// Copies all parts from the template into the target package.
    /// </summary>
    /// <param name="package">The OPC package to populate with template contents.</param>
    /// <exception cref="FileNotFoundException">If Template.pptx doesn't exist.</exception>
    public static void LoadTemplate(Opc.OpcPackage package)
    {
        var templatePath = GetTemplatePath();
        var templatePackage = Opc.OpcPackage.Open(templatePath);

        foreach (var partName in templatePackage.GetPartNames())
        {
            var content = templatePackage.GetPart(partName);
            if (content is not null)
            {
                package.SetPart(partName, content);
            }
        }
    }
}
