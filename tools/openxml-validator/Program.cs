using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

if (args.Length == 0)
{
    WriteDiagnostic(new("usage", "No DOCX path was supplied.", null, null,
        "Pass one or more generated .docx files to the validator."));
    return 2;
}

var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
var failed = false;

foreach (var path in args)
{
    try
    {
        using var document = WordprocessingDocument.Open(path, false);
        var errors = validator.Validate(document).ToList();
        foreach (var error in errors)
        {
            WriteDiagnostic(new(
                error.Id ?? error.ErrorType.ToString(),
                error.Description,
                error.Part?.Uri.ToString(),
                error.Path?.XPath,
                SuggestFix(error)));
        }

        if (errors.Count == 0)
        {
            Console.WriteLine(JsonSerializer.Serialize(new
            {
                valid = true,
                file = path,
                validator = "DocumentFormat.OpenXml 3.5.1",
                format = "Microsoft365"
            }));
        }
        else
        {
            failed = true;
        }
    }
    catch (Exception exception) when (exception is OpenXmlPackageException
        or IOException or UnauthorizedAccessException or InvalidDataException
        or FileFormatException)
    {
        failed = true;
        WriteDiagnostic(new(
            "package-open-failed",
            exception.Message,
            path,
            null,
            "Confirm that the output is an unencrypted DOCX ZIP package with valid OPC content types and relationships."));
    }
}

return failed ? 1 : 0;

static void WriteDiagnostic(Diagnostic diagnostic) =>
    Console.Error.WriteLine(JsonSerializer.Serialize(diagnostic));

static string SuggestFix(ValidationErrorInfo error)
{
    var description = error.Description ?? string.Empty;
    if (description.Contains("relationship", StringComparison.OrdinalIgnoreCase))
    {
        return "Check the source part's .rels file, relationship type, target path, and referenced r:id.";
    }
    if (description.Contains("attribute", StringComparison.OrdinalIgnoreCase))
    {
        return "Check the JSX property mapping and the OOXML attribute type, namespace, and allowed value.";
    }
    if (description.Contains("child element", StringComparison.OrdinalIgnoreCase))
    {
        return "Check the OOXML child-element order and whether this component is legal in its parent.";
    }
    return "Inspect the reported package part and XPath, then compare the emitted markup with ISO/IEC 29500 or a Word-generated reference file.";
}

internal sealed record Diagnostic(
    string Code,
    string? Description,
    string? Part,
    string? XPath,
    string Suggestion);
