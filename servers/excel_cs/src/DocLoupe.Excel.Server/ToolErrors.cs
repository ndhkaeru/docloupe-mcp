namespace DocLoupe.Excel.Server;

public sealed record ToolError(string Code, string Message, bool Retryable);

public static class ToolErrors
{
    public static ToolError FromException(Exception error)
    {
        var (code, retryable) = error switch
        {
            InvalidOperationException { Message: "REVISION_CONFLICT" } => ("REVISION_CONFLICT", true),
            InvalidOperationException { Message: "SESSION_BUSY" } => ("SESSION_BUSY", true),
            InvalidOperationException { Message: "UNSAVED_CHANGES" } => ("UNSAVED_CHANGES", false),
            InvalidOperationException { Message: "SOURCE_CHANGED_ON_DISK" } => ("SOURCE_CHANGED_ON_DISK", false),
            InvalidOperationException { Message: "DESTINATION_CHANGED" } => ("DESTINATION_CHANGED", true),
            InvalidDataException exception when exception.Message.StartsWith("MERGED_NON_ORIGIN:", StringComparison.Ordinal) =>
                ("MERGED_NON_ORIGIN", false),
            InvalidDataException exception when exception.Message.StartsWith("AMBIGUOUS_FORMULA_TEXT:", StringComparison.Ordinal) =>
                ("AMBIGUOUS_FORMULA_TEXT", false),
            ArgumentException { Message: "AMBIGUOUS_FORMULA_TEXT" } => ("AMBIGUOUS_FORMULA_TEXT", false),
            InvalidDataException { Message: "RICH_CONTENT_REQUIRES_REPLACE" } => ("RICH_TEXT_POLICY_REQUIRED", false),
            KeyNotFoundException { Message: "Unknown session" } => ("SESSION_NOT_FOUND", false),
            KeyNotFoundException => ("TARGET_NOT_FOUND", false),
            FileNotFoundException or DirectoryNotFoundException => ("FILE_NOT_FOUND", false),
            UnauthorizedAccessException => ("PATH_NOT_ALLOWED", false),
            FormatException exception when exception.Message.StartsWith("Invalid A1 cell address:", StringComparison.Ordinal) =>
                ("INVALID_ADDRESS", false),
            FormatException exception when exception.Message is "Invalid sheet name" or "Invalid quoted sheet name" =>
                ("INVALID_ADDRESS", false),
            NotSupportedException { Message: "Unsupported workbook format" } => ("UNSUPPORTED_FORMAT", false),
            NotSupportedException => ("INVALID_OP", false),
            ArgumentException exception when exception.Message.Contains("exceeds 500 cells", StringComparison.Ordinal) =>
                ("LIMIT_EXCEEDED", false),
            ArgumentException or FormatException => ("INVALID_INPUT", false),
            InvalidDataException or System.Xml.XmlException => ("PACKAGE_INVALID", false),
            _ => ("INTERNAL_ERROR", false)
        };
        return new ToolError(code, error.Message, retryable);
    }
}
