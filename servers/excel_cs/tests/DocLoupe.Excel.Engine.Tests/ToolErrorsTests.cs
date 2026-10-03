using DocLoupe.Excel.Server;
using Xunit;

namespace DocLoupe.Excel.Engine.Tests;

public sealed class ToolErrorsTests
{
    [Fact]
    public void KnownFailuresExposeActionableCodesAndRetryability()
    {
        (Exception Error, string Code, bool Retryable)[] cases =
        [
            (new InvalidOperationException("REVISION_CONFLICT"), "REVISION_CONFLICT", true),
            (new InvalidOperationException("SESSION_BUSY"), "SESSION_BUSY", true),
            (new InvalidOperationException("UNSAVED_CHANGES"), "UNSAVED_CHANGES", false),
            (new InvalidDataException("MERGED_NON_ORIGIN: B1 belongs to A1:B1"), "MERGED_NON_ORIGIN", false),
            (new InvalidDataException("AMBIGUOUS_FORMULA_TEXT: use set_formula"), "AMBIGUOUS_FORMULA_TEXT", false),
            (new ArgumentException("AMBIGUOUS_FORMULA_TEXT"), "AMBIGUOUS_FORMULA_TEXT", false),
            (new InvalidDataException("RICH_CONTENT_REQUIRES_REPLACE"), "RICH_TEXT_POLICY_REQUIRED", false),
            (new KeyNotFoundException("Unknown session"), "SESSION_NOT_FOUND", false),
            (new FormatException("Invalid A1 cell address: ZZ0"), "INVALID_ADDRESS", false),
            (new ArgumentException("bad input"), "INVALID_INPUT", false),
            (new ArgumentException("Read exceeds 500 cells"), "LIMIT_EXCEEDED", false),
            (new NotSupportedException("Unsupported workbook format"), "UNSUPPORTED_FORMAT", false),
            (new InvalidDataException("Malformed workbook"), "PACKAGE_INVALID", false),
            (new Exception("unexpected"), "INTERNAL_ERROR", false)
        ];
        foreach (var (error, code, retryable) in cases)
        {
            var mapped = ToolErrors.FromException(error);
            Assert.Equal(code, mapped.Code);
            Assert.Equal(retryable, mapped.Retryable);
            Assert.Equal(error.Message, mapped.Message);
        }
    }
}
