namespace OrcaX.Office.Contracts;

/// <summary>
/// Mirrors @orcax/api-contracts createRuleResult. Shared shape only — not a binary dependency for Office.js.
/// </summary>
public sealed record RuleResult(
    string RuleId,
    string Severity,
    string? Location,
    string? Explanation,
    string? Correction
);

public static class RuleSeverities
{
    public const string Pass = "pass";
    public const string Fail = "fail";
    public const string Warning = "warning";
}
