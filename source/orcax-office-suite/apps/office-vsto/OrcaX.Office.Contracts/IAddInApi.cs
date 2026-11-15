namespace OrcaX.Office.Contracts;

public interface IAddInApi
{
    Task<AddInBootstrap> BootstrapWordAsync(CancellationToken cancellationToken = default);

    Task<AddInBootstrap> BootstrapExcelAsync(CancellationToken cancellationToken = default);

    Task RecordWordValidationAsync(ValidationRunRequest request, CancellationToken cancellationToken = default);

    Task RecordExcelPrintAsync(PrintRunRequest request, CancellationToken cancellationToken = default);
}

public sealed record AddInBootstrap(
    string Host,
    string WorkspaceId,
    bool EntitlementAllowed,
    IReadOnlyList<ProfileSummary> Profiles
);

public sealed record ProfileSummary(string Id, string Code, string Name, int CurrentVersion);

public sealed record ValidationRunRequest(
    Guid ProfileId,
    Guid ProfileVersionId,
    int ProfileVersion,
    string ProfileCode,
    int Pass,
    int Fail,
    int Warning,
    int Corrected
);

public sealed record PrintRunRequest(
    Guid ProfileId,
    Guid ProfileVersionId,
    int ProfileVersion,
    string ProfileCode,
    string PaperSize,
    string Orientation,
    bool Applied,
    bool PdfExported
);
