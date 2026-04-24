namespace DataRefineX.Models;

public sealed record ProgressUpdate(
    int FileIndex,
    int FileTotal,
    string? CurrentFile,
    long TotalRowsRead,
    long DuplicatesRemoved,
    long InvalidRemoved,
    long ValidRecords,
    double OverallPercent);
