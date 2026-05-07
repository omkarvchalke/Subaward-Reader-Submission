namespace SubawardReader;

public sealed record SubawardRecord(
    string FileName,
    string SubrecipientName,
    decimal Amount
);
