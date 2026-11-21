# AriaDamInvoice

Windows desktop application for creating and managing livestock sale contracts for Aria Dam Khatoun.

## Technology
- WPF / .NET 10
- C#
- ClosedXML for editing Excel workbooks
- CommunityToolkit.Mvvm

## Current status
- Initial WPF application shell
- First-run output-folder selection
- Persisted output-folder setting in `%AppData%\AriaDamInvoice\settings.json`
- Basic dashboard entry points for new and previous contracts (workflows are placeholders)

## Build
Open `AriaDamInvoice.slnx` with Visual Studio 2022 / a compatible .NET 10 SDK, or run:

```powershell
dotnet restore
dotnet build
dotnet run --project AriaDamInvoice.App
```

## Planned
- Add the approved Excel contract template
- Generate variable-length animal and ear-tag tables
- Save editable Excel contracts and export matching PDFs
- Browse, reopen, and edit existing contracts
