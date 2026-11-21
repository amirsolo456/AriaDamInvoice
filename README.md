# AriaDamInvoice

Windows desktop application for creating and managing livestock sale contracts for Aria Dam Khatoun.

## Technology
- WPF / .NET 10
- C#
- ClosedXML for editing Excel workbooks
- CommunityToolkit.Mvvm

## Portability and project conventions
- The application root namespace and assembly name are explicitly set to `AriaDamInvoice`; they do not depend on the local folder or project filename.
- No user-specific Windows paths are stored in source code.
- User settings are stored under the current Windows user's `%AppData%\AriaDamInvoice\settings.json`.
- The output folder is selected by the user. If the saved folder is unavailable on another computer, the application asks the user to select a valid folder again.
- The WPF window is resizable and its content can scroll on smaller displays.

## Current foundation
- Persian WPF application shell
- First-run output-folder selection
- Persistent output-folder setting
- Dashboard entry points for new and previous contracts (workflow placeholders)

## Build and run
Requires a Windows machine with the .NET 10 SDK and Windows Desktop targeting support.

```powershell
dotnet restore
dotnet build
dotnet run --project AriaDamInvoice.App
```

## Planned implementation
1. Add the approved Excel contract template.
2. Create contract data models and input screens.
3. Generate variable-length animal and ear-tag tables.
4. Save editable Excel contracts and export matching PDFs.
5. Browse, reopen, and edit existing contracts.
