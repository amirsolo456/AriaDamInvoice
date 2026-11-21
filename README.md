# AriaDamInvoice

Windows desktop application for creating livestock sale contracts for Aria Dam Khatoun.

## Technology

- WPF / .NET 10
- C#
- ClosedXML for editing the selected Excel template
- Microsoft Excel desktop application for PDF export through Excel COM automation

## First run and saved settings

On first launch, the application asks the user to select:
1. The output folder for generated Excel and PDF contracts.
2. The original Excel contract template (`.xlsx`).

These paths and other preferences are stored in:

`%AppData%\\AriaDamInvoice\\settings.json`

If the application is moved to another computer and the old paths are unavailable, it asks the user to select the two locations again. The settings file is not required to move the application.

## Portable deployment (recommended)

A self-contained Windows x64 publish folder can be created without requiring the .NET Desktop Runtime to be installed on the destination computer.

On a Windows development computer with the .NET 10 SDK, run from the repository root:

```powershell
.\\scripts\\Publish-Portable.ps1
```

The script creates:

`artifacts\\AriaDamInvoice-Portable\\`

Copy the entire `AriaDamInvoice-Portable` folder to the new computer. Do not copy only the EXE; all files in the publish folder are required.

Also copy the exact original Excel template workbook separately. The workbook is intentionally not embedded in the application or stored in this repository, because the user chooses the trusted original template on first run.

### Destination computer requirements

- Windows 10/11 x64.
- Microsoft Excel desktop installed and working. The current PDF exporter automates Excel; Excel Viewer or a browser spreadsheet is not sufficient.
- Install the same fonts used by the contract: B Nazanin and B Yekan. Missing fonts may change line wrapping, sizes, or the PDF appearance.
- No separate .NET runtime installation is needed when using the self-contained publish folder.

### First launch on the new computer

1. Start `AriaDamInvoice.exe` from the copied folder.
2. Choose the destination folder for contract outputs.
3. Choose the exact original Excel template file.
4. Save and continue.

The selected paths are saved in the new Windows user's AppData settings. They are not hard-coded to the old computer's drive letters or user profile.

### Optional: preserve numbering and custom breed options

If the next document number and custom breed list should continue from the old computer, copy this file from the old computer:

`%AppData%\\AriaDamInvoice\\settings.json`

Place it at the same path under the new Windows user profile **before the first launch**. The old absolute folder/template paths may no longer exist, so select the new paths when the application asks. The next document number and breed options remain in the settings.

If you do not copy this optional file, the new installation starts with the default document number 1000 and default breed options.

## Current application features

- Persian WPF dashboard and persistent output/template path selection.
- Contract entry for buyer, phone, address, Solar Hijri date, and document number.
- Document number increment/decrement controls and automatic next-number persistence.
- Add/remove livestock rows with breed selection, totals, and thousands separators.
- Independently editable ear-tag table with one tag per row.
- Fixed obligations text from `Models/ContractText.cs`.
- Creates an editable Excel copy from the selected reference workbook and exports a matching PDF.
- Opens the generated PDF after successful export.

## Excel template compatibility

The selected workbook must be the same approved reference template used during development. The application expects the worksheet named `قرارداد فروش دام` and the cell/row layout encoded in `Services/ExcelContractService.cs`. Selecting an arbitrary Excel workbook is not supported. Keep the original workbook unchanged and make a backup before replacing it.

## Development build

Requires Windows and the .NET 10 SDK:

```powershell
dotnet restore
dotnet build
dotnet run --project AriaDamInvoice.App
```
