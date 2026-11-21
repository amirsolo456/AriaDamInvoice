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
- The WPF windows are resizable and their content can scroll on smaller displays.

## Current implementation
- Persian WPF dashboard and persistent output-folder selection.
- Contract entry window with buyer, phone, address, and contract date.
- Add/remove livestock rows independently, with automatic row numbering and line/overall totals.
- Enter ear tags one per line; the ear-tag count is independent of livestock rows.
- Contract notes field.

The current Save button is still a placeholder: Excel template population, persistence of the contract, and PDF export are not connected yet.

## Build and run
Requires Windows with the .NET 10 SDK and Windows Desktop targeting support.

```powershell
dotnet restore
dotnet build
dotnet run --project AriaDamInvoice.App
```

## Next steps
1. Connect the approved Excel template and map its cells.
2. Save a contract as an editable Excel file.
3. Export a matching PDF with variable-length tables and page layout.
4. Browse, reopen, and edit previous contracts.


## Fonts and dates
- Persian text uses the Windows-installed **B Nazanin** font.
- Numeric fields and calculated values use the Windows-installed **B Yekan** font.
- Contract date entry uses the Solar Hijri calendar (Jalali), in `yyyy/MM/dd` format. Persian and Arabic-Indic digits are accepted; the date is validated and converted internally to a Gregorian `DateTime` for processing.
- Both fonts should be installed on any Windows computer where the application is run. If they are missing, Windows will substitute another font.


## Fixed livestock sale obligations
The contract entry screen no longer asks the user to enter free-form contract notes. The fixed legal/obligations section is defined in `Models/ContractText.cs`:
- Excel section title: `تعهدات و بند قوانین فروش دام زنده`
- Body: the fixed text in `ContractText.ObligationsBody`

When the Excel template population step is implemented, this constant must be written into the matching obligations area in the reference workbook. The reference Excel template is not currently stored in this repository, so the workbook itself has not yet been edited.
