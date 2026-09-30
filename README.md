# Win32 API Demos

Small standalone C# projects exploring how managed code calls into native Windows APIs — Win32, WLAN, Office COM interop — via P/Invoke and interop.

Each folder is a self-contained example from the "Win32 APIs" section of the Programming Advices roadmap (Level 25, Introduction to RESTful APIs — context section).

---

## Examples

| # | Project | API / Library | What it does |
|---|---|---|---|
| 01 | Change WallPaper | `SystemParametersInfo` (user32) | Sets the desktop wallpaper programmatically |
| 02 | Getting Screen Resolution | `GetSystemMetrics` (user32) | Reads screen width/height in pixels |
| 03 | Fetching Battery Information | `GetSystemPowerStatus` (kernel32) | Reports battery charge level and AC status |
| 04 | Shutdown Windows | `ExitWindowsEx` (user32) | Initiates system shutdown / restart |
| 05 | MessageBox | `MessageBox` (user32) | Displays a native Windows message dialog |
| 06 | RunningProcesses | `System.Diagnostics.Process` | Lists active processes on the machine |
| 07 | Audio API | NAudio 3.x (`NAudio.Wasapi`) | Controls the master volume of the default audio endpoint |
| 08 | Wifi Networks | `WlanEnumInterfaces` (Wlanapi) | Enumerates the wireless adapters on the machine |
| 09 | OutlookEmailSender | Outlook COM interop | Sends an email via the locally installed Outlook client |
| 10 | WordDocumentCreator | Word COM interop | Creates a Word document programmatically |
| 11 | ExcelSheetCreator | Excel COM interop | Creates an Excel workbook with data |
| 12 | PowerPointCreator | PowerPoint COM interop | Creates a PowerPoint presentation |

---

## Requirements

- **Windows 10 / 11** — most examples call Windows-only APIs
- **.NET 9 SDK** or later (examples 07, 08 use SDK-style projects targeting `net9.0-windows`)
- **.NET Framework 4.8** (examples 01–06, 09–12 use classic `.csproj` style)
- **Microsoft Office** (for examples 09–12 only) — the COM interop demos require Word, Excel, Outlook, or PowerPoint installed

---

## Building and running

Each folder contains its own solution (`.sln`). Open in Visual Studio or build with the .NET CLI:

```bash
cd "01- Change WallPaper"
dotnet build
dotnet run
