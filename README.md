# ExcelTrainingMonitor

ExcelTrainingMonitor is a Windows desktop application for monitoring training-status spreadsheets and maintaining editable Excel-based records.

## Download The Prebuilt Application

Prebuilt Windows releases are available from the repository's [GitHub Releases](https://github.com/SuperEgg9562/ExcelTrainingMonitor/releases) page.

Download either:

- `ExcelTrainingMonitor-<version>-win-x64.exe` for the standalone executable.
- `ExcelTrainingMonitor-<version>-win-x64.zip` for the complete portable package including theme assets.

## Build From Source

Requirements:

- Windows 10 or Windows 11 x64
- Visual Studio 2022
- The **.NET desktop development** Visual Studio workload
- .NET 8 SDK

Build steps:

1. Clone or download this repository.
2. Open `src/ExcelTrainingMonitor.slnx` in Visual Studio 2022.
3. Allow Visual Studio to restore the ClosedXML NuGet dependency.
4. Select the `Release` configuration.
5. Use **Build > Build Solution**.

Normal Visual Studio build output is generated under `src/bin/`.

To create the same self-contained Windows x64 package used for prebuilt releases:

1. Right-click the `ExcelTrainingMonitor` project in Solution Explorer.
2. Select **Publish**.
3. Select the `Windows-x64` profile.
4. Click **Publish**.

The publish profile writes the executable and required theme assets to `prebuilt/windows-x64/`.

## Features

- Watches Excel workbooks for training-status changes.
- Displays training alerts, history, dashboard progress, and charts.
- Creates and edits GridBook, Compliance Plan, Process Record, and Daily Production grids.
- Saves editable grids to Excel workbooks.
- Supports printing, chart export, reminders, and optional ntfy notifications.
