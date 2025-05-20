# Drill Namer

This repository contains a Windows Forms add-in for AutoCAD to manage drill naming.

## Building

1. Open **Drill Namer.sln** in Visual Studio 2019 or later.
2. Restore NuGet packages when prompted.
3. Ensure AutoCAD 2022 is installed so referenced DLLs can be resolved.
4. Build the solution using the desired configuration.

The application uses Fody for assembly weaving and NLog for logging. Logging settings are stored in `Drill Namer/NLog.config`.
