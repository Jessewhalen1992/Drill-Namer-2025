# Upgrade Notes

To load `DrillNamer` inside AutoCAD, NETLOAD the DLL matching your AutoCAD version.

- AutoCAD 2014-2015: `net40` build
- AutoCAD 2016-2025: `net48` build

Both assemblies are produced under `bin/Release/<TFM>/DrillNamer.dll` after running `dotnet pack -c Release`.
