# OpenMD for WordCraft source

See the [project README](../README.md) for setup, behavior, configuration, and limitations.

```powershell
.\build.cmd
.\test.ps1 -WordCraft "C:\Apps\WordCraft\wordcraft.exe"
.\package.ps1
```

- `OpenMd.cs`: C# 5 / .NET Framework Windows application.
- `body-only.docx.xml`: XML template embedded during compilation, extracted to Windows TEMP at runtime.
- `openmd.exe`: portable launcher, with no console or main window.
- `build.cmd`: compile with the .NET Framework compiler.
- `test.ps1`: synthetic conversion and live synchronization checks.
- `package.ps1`: create an allowlisted portable ZIP and checksum.

Licensed under the [MIT License](../LICENSE).
