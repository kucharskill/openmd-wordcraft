# OpenMD for WordCraft source

See the [project README](../README.md) for setup, behavior, configuration, and limitations.

```powershell
.\build.cmd
.\test.ps1 -WordCraft "C:\Apps\WordCraft\wordcraft.exe"
```

- `OpenMd.cs`: C# 5 / .NET Framework Windows application.
- `body-only.docx.xml`: XML template embedded during compilation, extracted to Windows TEMP at runtime.
- `openmd.exe`: generated portable launcher, excluded from Git and available through GitHub Releases.
- `build.cmd`: compile with the .NET Framework compiler.
- `test.ps1`: synthetic conversion and live synchronization checks.

Licensed under the [MIT License](../LICENSE).
