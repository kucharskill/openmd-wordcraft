# OpenMD for WordCraft

WordCraft, part of the ArtCraft suite, is a great application. Its Markdown support is still at an early stage, however: Markdown export does not save image files, and YAML front matter is displayed as document content and can be damaged when the formatted document is saved. OpenMD is a temporary workaround that makes it possible to work with Markdown files in WordCraft while its native Markdown support is still incomplete.

Edit Markdown in WordCraft through a temporary DOCX, then synchronize saved changes back to the original Markdown file. OpenMD preserves YAML front matter, retains HTML comments, restores image paths, and optionally commits saved changes to a local Git repository.

A small Windows application with no console or main window. File selection, conflict dialogs, and a notification-area icon provide the interface.

Source is available on [GitHub](https://github.com/kucharskill/openmd-wordcraft). Download the [latest portable executable](https://github.com/kucharskill/openmd-wordcraft/releases/latest/download/openmd.exe) from [GitHub Releases](https://github.com/kucharskill/openmd-wordcraft/releases), or build it from source. No ZIP extraction or installer is needed.

## Requirements

- Windows with .NET Framework 4.8. No .NET SDK, Rust, or Visual Studio is needed to run the launcher.
- [WordCraft](https://github.com/storytold/wordcraft) with DOCX support.
- [Pandoc](https://pandoc.org/) available on PATH or selected at first launch. Verified with Pandoc 3.9.0.2 and WordCraft 0.18.1.
- Optional: Git on PATH and a configured author identity for automatic local commits.

OpenMD does not bundle or install WordCraft or Pandoc. It uses no hosted API or paid service. Pandoc may access remote resources referenced by the document.

## Quick start

Download `openmd.exe`, then drag an `.md` file onto it, choose OpenMD in Windows **Open with**, or run:

```powershell
.\openmd.exe "C:\Documents\notes.md"
```

Running without a document opens a file picker. OpenMD does not change file associations automatically.

The launcher detects the executable of a running WordCraft process. Otherwise, it uses a saved path or asks you to locate `wordcraft.exe`. Pandoc is resolved from PATH, the saved configuration, or a file picker. Canceling either picker ends the operation.

Explicit executable paths are also supported:

```powershell
.\openmd.exe "C:\Documents\notes.md" --wordcraft "C:\Apps\WordCraft\wordcraft.exe" --pandoc "C:\Tools\pandoc.exe"
```

When building from a source checkout, the output is `WordCraft\openmd.exe`. Generated executables are excluded from Git and distributed as release assets.

## Editing and synchronization

1. OpenMD keeps the original YAML and HTML comments outside the editable content, records local image paths, and converts the document to DOCX.
2. WordCraft opens the temporary DOCX. Save normally in WordCraft to update the source Markdown.
3. After a stable DOCX save, OpenMD converts it back, extracts images, restores metadata and comments, and atomically replaces Markdown.
4. OpenMD exits after its WordCraft process exits or its document is confirmed closed. Unsaved changes discarded in WordCraft cannot be recovered by the launcher.

**Save As to a different DOCX path does not redirect monitoring.** Continue saving the original temporary DOCX for synchronization.

If Markdown was edited externally, OpenMD asks whether to overwrite it. Declining keeps the external file and retains the recovered candidate. Accepting preserves a copy of the replaced file in the session directory. Headless diagnostics never overwrite a conflict.

## Metadata, comments, and images

- YAML front matter is retained verbatim. UTF-8 BOM and newline style are preserved.
- YAML metadata is also passed to Pandoc: `lang: pl-PL` sets Polish document and proofing language, while title, author, and other supported fields populate DOCX properties. An embedded body-only XML template prevents an extra visible title/author/date block. It is extracted to `%TEMP%\openmd-wordcraft\templates\body-only.docx.xml` and restored if missing or modified. No separate template file is needed beside the executable.
- HTML comments are collected in source order and placed after YAML, before the content. One generated HTML comment explains their relocation. Repeated opening does not duplicate this notice. Comment examples in code blocks remain code.
- Word comments use Pandoc `.comment-start` and `.comment-end` spans. Highlights use `markdown+mark`, including `==highlighted text==`.
- Unchanged images are recognized by SHA-256 and mapped back to their original paths. Identical images originally stored at different paths use the first recorded match.
- New or modified images are written into the first recorded relative image directory, or `<document-stem>_img` if no such directory was recorded. Filenames include a content hash.
- Removing an image from the document never deletes its original file. Existing image files are not overwritten with different content.
- If at least one input table uses pipe syntax, the Markdown writer is configured to prefer pipe tables. Otherwise, Pandoc chooses its normal table representation.

## Configuration and recovery

Configuration is stored in `%TEMP%\openmd-wordcraft\config.json`.

Working files use `%TEMP%\openmd-wordcraft\sessions\<source-path-hash>\<document-stem>.docx`. Reopening the same source reuses this path. If the DOCX is locked or read-only, a fallback directory adds a unique suffix while preserving the document name.

The notification-area menu opens the session directory. It contains `session.json`, the original Markdown, export candidates, previous versions, and logs. **Windows TEMP is disposable**: copy recovery files elsewhere before clearing it. These files contain document content and should not be published.

## Optional Git commits

After a save that changes Markdown bytes, OpenMD checks whether Git is available and the document is already tracked. It attempts a local commit of the document and applicable extracted images, using this message:

```text
Apply user changes made in an external editor (WordCraft)
```

A separate index excludes unrelated staged changes. OpenMD does not push. Missing Git, untracked documents, and failed commits do not display an error or interrupt saving. Commit failures are recorded in the session's `git.log`. An existing Git author identity is required.

## Build and verification

```powershell
.\WordCraft\build.cmd
.\WordCraft\test.ps1 -WordCraft "C:\Apps\WordCraft\wordcraft.exe"
```

The build uses the Windows .NET Framework `csc.exe` and standard framework assemblies, without downloading packages. The test suite looks for `wordcraft-cli.exe` beside WordCraft; use `-WordCraftCli PATH` to override it.

Synthetic integration tests cover metadata, HTML and Word comments, highlights, image addition and removal, fallback directories, Unicode paths, save conflicts, stable DOCX paths, locked-file fallback, and live GUI save synchronization through WordCraft's authenticated localhost control channel. Fixtures stay in the repository's `.tmp/`; runtime sessions use Windows TEMP. No production documents are needed.

GitHub Releases distribute the single `openmd.exe` file. The release notes identify the version and SHA-256 checksum. The embedded template requires no companion file.

## Diagnostic options

```powershell
.\openmd.exe --version --headless --result version.json
.\openmd.exe document.md --prepare --headless --result prepared.json
.\openmd.exe --sync "C:\Temp\session.json" --headless
```

`--prepare` creates a working DOCX without launching WordCraft. The result JSON identifies the session and DOCX paths. `--sync` exports a changed DOCX from an existing session. `--headless` disables dialogs; errors return exit code 1, success returns 0. `--result` writes a JSON result or error.

## Limitations

This is conversion through DOCX, not a byte-preserving Markdown editor. Lists, quotes, whitespace, image dimensions, and unsupported Markdown extensions may change. Comments and highlights survive the tested cases, but replies, overlapping annotations, and tracked revisions need additional verification. DOCX reading uses `--track-changes=all` to retain comments, so revision spans can also appear.

WordCraft editing can affect paragraph styles when deleting paragraph boundaries. Images inserted inside code-styled paragraphs may be omitted during conversion. Inspect important documents after editing and keep a versioned source or backup.

## License

Licensed under the [MIT License](LICENSE). WordCraft and Pandoc are separate products distributed under their own licenses.
