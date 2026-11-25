# AX 2012 R3 Reference Extractor

`AxReference.ps1` is a **sanitized context extractor** for Microsoft Dynamics AX 2012 R3: a little bridge between the AOT and modern AI agent harnesses such as GitHub Copilot, Codex, and Claude. AX developers know the drill: the system still needs care, the codebase has years of history, and AX didn’t come with a built-in AI coding buddy. This tool gathers a focused, readable slice of the XPOs so your AI assistant can help untangle the legacy without being handed the whole kingdom.

The extractor reads a TOP export (the full layer, including merged customizations) and a CUS export (customizations only). It finds objects by source keywords and/or object-name fragments, can follow callers and called objects, and writes sanitized X++ files and a manifest explaining why each object was included. It can optionally generate an AX project XPO listing the selected objects.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7.
- AX 2012 R3 XPO exports for the TOP and CUS folders.
- `rg` (ripgrep) is optional. The script uses it to shortlist keyword-hit files when available and falls back to PowerShell otherwise.
- The script uses the Windows-provided `winsqlite3.dll`; no external SQLite package is required.

### Optional: install ripgrep

On Windows, if `winget` is available, install ripgrep from PowerShell:

```powershell
winget install --id BurntSushi.ripgrep.MSVC
```

Close and reopen PowerShell, then confirm it is on `PATH`:

```powershell
rg --version
```

You can also use the installation options on the [official ripgrep download page](https://ripgrep.dev/download/). Installing ripgrep only improves keyword search startup on large exports; the extractor still works without it.

## Getting XPO source folders from AX

The extractor needs two source folders: TOP (the full export for the layer you want to analyze) and CUS (your customizations). This repository includes [AxExportRef.xpp](./AxExportRef.xpp), an optional AX-side helper for exporting AOT objects into XPO files under an AOT-like folder tree. It can help prepare the TOP source folder.

`AxExportRef.xpp` is X++ source code, not an importable XPO. Run it in an AX 2012 development environment: create an AOT class named `AotExportRef`, paste the class declaration and methods from the file into the corresponding class nodes, compile it, then open/run the class (or press F5). Leave `RunOn` as `Called from`; this makes the exporter run on your client and write files to your disk. The source file itself includes the dialog options and supported type keys.

In the dialog, choose an object type, optional name prefix, export root, and whether to export SYS-layer objects only. The defaults are a dry run and SYS-only filtering. Review the dry-run list first; turn off dry run to write the XPOs. The exporter skips files already present, supports limiting the number of new exports, can pause between exports, and writes `_export_log.csv` in the export root. If you need a complete TOP export that includes merged higher-layer customizations, make sure the AX-side layer selection matches that intent; SYS-only mode intentionally omits objects that exist only outside SYS.

The CUS folder should normally be fetched or downloaded from the corresponding Azure DevOps/TFS repository, since it contains your team's customization exports. Custom-only objects can alternatively come from your team's existing source-control XPO export process.

After preparing both folders, pass their paths with `-TopFolder` and `-CusFolder`, or put folders named `top` and `cus` in your working directory. The extractor remembers the last paths entered interactively.

## Set up

Save the script somewhere stable, then dot-source it in the PowerShell session where you want to use the `axref` command:

```powershell
. .\AxReference.ps1
```

To make the command available in every PowerShell session, open your PowerShell profile in Notepad (or your preferred text editor):

```powershell
notepad $PROFILE
```

If PowerShell says the profile file does not exist, create it first:

```powershell
New-Item -ItemType File -Path $PROFILE -Force
notepad $PROFILE
```

Add this line to the profile text file, replacing the example path if you saved the script somewhere else:

```powershell
. 'C:\tools\AxReference.ps1'
```

Save the file, then open a new PowerShell session (or run `. $PROFILE` in the current session) to load the `axref` command.

Run `axref` from a working folder. By default, the extractor looks for folders named `top` and `cus` in the current folder. You can provide different folder names or paths with `-TopFolder` and `-CusFolder`.

## Use it

### Interactive extraction

```powershell
axref
```

The command asks for the TOP and CUS folders, then for search criteria. Enter one or more keywords to find source-code matches, or press Enter to skip keyword search and use object-name fragments alone. Name fragments are case-insensitive substrings; for example, `PurchAgree` matches object names containing that text. If keywords are omitted, caller and called depth default to zero so a name-only run stays focused on the matching objects. You can change either depth when prompted.

It also asks whether to include optional object types (such as jobs, menus, EDTs, and security objects), and whether to create a project XPO. When an output object already exists, answer `y` or `n` for that object, `a` to replace all remaining existing objects, or `s` to skip all remaining existing objects.

If project generation is enabled, the extractor shows the selected objects grouped by AX type in a checkbox-style list. All objects start checked. Enter one or more object numbers (comma or space separated) to toggle them, `all` or `none` to check or clear the current type, Enter to move to the next type, or `done` to finish. For example, enter `1, 4` to toggle objects 1 and 4 in the displayed type. Only checked objects are added to the generated project XPO; the sanitized reference files are still written for every object selected during extraction.

### Search by keyword

```powershell
axref -TopFolder .\top -CusFolder .\cus -Keyword validateWrite -CallerDepth 2 -CalledDepth 1
```

Keywords search XPO source text. The extractor includes matching objects and can follow references up to the caller and called depths. Depth `0` disables that direction of expansion.

### Search by object name only

```powershell
axref -TopFolder .\top -CusFolder .\cus -ObjectNameContains PurchAgree
```

Pass multiple fragments as an array:

```powershell
axref -ObjectNameContains PurchAgree,SalesLine
```

Keyword and name searches can be combined:

```powershell
axref -Keyword validateWrite -ObjectNameContains PurchAgree,SalesLine -CusObjectPrefix Zzz
```

`-CusObjectPrefix` optionally limits the CUS-only object handling to the specified object-name prefix.

### Build or refresh the index only

```powershell
axref -BuildIndex
```

The index is stored under `$HOME\.axref`. Normal extraction runs update and reuse this index, which avoids reparsing unchanged XPO files. Use `-RebuildIndex` on an extraction run to rebuild it from the selected source folders.

## Output

The extractor writes a `reference` folder under the current working directory. It contains sanitized `.xpp` files organized by AX object type, `MANIFEST.md` with the included objects and their inclusion roles, and `AGENTS.md` with guidance for using the reference set. If project generation is enabled and a project template is available, it also writes a project `.xpo` file.

The X++ output is intended for reading and analysis; it is not an importable XPO export. Comments and XPO metadata are removed. When both TOP and CUS versions exist, they are written separately with `_top` and `_cus` suffixes. The generated project XPO is an importable AX project that lists the selected objects; it does not contain their source code.

## Use it with an AI coding agent

Open the generated `reference` folder as the workspace or root folder in your preferred agent harness, such as VS Code with GitHub Copilot, Codex, or Claude. The generated `AGENTS.md` explains how the reference set is organized, and `MANIFEST.md` lists the extracted objects and why they were included. Then describe the issue in the agent chat and ask it to write a Markdown document with a proposed fix.

For example:

> Help me investigate a bug where posting a sales invoice fails when the order contains a specific discount. Use the AX reference files to find the relevant posting flow and likely cause. Create `fix-plan.md` in this workspace with the objects and methods involved, your diagnosis, step-by-step instructions for fixing it, suggested tests, and any assumptions or missing context. Include the proposed code changes as a unified diff and provide complete, copy-paste-ready replacement methods or code blocks. Don't modify the reference files.

The reference folder gives the agent focused AX context to reason from. Use the plan to guide changes in your actual AX development environment or source repository; the sanitized `.xpp` files are for analysis, not direct import into AX.

## Main options

| Option | Purpose |
|---|---|
| `-TopFolder`, `-CusFolder` | Source export folders; defaults to `top` and `cus` in the current directory, or the last used values. |
| `-Keyword` | One or more source-code search terms. Optional when using `-ObjectNameContains`. |
| `-ObjectNameContains` | One or more case-insensitive object-name fragments. |
| `-CusObjectPrefix` | Optional prefix for CUS-only object handling. |
| `-CallerDepth`, `-CalledDepth` | Number of relationship levels to follow in either direction. |
| `-MaxObjects` | Maximum number of objects to include (default `400`). |
| `-BuildIndex` | Build/update the index without extracting a reference set. |
| `-RebuildIndex` | Rebuild the index during extraction. |
| `-Project`, `-ProjectName`, `-ProjectSample` | Generate an AX project XPO, choose its objects interactively by type, set its name, and/or provide a sample project used to learn the project layout. |

To inspect the command syntax after dot-sourcing the script:

```powershell
Get-Command axref -Syntax
```

## Technical design

The extractor uses ripgrep when available to find XPO files containing requested keywords quickly; PowerShell retains a fallback search path. In either case, it parses the XPO files with compiled .NET regular expressions and stores file metadata, parsed object data, methods, and reference edges in a SQLite index under `$HOME\.axref`. The index is keyed to the source roots and reuses cached parses for unchanged files; changed files are reparsed during later runs.

It builds an in-memory model from TOP and CUS entries, resolves object types and references, and uses the CUS data for customizations while retaining TOP data for the full object view. Keyword matches and object-name matches seed selection. Caller traversal uses reverse reference edges; called-object traversal follows references outward. Both honor the configured depth, included types, and object limit.

For each selected object, the exporter writes a sanitized reading copy under `reference/`, preserving separate TOP and CUS variants when both exist. It removes comments and XPO metadata and records inclusion roles in `MANIFEST.md`. When requested, project generation writes a project containing only the objects checked in the interactive type-grouped selection. That generated project is an AX-importable project list; the sanitized `.xpp` files remain analysis copies and are not importable XPO objects.
