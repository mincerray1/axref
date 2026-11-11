Yes. I reviewed AxReference_Optimized.ps1, and I think we can make it considerably faster without sacrificing its existing functionality or requiring a complete rewrite.

My recommendation is to use a hybrid approach: ripgrep (rg) for keyword searching, PowerShell/.NET for XPO parsing and relationship analysis, and a more efficient persistent index for subsequent runs.

This is particularly appropriate for your AX 2012 R3 reference extractor because searching thousands of XPO files and understanding their X++ relationships are two different problems. We should use the fastest tool for each.
1. What I found in your current implementation
Your script already has some good optimizations:
- Compiled regular expressions.
- An incremental cache based on file modification time and size.
- Separate TOP and CUS layers, with CUS methods taking precedence.
- Reverse indexes for finding callers.
- A maximum object limit to prevent relationship expansion from exploding.
- Parsing only changed files rather than rebuilding every parsed object.
However, there are several performance bottlenecks.
Area	Current implementation	Opportunity
File discovery	.NET recursive enumeration	Keep, but optimize
Keyword search	Reads every XPO and uses IndexOf()	Replace with ripgrep
XPO parsing	PowerShell regex and object creation	Optimize selectively
Cache	TSV file, loaded into memory	Improve persistence and incremental updates
Hit identification	Reads matching XPO files again	Reuse search results and avoid unnecessary reads
Model building	PowerShell loops and dictionaries	Reduce allocations and repeated lookups
Caller expansion	Reverse indexes	Already reasonably optimized
Called expansion	Iterates through object methods	Add reverse method indexes
Progress reporting	Frequent Write-Progress	Reduce UI overhead

The biggest opportunity is the initial keyword search and the repeated work surrounding it.
2. Use ripgrep for keyword searching










ripgrep (rg)
My first recommendation
Written in Rust, optimized for recursive searching, parallel directory traversal, and fast literal matching.

Official documentation: ripgrep documentation


VS Code uses ripgrep for its workspace text search. We can adopt the same underlying search utility.
What I would implement
First, detect whether rg.exe exists:
$script:AxRefRg = (Get-Command rg -ErrorAction SilentlyContinue).Source

if ($script:AxRefRg) {
    Write-Host "Using ripgrep: $script:AxRefRg"
}
else {
    Write-Host "ripgrep not found. Using PowerShell fallback."
}


Then use ripgrep to identify files containing any of the requested keywords.

Conceptually:
rg -i -F -l `
   -g "*.xpo" `
   -e "SalesLineGrouping" `
   -e "SalesTable" `
   "C:\AX\top"


Important options:
Option	Purpose
-i	Case-insensitive
-F	Literal matching, not regex
-l	Return filenames only
-g "*.xpo"	Only XPO files
-e	Specify search terms

I'd also use --hidden --no-ignore to preserve the current script's behavior of searching all recursively enumerated XPO files, rather than silently skipping ignored files.
How it fits your existing code
TOP + CUS XPO directories

Complete source corpus


ripgrep

If installed


PowerShell

Fallback




Matching XPO filenames

Candidate objects only


PowerShell XPO parser

Identify matching methods, declarations and properties


Existing relationship engine

Caller / called expansion, TOP/CUS merging, output generation


One important distinction: ripgrep identifies matching files, not necessarily matching AX methods. We should retain your existing XPO parser to identify the exact method containing the keyword.

Also, I would initially retain the current PowerShell fallback so the script remains usable on machines without additional dependencies.

Documentation: Installation and quick start. 
Tutorial & Quick Start Guide+1

3. The biggest optimization: fix the current indexing sequence
I found something particularly important in your existing AxRef-BuildIndex().

Currently, the sequence is:
1. Enumerate all XPO files.
2. Read the entire contents of each file.
3. Search its contents for keywords.
4. Check whether the file has changed.
5. If changed, parse it.
6. Otherwise, reuse the cache.
Notice the problem?

You read every file before checking whether it needs to be read.

That defeats a substantial part of your incremental cache.

I'd change this to:
Proposed indexing sequence
1. Enumerate files and collect metadata
   
   Full path, relative path, modification timestamp and file size.
2. Compare metadata against cache
   
   Identify new, modified, unchanged and deleted files.
3. Parse only new and modified files
   
   Read their contents once and update the cached object/method relationships.
4. Search keywords using ripgrep
   
   Search the corpus in parallel, returning matching filenames only.
5. Process matching XPO files
   
   Identify exact matching methods and non-method matches.
6. Build the relationship model and export
   
   Continue with your existing logic.

This means unchanged files won't need to be loaded into PowerShell memory just to rebuild the index.

There is still one unavoidable operation: ripgrep needs to inspect the source files to find new keyword matches. But its native parallel search should do that much more efficiently.

For even better performance, we could eventually maintain a persistent keyword-independent content index, but I wouldn't start there.
4. Further optimizations I'd make
These are the improvements I'd prioritize after introducing ripgrep.
A. Optimize XPO parsing
Your parser is doing several operations per object:
- Extracting SOURCE blocks.
- Removing XPO line prefixes.
- Stripping comments and string literals.
- Running multiple regexes to identify object references and method calls.
- Building PowerShell objects for every method.
There are two improvements worth making.

1. Optimize AxRef-GetSources()

Currently, for Forms and Reports, every SOURCE block iterates through the collection of all Name matches to determine its owner.

This can become expensive for large forms with many controls and methods.

Instead, maintain a moving pointer through the Name matches as SOURCE blocks are encountered. Both collections are already ordered by position, so we can avoid repeatedly scanning previous names.

2. Avoid unnecessary intermediate strings

AxRef-SourceToCode() splits and joins every source block. AxRef-Strip() then performs another regex replacement.

We can investigate consolidating these operations into fewer passes, while preserving comment and string-literal handling.

I wouldn't replace the parser with ripgrep. The parser understands XPO structure; ripgrep doesn't.
B. Optimize model building
Your AxRef-BuildModel() is reasonably well designed, particularly its reverse indexes.

However, I see three opportunities:
Improvement	Benefit
Use case-insensitive dictionaries consistently	Correct AX name resolution and fewer unnecessary conversions
Precompute method lookup indexes	Faster called-method resolution
Reduce repeated PowerShell object creation	Lower CPU and memory consumption

One specific opportunity is in AxRef-ExpandCalled().

You currently iterate through all methods of a referenced object to find methods matching the calls made by the current method.

We could introduce:
MethodsByName
    ObjectName
        MethodName -> Method


Then resolving called methods becomes a dictionary lookup instead of scanning every method.

Also, your ByName dictionary currently uses StringComparer.Ordinal, which is case-sensitive. AX object and method identifiers are case-insensitive, so I'd correct this consistently.
C. Make the cache more efficient
Your TSV cache is a reasonable lightweight solution. I wouldn't immediately replace it.

I'd improve it in stages:
1. Keep separate metadata for TOP and CUS.
2. Avoid loading unchanged XPO contents.
3. Update only changed entries.
4. Remove deleted objects from the cache.
5. Write cache updates atomically using a temporary file.
6. Add optional content-hash verification for situations where modification timestamps cannot be trusted.
One limitation of your existing timestamp-plus-file-size check: a file can change while retaining both values. A hash can catch that, although calculating hashes for every file on every run would undermine the performance improvement.
D. Improve progress reporting
AxRef-FindHits() currently calls Write-Progress for every entry.

I'd throttle this to approximately every 250–500 milliseconds, like your indexing function.

This is a small optimization, but it also makes the interface feel more responsive.
E. Consider parallel parsing
After the ripgrep improvement, the next candidate is parallel parsing of new and modified XPO files.

I would use PowerShell 7 runspaces or a runspace pool rather than Start-Job.

Why?
- Lower overhead than separate PowerShell processes.
- Multiple XPO files can be parsed simultaneously.
- We can control concurrency.
- Each worker can return a parsed object for merging into the main index.
PowerShell 7 supports ForEach-Object -Parallel, although Microsoft's documentation correctly cautions that parallelism has overhead and needs benchmarking. 
GitHub+1


I'd make this optional and benchmark a few concurrency levels rather than assuming more threads are always faster.
5. What I recommend installing
You don't need many additional tools.
1. ripgrep
Recommended

Fast file searching. This is the most important installation.
winget install BurntSushi.ripgrep.MSVC


Verify:
rg --version


Free and open source. GitHub repository

2. PowerShell 7
Recommended

Better foundation for runspaces and parallel processing. Keep compatibility with Windows PowerShell 5.1 if you still need it.
winget install Microsoft.PowerShell


PowerShell documentation

3. SQLite
Optional, later
Useful if we eventually move to a persistent searchable database containing objects, methods, references and metadata. Not necessary for the first optimization.

SQLite official website

I wouldn't install Elasticsearch, OpenSearch, Redis, or a dedicated search server for this tool. That's unnecessary infrastructure for an XPO reference extractor.

One thing to validate before adopting ripgrep: the actual encoding of your XPO files. Your script already handles different project-XPO encodings, but we should test the source corpus too. We must ensure ripgrep doesn't silently skip files it interprets as binary.

6. My proposed improvement roadmap
I'd implement these in stages so we can measure actual improvements rather than introducing unnecessary complexity.
Priority	Improvement	Expected impact
1	Detect and use ripgrep	Very high
2	Check cache before reading files	Very high
3	Reuse matching file information	High
4	Optimize SOURCE owner detection	Medium–high
5	Add method lookup indexes	Medium
6	Optimize model construction	Medium
7	Parallel parsing with PS7	Potentially high
8	SQLite persistent index	Future option

A further idea: separate indexing from extraction
I would eventually introduce two distinct operations:
- Index: Build or update the AX reference index. This is done once and incrementally thereafter.
- Extract: Search keywords, resolve callers/called objects, and generate the reference package.
This would make the tool more useful beyond its current purpose.

For example, eventually you could ask it to:
- Find all objects referencing a particular table.
- Find all callers of a particular method.
- Identify all forms using a particular EDT.
- Find where a particular enum value is used.
- Trace dependencies from a specific AX class.
- Compare TOP and CUS customizations.
The index would become a reusable AX code intelligence database rather than just an intermediate step for keyword searches.
7. What I would build first
My preferred target architecture:
AX Reference Extractor v2
Search engine

ripgrep with automatic PowerShell fallback

Incremental indexing engine

Metadata comparison, changed-file parsing, persistent cache

XPO analysis engine

Optimized parsing, method extraction, object/reference identification

Relationship engine

Reverse indexes, caller/called traversal, TOP/CUS resolution

Output engine

Sanitized XPP files, MANIFEST.md, AGENTS.md, importable project XPO


My recommendation: Don't rewrite this in C#, Rust, or Python yet. Your current PowerShell implementation already has substantial AX-specific logic that works. Optimize its bottlenecks first.

I'd also add performance instrumentation to each stage:
- File enumeration time
- Ripgrep search time
- Files read
- Files parsed
- Cache load/save time
- Model construction time
- Relationship expansion time
- Total execution time
That will tell us exactly where the next bottleneck is.

The next thing I'd do is produce an optimized v2 of your actual PowerShell script, preserving its existing commands, output structure, sanitization, project XPO generation, and relationship behavior, while implementing the first six improvements. Then we can benchmark the original against v2 using your actual AX corpus.

I would keep the first version compatible with both PowerShell 5.1 and 7, with parallel parsing enabled only under PowerShell 7.

## Implementation status

The following low-risk improvements are implemented in `AxReference_Optimized.ps1`:

- Ripgrep is used when available to shortlist literal, case-insensitive keyword matches; the PowerShell search remains the fallback.
- File metadata is checked before reading cached files. Only changed files are parsed, while matching files are read for exact method/property identification.
- Text read while parsing changed matches is reused by hit identification.
- SOURCE owner lookup advances through ordered NAME and SOURCE matches instead of rescanning prior names.
- Called-method lookup uses per-object method-name indexes, and object-name lookup is case-insensitive.
- Keyword-hit progress updates are throttled, and cache replacement is atomic.

Parallel parsing and a database-backed index remain deferred because they add complexity and need measurements from a representative AX corpus before they are justified.