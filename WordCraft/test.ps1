param(
    [Parameter(Mandatory=$true)][string]$WordCraft,
    [string]$WordCraftCli
)
$ErrorActionPreference = 'Stop'
if (-not $WordCraftCli) { $WordCraftCli = Join-Path (Split-Path $WordCraft) 'wordcraft-cli.exe' }
$launcher = Join-Path $PSScriptRoot 'openmd.exe'
$testRoot = Join-Path (Split-Path $PSScriptRoot) ('.tmp\openmd-tests-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
$script:checks = 0
function Assert($value, $message) {
    if (-not $value) { throw "FAIL: $message" }
    $script:checks++
}
function Run-Launcher([string[]]$Arguments, [int]$Expected = 0) {
    $p = Start-Process -FilePath $launcher -ArgumentList ($Arguments | ForEach-Object { '"' + $_ + '"' }) -Wait -PassThru
    Assert ($p.ExitCode -eq $Expected) "Launcher exit $($p.ExitCode), expected $Expected"
}
function Prepare($source) {
    $result = Join-Path $testRoot ([guid]::NewGuid().ToString('N') + '.json')
    Run-Launcher @($source, '--prepare', '--headless', '--wordcraft', $WordCraft, '--result', $result)
    return (Get-Content -Raw -LiteralPath $result | ConvertFrom-Json)
}
function Edit($prepared, $commands) {
    $scriptPath = Join-Path $testRoot 'edit.json'
    [IO.File]::WriteAllText($scriptPath, (@{ file=$prepared.docx; commands=@($commands); save=@($prepared.docx) } | ConvertTo-Json -Depth 20))
    & $WordCraftCli run --script $scriptPath
    Assert ($LASTEXITCODE -eq 0) 'WordCraft edit and DOCX save'
}
function Sync($prepared) { Run-Launcher @('--sync', $prepared.session, '--headless') }
function Invoke-Control($session, $profile, $method, $parameters) {
    $key = [IO.File]::ReadAllText((Join-Path $profile "control-key.local-$($session.Port)")).Trim()
    $client = [Net.Sockets.TcpClient]::new('127.0.0.1', $session.Port)
    try {
        $stream=$client.GetStream(); $stream.ReadTimeout=8000
        $writer=[IO.StreamWriter]::new($stream,[Text.UTF8Encoding]::new($false)); $writer.AutoFlush=$true
        $reader=[IO.StreamReader]::new($stream)
        $writer.WriteLine((@{id=1;key=$key;method=$method;params=$parameters}|ConvertTo-Json -Depth 15 -Compress))
        $response=$reader.ReadLine()|ConvertFrom-Json
        if(-not $response.ok){throw $response.error}
        return $response.result
    } finally { $client.Dispose() }
}
Add-Type -AssemblyName System.Drawing
$docs=Join-Path $testRoot 'Dokumenty ze spacją i ą'
$assets=Join-Path $docs 'grafiki ze spacją'
New-Item -ItemType Directory -Path $assets -Force | Out-Null
foreach($item in @(@('old',[Drawing.Color]::Red),@('new',[Drawing.Color]::Blue))) {
    $bmp=[Drawing.Bitmap]::new(8,8); $g=[Drawing.Graphics]::FromImage($bmp)
    try { $g.Clear($item[1]); $bmp.Save((Join-Path $assets ($item[0]+'.png')),[Drawing.Imaging.ImageFormat]::Png) }
    finally { $g.Dispose(); $bmp.Dispose() }
}
$front="---`r`ntitle: 'Dosłowny: YAML'`r`ntags: [a, b]`r`ncustom: {x: 1}`r`n---`r`n"
$md=$front+@'

<!-- openmd: Komentarze HTML przeniesiono na początek dokumentu z ich pierwotnych miejsc w treści z powodu konwersji MD -> DOCX -> MD. -->
<!-- pierwszy -->

# Próba

| A | B |
|---|---|
| 1 | 2 |

==MARK== oraz [Uwaga.]{.comment-start id="7" author="Tester" date="2026-11-17T10:00:00Z"}Fragment.[]{.comment-end id="7"}

![old](<grafiki ze spacją/old.png>)

<!-- drugi -->

```html
<!-- kod -->
```

Ostatni akapit.
'@
$source=Join-Path $docs 'próba.md'
$originalImageHash=(Get-FileHash -LiteralPath (Join-Path $assets 'old.png')).Hash
[IO.File]::WriteAllText($source,$md,[Text.UTF8Encoding]::new($true))
$prepared=Prepare $source
Assert ([IO.Path]::GetFileName($prepared.docx) -ceq 'próba.docx') 'DOCX uses original document name'
$before=[Convert]::ToBase64String([IO.File]::ReadAllBytes($source))
Sync $prepared
Assert ($before -ceq [Convert]::ToBase64String([IO.File]::ReadAllBytes($source))) 'No save means no Markdown rewrite'
Edit $prepared @(
    @{command='caret.docEnd';params=@{}},
    @{command='text.newParagraph';params=@{}},
    @{command='para.style';params=@{style='Normal'}},
    @{command='insert.picture';params=@{path=(Join-Path $assets 'new.png');alt='NEW'}}
)
Sync $prepared
$saved=Get-Content -Raw -LiteralPath $source
Assert ($saved.StartsWith($front)) 'YAML preserved exactly, including CRLF'
Assert ([IO.File]::ReadAllBytes($source)[0] -eq 239) 'UTF-8 BOM preserved'
Assert ($saved.IndexOf('<!-- pierwszy -->') -lt $saved.IndexOf('# Próba')) 'HTML comments moved to top'
Assert ($saved.IndexOf('<!-- drugi -->') -lt $saved.IndexOf('# Próba')) 'Trailing HTML comment moved to top'
Assert ($saved.Contains('<!-- kod -->')) 'Comment inside code retained as code'
$state=Get-Content -Raw -LiteralPath $prepared.session|ConvertFrom-Json
Assert ($state.Comments.Count -eq 2) 'Code example not harvested as a comment'
Assert $state.UsePipeTables 'Input pipe table detected'
Assert ($saved -match '(?m)^\|.*A.*\|.*B.*\|') 'Output keeps pipe table syntax'
Assert ($saved.Contains('HTML comments were moved')) 'Comment relocation explained'
Assert ($saved.Contains('==MARK==')) 'Mark retained'
Assert ($saved.Contains('.comment-start') -and $saved.Contains('author="Tester"') -and $saved.Contains('.comment-end')) 'Word comment and anchor retained'
Assert ($saved.Contains('old.png')) 'Original image path retained'
Assert (@(Get-ChildItem -LiteralPath $assets -Filter 'image_*.png').Count -eq 1) 'New image extracted into original relative directory'
$newImage=Get-ChildItem -LiteralPath $assets -Filter 'image_*.png' | Select-Object -First 1
Assert ((Get-FileHash -LiteralPath $newImage.FullName).Hash -eq (Get-FileHash -LiteralPath (Join-Path $assets 'new.png')).Hash) 'New image bytes preserved'
Edit $prepared @(@{command='select.objects';params=@{index=0}},@{command='text.delete';params=@{}})
Sync $prepared
$saved=Get-Content -Raw -LiteralPath $source
Assert (-not $saved.Contains('old.png')) 'Deleted image removed from Markdown'
Assert (Test-Path -LiteralPath (Join-Path $assets 'old.png')) 'Deleted image file left on disk'
Assert ((Get-FileHash -LiteralPath (Join-Path $assets 'old.png')).Hash -eq $originalImageHash) 'Deleted image file content unchanged'
Assert ([regex]::Matches($saved,'<!-- openmd:').Count -eq 1) 'Relocation notice emitted once'
$reopened=Prepare $source
Assert ($reopened.docx -ceq $prepared.docx) 'Reopening uses the same DOCX path'
$reopenedState=Get-Content -Raw -LiteralPath $reopened.session|ConvertFrom-Json
Assert ($reopenedState.Comments.Count -eq 2) 'Reopening does not harvest generated notice as a user comment'
$lock=[IO.File]::Open($reopened.docx,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
try {
    $alternate=Prepare $source
    Assert ($alternate.docx -cne $reopened.docx) 'Locked DOCX selects an alternate directory'
    Assert ([IO.Path]::GetFileName($alternate.docx) -ceq 'próba.docx') 'Alternate DOCX retains document name'
} finally { $lock.Dispose() }

$fallback=Join-Path $docs 'fallback.md'
[IO.File]::WriteAllText($fallback,"# Fallback`n`nText.`n")
$p2=Prepare $fallback
$fallbackState=Get-Content -Raw -LiteralPath $p2.session|ConvertFrom-Json
Assert (-not $fallbackState.UsePipeTables) 'No pipe table means no forced pipe export'
Edit $p2 @(@{command='caret.docEnd';params=@{}},@{command='text.newParagraph';params=@{}},@{command='insert.picture';params=@{path=(Join-Path $assets 'new.png');alt='fallback'}})
Sync $p2
Assert (@(Get-ChildItem -LiteralPath (Join-Path $docs 'fallback_img') -Filter '*.png').Count -eq 1) 'Default relative image directory'

$conflict=Join-Path $docs 'conflict.md'
[IO.File]::WriteAllText($conflict,"# Original`n")
$p3=Prepare $conflict
Edit $p3 @(@{command='text.insert';params=@{text='Edited '}})
[IO.File]::WriteAllText($conflict,'EXTERNAL EDIT')
Run-Launcher @('--sync',$p3.session,'--headless') 1
Assert ((Get-Content -Raw -LiteralPath $conflict) -ceq 'EXTERNAL EDIT') 'External edit not overwritten'
$conflictSession=Get-Content -Raw -LiteralPath $p3.session|ConvertFrom-Json
Assert (@(Get-ChildItem -LiteralPath (Join-Path $conflictSession.Directory 'exports') -Recurse -Filter candidate.md).Count -gt 0) 'Conflict candidate retained for recovery'

$missing=Join-Path $docs 'missing.md'; [IO.File]::WriteAllText($missing,'![missing](absent.png)')
Run-Launcher @($missing,'--prepare','--headless','--wordcraft',$WordCraft) 1
Assert ((Get-Content -Raw -LiteralPath $missing) -ceq '![missing](absent.png)') 'Missing image stops conversion without changing source'

# Real launcher process: GUI save triggers Markdown export; WordCraft exit ends the launcher.
$live=Join-Path $docs 'live.md'; [IO.File]::WriteAllText($live,"# Live`n`nText.`n")
$result=Join-Path $testRoot 'live.json'; $profile=Join-Path $testRoot 'profile'
$oldProfile=$env:WORDCRAFT_PROFILE_DIR; $oldNoPrefs=$env:WORDCRAFT_NO_PREFS
$owned=$null; $session=$null
try {
    $env:WORDCRAFT_PROFILE_DIR=$profile; $env:WORDCRAFT_NO_PREFS='1'
    $arguments=@($live,'--headless','--wordcraft',$WordCraft,'--result',$result)
    $owned=Start-Process -FilePath $launcher -ArgumentList ($arguments|ForEach-Object {'"'+$_+'"'}) -PassThru -WindowStyle Hidden
    $deadline=[DateTime]::UtcNow.AddSeconds(35)
    do {
        Start-Sleep -Milliseconds 200
        if($owned.HasExited){throw 'Launcher exited before control startup'}
        if([DateTime]::UtcNow -gt $deadline){throw 'Live startup timeout'}
        if(Test-Path -LiteralPath $result){
            $entry=Get-Content -Raw -LiteralPath $result|ConvertFrom-Json
            $session=Get-Content -Raw -LiteralPath $entry.session|ConvertFrom-Json
        }
    }until($session -and $session.Port -and (Test-Path -LiteralPath (Join-Path $profile "control-key.local-$($session.Port)")))
    Invoke-Control $session $profile 'engine.execute' @{command='caret.docEnd';params=@{}}|Out-Null
    Invoke-Control $session $profile 'engine.execute' @{command='text.insert';params=@{text=' LIVE SAVED'}}|Out-Null
    Invoke-Control $session $profile 'engine.execute' @{command='file.save';params=@{}}|Out-Null
    $deadline=[DateTime]::UtcNow.AddSeconds(20)
    while(-not (Get-Content -Raw -LiteralPath $live).Contains('LIVE SAVED')){
        if([DateTime]::UtcNow -gt $deadline){throw 'Live synchronization timeout'}
        Start-Sleep -Milliseconds 250
    }
    Assert $true 'Live GUI save synchronized automatically'
    try { Invoke-Control $session $profile 'app.quit' @{}|Out-Null }
    catch {
        # Quit can close the socket before its reply is read. Never replay it.
        $appProcess=Get-Process -Id $session.Pid -ErrorAction SilentlyContinue
        if($appProcess -and -not $appProcess.WaitForExit(8000)){throw}
    }
    Assert ($owned.WaitForExit(15000)) 'Launcher ended after WordCraft exit'
    Assert ($owned.ExitCode -eq 0) 'Live launcher exit successful'
}finally{
    if($session -and $session.Pid){Get-Process -Id $session.Pid -ErrorAction SilentlyContinue|Stop-Process -ErrorAction SilentlyContinue}
    if($owned -and -not $owned.HasExited){$owned.Kill()}
    $env:WORDCRAFT_PROFILE_DIR=$oldProfile; $env:WORDCRAFT_NO_PREFS=$oldNoPrefs
}
Write-Host "PASS: $script:checks checks. Test files: $testRoot"
