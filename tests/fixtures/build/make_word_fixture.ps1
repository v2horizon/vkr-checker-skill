param(
    [Parameter(Mandatory = $true)][string]$Python,
    [string]$WorkDir = (Join-Path $env:TEMP "vkr-word-fixture"),
    [int]$TimeoutSec = 300
)
# Пересоздаёт tests/fixtures/build/vkr-word-saved.docx: сборка фикстурного проекта,
# открытие копии в Word 16, обновление полей и содержания, полная раскладка,
# сохранение в новый файл и копирование в фикстуру.
# Word работает в отдельном задании с тайм-аутом; при зависании завершается только
# процесс окна открытого документа (GetWindowThreadProcessId по ActiveWindow.Hwnd).
# Чужие процессы Word не трогаются.
$ErrorActionPreference = 'Stop'
$fixtureDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repo = Resolve-Path (Join-Path $fixtureDir "..\..\..")
$project = Join-Path $WorkDir "project"
if (Test-Path $WorkDir) { Remove-Item -Recurse -Force $WorkDir }
New-Item -ItemType Directory -Force $WorkDir | Out-Null
Copy-Item -Recurse (Join-Path $fixtureDir "project") $project
$env:PYTHONDONTWRITEBYTECODE = "1"
& $Python (Join-Path $repo "vkr-mpgu\scripts\build_vkr.py") $project | Out-Null
if ($LASTEXITCODE -ne 0) { throw "build_vkr.py завершился с кодом $LASTEXITCODE" }
$source = Join-Path $WorkDir "vkr-copy.docx"
Copy-Item (Join-Path $project "final\vkr.docx") $source
$saved = Join-Path $WorkDir ("vkr-word-saved-" + [guid]::NewGuid().ToString("N") + ".docx")
$pidFile = Join-Path $WorkDir "word-pid.txt"
$target = Join-Path $fixtureDir "vkr-word-saved.docx"

$job = Start-Job -ArgumentList $source, $saved, $pidFile -ScriptBlock {
    param($source, $saved, $pidFile)
    $ErrorActionPreference = 'Stop'
    $before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $word = New-Object -ComObject Word.Application
    $doc = $null
    $shared = $false
    try {
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($source, $false, $false, $false)
        # PID процесса окна документа (как в update_docx_fields.py): останавливается только он.
        if (-not ('VkrFixtureUser32' -as [type])) {
            Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class VkrFixtureUser32 { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
        }
        [uint32]$found = 0
        [void][VkrFixtureUser32]::GetWindowThreadProcessId([IntPtr][int64]$doc.ActiveWindow.Hwnd, [ref]$found)
        if ($before -contains [int]$found) { $shared = $true } else { Set-Content -Path $pidFile -Value ([int]$found) }
        $doc.Fields.Update() | Out-Null
        foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
        $doc.Repaginate()
        $pages = $doc.ComputeStatistics(2)
        foreach ($paragraph in $doc.Paragraphs) { $null = $paragraph.Range.Information(3) }
        $doc.SaveAs2($saved, 16)
        "pages=$pages"
    }
    finally {
        if ($doc -ne $null) { $doc.Close($false) | Out-Null }
        if (-not $shared) { $word.Quit() | Out-Null }
    }
}
if (-not (Wait-Job $job -Timeout $TimeoutSec)) {
    Stop-Job $job
    if (Test-Path $pidFile) {
        foreach ($id in ((Get-Content $pidFile) -split ",")) {
            if ($id -and (Get-Process -Id ([int]$id) -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -eq 'WINWORD' })) { Stop-Process -Id ([int]$id) -Force }
        }
    }
    throw "Word не завершил работу за $TimeoutSec с; экземпляр, запущенный скриптом, остановлен"
}
$output = Receive-Job $job
Remove-Job $job
if (-not (Test-Path $saved)) { throw "Word не сохранил файл: $output" }
Copy-Item -Force $saved $target
Write-Output "$output saved=$target"
