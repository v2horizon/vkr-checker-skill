param(
  [string]$InDir,
  [string]$OutDir,
  [string]$PdfDir = "",
  [string]$Only = ""
)
# Фикстуры валидатора через Microsoft Word 16 (COM). Только свои копии документов.
# Каждый вариант: открыть копию -> изменить средствами Word -> обновить оглавление и поля -> SaveAs2.
$ErrorActionPreference = "Stop"
$log = New-Object System.Collections.Generic.List[string]
function L($s) { $line = ("{0} {1}" -f (Get-Date -Format "HH:mm:ss"), [string]$s); $log.Add($line); Add-Content -Path (Join-Path $OutDir "word_fixtures.live.log") -Value $line -Encoding UTF8 }
function Retry([scriptblock]$b) {
  for ($i = 0; $i -lt 10; $i++) {
    try { return & $b } catch {
      if ($_.Exception.Message -match "RPC_E_CALL_REJECTED|0x80010001|0x8001010A|busy") { Start-Sleep -Milliseconds 700 } else { throw }
    }
  }
  return & $b
}
function MmPt([double]$mm) { return $mm * 72.0 / 25.4 }

$word = $null
$missing = [System.Reflection.Missing]::Value

function OpenDoc([string]$path) {
  return Retry { $word.Documents.Open($path, $false, $true, $false) }
}

function FindParagraph($doc, [string]$prefix, [int]$minIndex = 1) {
  $n = $doc.Paragraphs.Count
  for ($i = $minIndex; $i -le $n; $i++) {
    $p = $doc.Paragraphs.Item($i)
    $t = $p.Range.Text
    if ($t.StartsWith($prefix)) {
      # строки оглавления содержат табуляцию с номером страницы — пропускаем
      if ($t.Contains("`t")) { continue }
      return $p
    }
  }
  return $null
}

function UpdateAndSave($doc, [string]$name) {
  L "  update: toc"
  if ($doc.TablesOfContents.Count -ge 1) { Retry { $doc.TablesOfContents.Item(1).Update() } | Out-Null }
  L "  update: fields"
  Retry { $doc.Fields.Update() } | Out-Null
  L "  update: footers"
  foreach ($s in $doc.Sections) {
    foreach ($idx in 1, 2, 3) {
      try { $s.Footers.Item($idx).Range.Fields.Update() | Out-Null } catch {}
    }
  }
  L "  update: repaginate"
  $doc.Repaginate()
  L "  update: save"
  [string]$out = [string](Join-Path $OutDir $name)
  Retry { $doc.SaveAs2($out, 16) } | Out-Null
  $pages = $doc.ComputeStatistics(2)
  L ("saved $name pages=$pages sections=" + $doc.Sections.Count)
  if ($PdfDir -ne "") { try { Retry { $doc.ExportAsFixedFormat([string](Join-Path $PdfDir ($name + ".pdf")), 17) } | Out-Null } catch { L ("pdf failed: " + $_.Exception.Message) } }
  return $out
}

function PageOf($doc, [string]$prefix) {
  $p = FindParagraph $doc $prefix
  if ($p -eq $null) { return "not found" }
  return ("adjusted=" + $p.Range.Information(1) + " physical=" + $p.Range.Information(3))
}

function BuildingBlock([int]$type, [string]$pattern) {
  $word.Templates.LoadBuildingBlocks()
  foreach ($t in $word.Templates) {
    if ($t.Name -notlike "Built-In Building Blocks*") { continue }
    $count = $t.BuildingBlockEntries.Count
    for ($i = 1; $i -le $count; $i++) {
      $e = $t.BuildingBlockEntries.Item($i)
      if ($e.Type.Index -eq $type -and $e.Name -match $pattern) { return $e }
    }
  }
  return $null
}

function Want([string]$name) { return ($Only -eq "" -or $name -match $Only) }

$cfg = Get-Content -Raw -Encoding UTF8 (Join-Path $PSScriptRoot "word_fixtures.json") | ConvertFrom-Json

try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  L ("Word " + $word.Version + " build " + $word.Build)

  [string]$genIn = [string](Join-Path $InDir "gen_project.docx")
  [string]$base = [string](Join-Path $OutDir "gen_project_word.docx")

  if (Want "gen_project_word") {
    $doc = OpenDoc $genIn
    UpdateAndSave $doc "gen_project_word.docx" | Out-Null
    L ("  annotation page " + (PageOf $doc $cfg.annotation) + "; intro " + (PageOf $doc $cfg.introduction))
    $doc.Close([ref]0)
  }

  # (б) номер страницы стандартным блоком галереи «Простой номер 2»
  if (Want "ok_pagenum_building_block") {
    try {
      $bb = BuildingBlock 12 $cfg.plain_number_pattern
      if ($bb -eq $null) { throw "building block Plain Number 2 not found" }
      $doc = OpenDoc $base
      $f = $doc.Sections.Item(1).Footers.Item(1)
      $f.Range.Delete() | Out-Null
      $null = $bb.Insert($f.Range, $true)
      L ("  inserted building block: " + $bb.Name + " / " + $bb.Category.Name)
      UpdateAndSave $doc "ok_pagenum_building_block.docx" | Out-Null
      L ("  footer text='" + $doc.Sections.Item(1).Footers.Item(1).Range.Text.Trim() + "' contentcontrols=" + $doc.Sections.Item(1).Footers.Item(1).Range.ContentControls.Count)
      $doc.Close([ref]0)
    } catch { L ("ERROR ok_pagenum_building_block: " + $_.Exception.Message) }
  }

  # (в) автособираемое оглавление (w:sdt) вместо поля TOC генератора
  if (Want "ok_toc_sdt") {
    try {
      $bb = BuildingBlock 14 $cfg.auto_toc_pattern
      if ($bb -eq $null) { throw "building block Automatic Table 1 not found" }
      $doc = OpenDoc $base
      $heading = FindParagraph $doc ([string]$cfg.contents)
      $doc.TablesOfContents.Item(1).Delete()
      $rng = $heading.Range
      $rng.Text = ""
      $rng = $heading.Range
      $rng.Collapse(1)
      $null = $bb.Insert($rng, $true)
      L ("  inserted " + $bb.Name + " tocs=" + $doc.TablesOfContents.Count + " cc=" + $doc.ContentControls.Count)
      UpdateAndSave $doc "ok_toc_sdt.docx" | Out-Null
      $doc.Close([ref]0)
    } catch { L ("ERROR ok_toc_sdt: " + $_.Exception.Message) }
  }

  # (г) титульный лист отдельной секцией без колонтитула
  if (Want "ok_title_section") {
    try {
      $doc = OpenDoc $base
      $ann = FindParagraph $doc ([string]$cfg.annotation)
      $prev = $ann.Previous(1)
      if ($prev.Range.Text -match [char]12) { $prev.Range.Delete() | Out-Null }
      $ann = FindParagraph $doc ([string]$cfg.annotation)
      $r = $ann.Range
      $r.Collapse(1)
      $r.InsertBreak(2)
      $s1 = $doc.Sections.Item(1)
      $s2 = $doc.Sections.Item(2)
      $s2.Footers.Item(1).LinkToPrevious = $false
      $s2.Footers.Item(2).LinkToPrevious = $false
      $s2.PageSetup.DifferentFirstPageHeaderFooter = $false
      $s2.Footers.Item(1).PageNumbers.RestartNumberingAtSection = $false
      $s1.PageSetup.DifferentFirstPageHeaderFooter = $false
      $s1.Footers.Item(1).Range.Delete() | Out-Null
      UpdateAndSave $doc "ok_title_section.docx" | Out-Null
      L ("  annotation page " + (PageOf $doc $cfg.annotation) + " s1 footer='" + $s1.Footers.Item(1).Range.Text.Trim() + "' s2 footer fields=" + $s2.Footers.Item(1).Range.Fields.Count)
      $doc.Close([ref]0)
    } catch { L ("ERROR ok_title_section: " + $_.Exception.Message) }
  }

  # (д) стили заголовков синие (тема), заголовки перекрашены студентом в «Авто»; и испорченная копия без перекраски
  foreach ($variant in @("ok_headings_recolored_auto", "bad_headings_blue")) {
    if (-not (Want $variant)) { continue }
    try {
      $doc = OpenDoc $base
      foreach ($sid in -2, -3) { $doc.Styles.Item($sid).Font.TextColor.ObjectThemeColor = 4 }
      $n = $doc.Paragraphs.Count
      $changed = 0
      for ($i = 1; $i -le $n; $i++) {
        $p = $doc.Paragraphs.Item($i)
        if ($p.OutlineLevel -le 2) {
          if ($variant -eq "ok_headings_recolored_auto") { $p.Range.Font.Color = -16777216 } else { $p.Range.Font.Reset() }
          $changed++
        }
      }
      L ("  $variant headings changed=$changed")
      UpdateAndSave $doc ($variant + ".docx") | Out-Null
      $doc.Close([ref]0)
    } catch { L ("ERROR ${variant}: " + $_.Exception.Message) }
  }

  # испорченные копии эталона средствами Word
  $simple = [ordered]@{
    "bad_margin_left30" = { param($d) $d.Sections.Item(1).PageSetup.LeftMargin = [double](MmPt 30) }
    "bad_normal_font_arial" = { param($d) $d.Styles.Item(-1).Font.Name = "Arial" }
    "bad_normal_size12" = { param($d) $d.Styles.Item(-1).Font.Size = 12 }
    "bad_normal_spacing_single" = { param($d) $d.Styles.Item(-1).ParagraphFormat.LineSpacingRule = 0 }
    "bad_normal_indent0" = { param($d) $d.Styles.Item(-1).ParagraphFormat.FirstLineIndent = 0 }
    "bad_body_align_left" = { param($d)
        $cnt = 0
        foreach ($p in $d.Paragraphs) { if ($p.OutlineLevel -eq 10 -and $p.Range.Text.Length -gt 100 -and $p.Range.Font.Name -ne "Consolas" -and $p.Alignment -eq 3) { $p.Alignment = 0; $cnt++ } }
        L ("  left-aligned paragraphs: $cnt") }
    "bad_no_page_field" = { param($d) $d.Sections.Item(1).Footers.Item(1).Range.Delete() | Out-Null }
    "bad_title_page_number" = { param($d)
        $f = $d.Sections.Item(1).Footers.Item(2)
        $null = $f.Range.Fields.Add($f.Range, 33)
        L ("  first-page footer fields=" + $f.Range.Fields.Count) }
    "bad_letter" = { param($d) foreach ($s in $d.Sections) { $s.PageSetup.PaperSize = 2 } }
    "bad_marker_in_header" = { param($d) $d.Sections.Item(1).Headers.Item(1).Range.Text = [string]$cfg.header_marker }
    "bad_second_section_margins" = { param($d)
        $c = FindParagraph $d ([string]$cfg.conclusion)
        $prev = $c.Previous(1)
        if ($prev.Range.Text -match [char]12) { $prev.Range.Delete() | Out-Null }
        $c = FindParagraph $d ([string]$cfg.conclusion)
        $r = $c.Range; $r.Collapse(1); $r.InsertBreak(2)
        $s2 = $d.Sections.Item(2)
        $s2.PageSetup.DifferentFirstPageHeaderFooter = $false
        $s2.Footers.Item(1).PageNumbers.RestartNumberingAtSection = $false
        $s2.PageSetup.LeftMargin = [double](MmPt 15)
        $s2.PageSetup.RightMargin = [double](MmPt 25) }
  }
  foreach ($name in $simple.Keys) {
    if (-not (Want $name)) { continue }
    try {
      $doc = OpenDoc $base
      & $simple[$name] $doc
      UpdateAndSave $doc ($name + ".docx") | Out-Null
      $doc.Close([ref]0)
    } catch { L ("ERROR ${name}: " + $_.Exception.Message) }
  }

  # замены текста (Find/Replace Word) — до обновления оглавления
  foreach ($item in $cfg.replacements) {
    if (-not (Want $item.name)) { continue }
    try {
      $doc = OpenDoc $base
      $rng = $doc.Content
      $ok = $rng.Find.Execute([string]$item.find, $true, $false, $false, $false, $false, $true, 0, $false, [string]$item.replace, 2)
      L ("  $($item.name) replace ok=$ok")
      if ($item.normal_bold) {
        $p = FindParagraph $doc ([string]$item.replace)
        if ($p -ne $null) { $p.Style = -1; $p.Range.Font.Bold = $true; $p.Alignment = 1; $p.FirstLineIndent = 0; L "  style->Normal bold" }
      }
      UpdateAndSave $doc ($item.name + ".docx") | Out-Null
      $doc.Close([ref]0)
    } catch { L ("ERROR $($item.name): " + $_.Exception.Message) }
  }

  foreach ($pair in @(@("gen_project_unsorted_bib.docx", "bad_bibliography_unsorted.docx"), @("gen_project_skiplast.docx", "bad_trailing_empty_page.docx"))) {
    $nm = [System.IO.Path]::GetFileNameWithoutExtension($pair[1])
    if (-not (Want $nm)) { continue }
    try {
      $doc = OpenDoc (Join-Path $InDir $pair[0])
      UpdateAndSave $doc $pair[1] | Out-Null
      $doc.Close([ref]0)
    } catch { L ("ERROR $($pair[1]): " + $_.Exception.Message) }
  }

  # стили SPEC 8 и спорные правильные конструкции
  if (Want "ok_spec8_styles") {
    try {
      $doc = OpenDoc (Join-Path $InDir "kitchen.docx")
      $bb = BuildingBlock 12 $cfg.plain_number_pattern
      $f = $doc.Sections.Item(1).Footers.Item(1)
      $f.Range.Delete() | Out-Null
      $null = $bb.Insert($f.Range, $true)
      $rng = $doc.Content
      if ($rng.Find.Execute([string]"FOOTNOTE_HERE")) {
        $rng.Text = ""
        $null = $doc.Footnotes.Add($rng, $missing, [string]$cfg.footnote_text)
      }
      $app2 = FindParagraph $doc ([string]$cfg.appendix2)
      $app2.PageBreakBefore = $false
      $r = $app2.Range; $r.Collapse(1); $r.InsertBreak(2)
      $last = $doc.Sections.Item($doc.Sections.Count)
      $last.PageSetup.DifferentFirstPageHeaderFooter = $false
      $last.PageSetup.Orientation = 1
      $ps = $last.PageSetup
      L ("  landscape W={0:N1} H={1:N1} T={2:N1} B={3:N1} L={4:N1} R={5:N1}" -f ($ps.PageWidth/2.8346), ($ps.PageHeight/2.8346), ($ps.TopMargin/2.8346), ($ps.BottomMargin/2.8346), ($ps.LeftMargin/2.8346), ($ps.RightMargin/2.8346))
      UpdateAndSave $doc "ok_spec8_styles.docx" | Out-Null
      L ("  footnotes=" + $doc.Footnotes.Count + " intro page " + (PageOf $doc $cfg.introduction) + " app2 page " + (PageOf $doc $cfg.appendix2))
      $doc.Close([ref]0)
    } catch { L ("ERROR ok_spec8_styles: " + $_.Exception.Message) }
  }

  # XML-варианты (docDefaults), подготовленные Python, пересохраняются Word
  foreach ($file in (Get-ChildItem -Path $InDir -Filter "xml_*.docx")) {
    $target = $file.Name.Substring(4)
    if (-not (Want ([System.IO.Path]::GetFileNameWithoutExtension($target)))) { continue }
    try {
      $doc = OpenDoc $file.FullName
      $st = $doc.Styles.Item(-1)
      L ("  $target Normal: font=" + $st.Font.Name + " size=" + $st.Font.Size + " rule=" + $st.ParagraphFormat.LineSpacingRule + " ls=" + $st.ParagraphFormat.LineSpacing)
      UpdateAndSave $doc $target | Out-Null
      $doc.Close([ref]0)
    } catch { L ("ERROR ${target}: " + $_.Exception.Message) }
  }
} catch {
  L ("FATAL: " + $_.Exception.Message)
} finally {
  if ($word -ne $null) {
    try { foreach ($d in @($word.Documents)) { $d.Close([ref]0) } } catch {}
    try { $word.Quit([ref]0) } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
  }
  [System.IO.File]::AppendAllLines((Join-Path $OutDir "word_fixtures.log"), $log, (New-Object System.Text.UTF8Encoding($false)))
}
