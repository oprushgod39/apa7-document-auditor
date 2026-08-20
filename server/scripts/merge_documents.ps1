param([Parameter(Mandatory=$true)][string]$ManifestPath)

$ErrorActionPreference = 'Stop'
$manifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$word = $null
$destination = $null
$wdCollapseEnd = 0
$wdAlignLeft = 0
$wdAlignCenter = 1
$wdPageBreak = 7
$wdFormatDocumentDefault = 16

function Add-TextParagraph {
  param($Document, [string]$Text, [bool]$Bold = $false, [double]$Size = 12, [int]$Alignment = 0)
  $range = $Document.Content
  $range.Collapse($wdCollapseEnd)
  $range.Text = $Text
  $range.Font.Name = 'Times New Roman'
  $range.Font.NameAscii = 'Times New Roman'
  $range.Font.NameOther = 'Times New Roman'
  $range.Font.Size = $Size
  $range.Font.Bold = $(if ($Bold) { 1 } else { 0 })
  $range.ParagraphFormat.Alignment = $Alignment
  $range.ParagraphFormat.SpaceBefore = 0
  $range.ParagraphFormat.SpaceAfter = 0
  $range.InsertParagraphAfter()
}

try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $destination = $word.Documents.Add()

  for ($i = 0; $i -lt $manifest.documents.Count; $i++) {
    $item = $manifest.documents[$i]
    Add-TextParagraph -Document $destination -Text ([string]$item.name) -Bold $true -Size 12 -Alignment $wdAlignCenter

    $source = $word.Documents.Open([string]$item.path, $false, $true, $false)
    try {
      $copyEnd = $source.Content.End - 1
      foreach ($paragraph in $source.Paragraphs) {
        $candidate = ([string]$paragraph.Range.Text).Replace("`r", '').Replace([char]7, ' ').Trim()
        if ($candidate -match '^(references|bibliography|works\s+cited)\s*:?[\s\.]*$') {
          $copyEnd = $paragraph.Range.Start
          break
        }
      }
      if ($copyEnd -gt 0) {
        $sourceRange = $source.Range(0, $copyEnd)
        $targetRange = $destination.Content
        $targetRange.Collapse($wdCollapseEnd)
        $targetRange.FormattedText = $sourceRange.FormattedText
      }
    }
    finally {
      $source.Close($false)
    }

    if ($i -lt ($manifest.documents.Count - 1)) {
      for ($blank = 0; $blank -lt 4; $blank++) {
        $r = $destination.Content
        $r.Collapse($wdCollapseEnd)
        $r.InsertParagraphAfter()
      }
    }
  }

  $requestedAppendixWords = [Math]::Max(0, [int]$manifest.appendixWords)
  if ($requestedAppendixWords -gt 0) {
    $endRange = $destination.Content
    $endRange.Collapse($wdCollapseEnd)
    $endRange.InsertBreak($wdPageBreak)
    Add-TextParagraph -Document $destination -Text 'IGNORE' -Bold $true -Size 20 -Alignment $wdAlignCenter

    foreach ($paragraph in $manifest.appendixParagraphs) {
      Add-TextParagraph -Document $destination -Text ([string]$paragraph.text) -Bold ([bool]$paragraph.bold) -Size 12 -Alignment $wdAlignLeft
    }
  }

  $destination.SaveAs2([string]$manifest.outputPath, $wdFormatDocumentDefault)
}
finally {
  if ($null -ne $destination) { $destination.Close($false) }
  if ($null -ne $word) { $word.Quit() }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
