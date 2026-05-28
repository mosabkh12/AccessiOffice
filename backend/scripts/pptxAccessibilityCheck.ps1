<#
.SYNOPSIS
  Opens a PPTX file in PowerPoint (visible), triggers the built-in Accessibility
  Checker pane via ExecuteMso, then reads the pane content through Windows
  UIAutomation. Outputs one JSON object to stdout. Debug messages go to stderr.

.PARAMETER FilePath
  Absolute path to the .pptx file to scan.

.NOTES
  Env vars read at runtime:
    PPTX_WORKER_DEBUG=true  - include rawOfficeText + verbose stderr logs
  Requires: Windows, Microsoft PowerPoint installed (desktop edition).
  The presentation is opened read-only with a visible window so ribbon commands
  and UIAutomation work. PowerPoint window will briefly appear on screen.
#>
param([Parameter(Mandatory=$true)][string]$FilePath)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$DEBUG   = ($env:PPTX_WORKER_DEBUG -eq 'true')
$OPEN_MS = 2000
$POLL_MS = 800
$MAX_MS  = 14000

# -- Load UIAutomation --------------------------------------------------------
try {
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
} catch {
    [ordered]@{ ok = $false; error = "UIAutomationClient not available: $($_.Exception.Message)" } |
        ConvertTo-Json -Compress
    exit 0
}

# -- Script-level text store (avoids PS pipeline / return-type issues) --------
[System.Collections.Generic.List[string]]$script:uiaTexts =
    [System.Collections.Generic.List[string]]::new()

function Add-WindowTexts([System.Windows.Automation.AutomationElement]$Root) {
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $queue  = [System.Collections.Generic.Queue[System.Windows.Automation.AutomationElement]]::new()
    [void]$queue.Enqueue($Root)
    $n = 0
    while ($queue.Count -gt 0 -and $n -lt 8000) {
        $el = $queue.Dequeue(); $n++
        try {
            $nm = $el.Current.Name
            if ($nm -and $nm.Trim().Length -gt 0) { [void]$script:uiaTexts.Add($nm.Trim()) }
        } catch {}
        try {
            $child = $walker.GetFirstChild($el)
            while ($null -ne $child) {
                [void]$queue.Enqueue($child)
                try { $child = $walker.GetNextSibling($child) } catch { break }
            }
        } catch {}
    }
}

function Test-PaneLoaded {
    # Returns true when a known Accessibility Checker sentinel appears in collected text.
    # English only - Hebrew strings cannot be embedded safely in PS1 files without UTF-8 BOM.
    $sentinels = @(
        'Inspection Results',
        'Accessibility Checker',
        'Accessibility Assistant',
        'No accessibility issues',
        'No issues found',
        'Congratulations',
        'Check Complete',
        'Errors',
        'Warnings',
        'Tips',
        'Accessibility'
    )
    foreach ($t in $script:uiaTexts) {
        foreach ($s in $sentinels) {
            if ($t -match [regex]::Escape($s)) { return $true }
        }
    }
    return $false
}

function Parse-CheckerTexts {
    # Priority order for count extraction:
    #   Pass 1 (HIGHEST): Adjacent label + count/checkmark elements.
    #     The real Accessibility Assistant renders each check as two adjacent UIAutomation
    #     elements: a pure-label string (no digits) followed by a digit string or checkmark.
    #     This format represents the whole-document totals and is always preferred.
    #   Pass 2 (FALLBACK): "Label (N)" parens, then "Label - N" dash.
    #     These appear in summary/section rows and may reflect per-slide or partial counts.
    #     Used only when Pass 1 found no adjacent match for a given rule key.
    $counts   = [ordered]@{}
    $statuses = [ordered]@{}

    $ruleMap = [ordered]@{
        missingAltText       = 'Missing Alt(?:ernative)? [Tt]ext|Missing alternative text'
        missingSlideTitle    = 'Missing Slide Title|Slide title.*missing|Missing.*slide title'
        duplicateSlideTitle  = 'Duplicate Slide Title|Duplicate.*slide title'
        readingOrder         = '(?:Check )?Reading Order'
        hardToReadText       = 'Hard.to.Read(?:\s+\w+)*|Hard-to-Read|Hard to Read|Color Contrast|Insufficient.*[Cc]ontrast'
        mediaCaptions        = 'Missing.*(?:Audio|Video|Subtitles|Captions)|(?:Audio|Video).*[Cc]aptions?.*[Mm]issing|Missing audio or video subtitles'
        missingTableHeader   = 'Missing Table Header|Table.*Header Row|Missing table header'
        mergedCells          = 'Merged(?:.*)[Cc]ells|Split Cells|Use of merged or split cells'
        restrictedAccess     = 'Restricted Access|Restricted access'
    }

    $ignCase = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    $textArr = $script:uiaTexts.ToArray()
    $textLen = $textArr.Length

    foreach ($entry in $ruleMap.GetEnumerator()) {
        $key = $entry.Key; $pat = $entry.Value
        $found = $false

        # -- Pass 1: Adjacent checklist format (highest priority) -----------------
        # Scan for a digit-free text element matching the rule, immediately followed
        # by a pure-digit count or a checkmark (U+2714 / U+2713 = passed/0).
        for ($i = 0; $i -lt ($textLen - 1) -and -not $found; $i++) {
            $t = $textArr[$i]
            if ([regex]::IsMatch($t, '\d')) { continue }           # skip texts with digits
            if (-not [regex]::IsMatch($t, "(?:$pat)", $ignCase)) { continue }
            $next = $textArr[$i + 1]
            if ($next -match '^\d+$') {
                $n = [int]$next
                if ($n -gt 0) { $counts[$key] = $n }
                $found = $true
            } elseif ($next.Length -ge 1 -and $next.Length -le 2) {
                $cp = [int][char]$next[0]
                if ($cp -eq 0x2714 -or $cp -eq 0x2713) { $found = $true }  # checkmark = 0/passed
            }
        }

        # -- Pass 2: Parens or dash summary lines (fallback only) -----------------
        if (-not $found) {
            for ($i = 0; $i -lt $textLen -and -not $found; $i++) {
                $t = $textArr[$i]
                # "Label (N)" - parentheses format
                $m = [regex]::Match($t, "(?:$pat)\s*\((\d+)\)", $ignCase)
                if ($m.Success) {
                    $n = [int]$m.Groups[1].Value
                    if ($n -gt 0) { $counts[$key] = $n }
                    $found = $true; break
                }
                # "Label - N" - whole string ends with dash + number
                $m2 = [regex]::Match($t, '^(.*\S)\s+-\s+(\d+)\s*$')
                if ($m2.Success -and [regex]::IsMatch($m2.Groups[1].Value, "(?:$pat)", $ignCase)) {
                    $n = [int]$m2.Groups[2].Value
                    if ($n -gt 0) { $counts[$key] = $n }
                    $found = $true; break
                }
            }
        }
    }

    # Statuses: key absent from counts = PPT showed no issue = passed
    $checkToCount = [ordered]@{
        contrast           = 'hardToReadText'
        mediaCaptions      = 'mediaCaptions'
        missingTableHeader = 'missingTableHeader'
        mergedCells        = 'mergedCells'
    }
    foreach ($sk in $checkToCount.Keys) {
        $ck = $checkToCount[$sk]
        $statuses[$sk] = if ($counts.Contains($ck) -and $counts[$ck] -gt 0) { 'failed' } else { 'passed' }
    }

    return [PSCustomObject]@{ counts = $counts; statuses = $statuses }
}

# -- Validate input -----------------------------------------------------------
if (-not (Test-Path $FilePath)) {
    [ordered]@{ ok = $false; error = "File not found: $FilePath" } | ConvertTo-Json -Compress
    exit 0
}

$ppt = $null; $pres = $null

try {
    # -- Start PowerPoint visible so ExecuteMso and UIAutomation work ----------
    $ppt         = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

    $msoT = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $msoF = [Microsoft.Office.Core.MsoTriState]::msoFalse

    # Open read-only WITH a window (ribbon commands require a visible window)
    $pres        = $ppt.Presentations.Open($FilePath, $msoT, $msoF, $msoT)
    $pptVersion  = $ppt.Version
    $slideCount  = $pres.Slides.Count

    if ($DEBUG) { [Console]::Error.WriteLine("[PPT DEBUG] Opened: version=$pptVersion slides=$slideCount") }

    Start-Sleep -Milliseconds $OPEN_MS

    # -- Trigger the Accessibility Checker pane via ribbon command -------------
    $executedMso = $null
    $msoAttempts = @(
        'ReviewCheckAccessibility',
        'CheckAccessibility',
        'AccessibilityChecker',
        'ReviewAccessibilityChecker'
    )
    foreach ($id in $msoAttempts) {
        try {
            $ppt.CommandBars.ExecuteMso($id)
            $executedMso = $id
            if ($DEBUG) { [Console]::Error.WriteLine("[PPT DEBUG] ExecuteMso '$id' succeeded") }
            break
        } catch {
            if ($DEBUG) { [Console]::Error.WriteLine("[PPT DEBUG] ExecuteMso '$id' failed: $($_.Exception.Message)") }
        }
    }

    if ($null -eq $executedMso) {
        $tried = $msoAttempts -join ', '
        throw "Could not open Accessibility Checker pane. All ExecuteMso IDs failed: $tried. Ensure PowerPoint 2016+ is installed and the Review tab is accessible."
    }

    # -- Find all top-level PowerPoint windows for UIAutomation ---------------
    $procs = [System.Diagnostics.Process]::GetProcessesByName('POWERPNT')
    if ($procs.Count -eq 0) { throw 'POWERPNT.EXE process not found after Presentations.Open()' }
    $pptPid = ($procs | Sort-Object StartTime -Descending | Select-Object -First 1).Id

    if ($DEBUG) { [Console]::Error.WriteLine("[PPT DEBUG] PID=$pptPid") }

    $root    = [System.Windows.Automation.AutomationElement]::RootElement
    $pidCond = [System.Windows.Automation.PropertyCondition]::new(
                   [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $pptPid)

    # -- Poll UIAutomation until pane sentinel text is found ------------------
    $elapsed = 0; $loaded = $false

    while ($elapsed -lt $MAX_MS) {
        Start-Sleep -Milliseconds $POLL_MS
        $elapsed += $POLL_MS

        # Collect text from ALL windows belonging to the PPT process.
        # The Accessibility pane may appear as a sibling window in some Office versions.
        $script:uiaTexts.Clear()
        $wins = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $pidCond)
        if ($null -ne $wins) {
            foreach ($w in $wins) { Add-WindowTexts $w }
        }

        if ($DEBUG) {
            $sample = ($script:uiaTexts | Select-Object -First 8) -join ' | '
            [Console]::Error.WriteLine("[PPT DEBUG] Poll ${elapsed}ms: nodes=$($script:uiaTexts.Count) sample=[$sample]")
        }

        if (Test-PaneLoaded) { $loaded = $true; break }
    }

    # -- Build output ---------------------------------------------------------
    if (-not $loaded) {
        $errMsg = "Accessibility Checker pane sentinel text not found after ${MAX_MS}ms. Enable PPTX_WORKER_DEBUG=true to see raw UIAutomation text. Possible cause: pane hosted in WebView2 (Office 365 insider build) - UIAutomation cannot read WebView2 content."
        $errOut = [ordered]@{
            ok          = $false
            error       = $errMsg
            pptVersion  = $pptVersion
            executedMso = $executedMso
        }
        if ($DEBUG) {
            $cap = [Math]::Min(199, $script:uiaTexts.Count - 1)
            $errOut['rawOfficeText'] = if ($script:uiaTexts.Count -gt 0) { $script:uiaTexts.ToArray()[0..$cap] } else { @() }
        }
        $errOut | ConvertTo-Json -Depth 3 -Compress

    } else {
        $parsed = Parse-CheckerTexts
        $cap    = [Math]::Min(199, $script:uiaTexts.Count - 1)
        $raw    = if ($script:uiaTexts.Count -gt 0) { $script:uiaTexts.ToArray()[0..$cap] } else { @() }

        $out = [ordered]@{
            ok            = $true
            engine        = 'powerpoint-ui-automation'
            pptVersion    = $pptVersion
            slideCount    = $slideCount
            executedMso   = $executedMso
            counts        = $parsed.counts
            statuses      = $parsed.statuses
            rawOfficeText = $raw
        }
        $out | ConvertTo-Json -Depth 5 -Compress
    }

} catch {
    [ordered]@{
        ok    = $false
        error = "PowerPoint Accessibility Assistant UI could not be read: $($_.Exception.Message)"
    } | ConvertTo-Json -Compress
} finally {
    if ($null -ne $pres) {
        try { $pres.Close() }  catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null } catch {}
        $pres = $null
    }
    if ($null -ne $ppt) {
        try { $ppt.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch {}
        $ppt = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
