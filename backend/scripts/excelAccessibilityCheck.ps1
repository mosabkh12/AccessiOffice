<#
.SYNOPSIS
  Opens an XLSX file in a NEW Excel instance (visible), triggers the Accessibility
  Checker pane via ExecuteMso, then reads the pane content through UIAutomation.
  Outputs one JSON object to stdout. Debug messages go to stderr.

.PARAMETER FilePath
  Absolute path to the .xlsx file to scan.

.NOTES
  Env vars read at runtime:
    EXCEL_WORKER_DEBUG=true  - verbose stderr logs + extra debug fields in JSON output
  Requires: Windows, Microsoft Excel installed (desktop edition, 2016+).

  Key isolation strategy:
  - [Activator]::CreateInstance forces a NEW EXCEL.EXE via CoCreateInstance,
    rather than reusing a running instance (GetActiveObject). This prevents
    reading a stale Accessibility pane from a prior scan session.
  - UIAutomation is filtered to that process's PID, obtained from the window
    handle (HWND) via GetWindowThreadProcessId.
  - Any pre-existing Accessibility pane is dismissed before re-triggering,
    ensuring the checker runs fresh for the current workbook.
  - The pane is not considered "ready" until result-category-level text
    appears (not just the pane title). If "Looks good" appears, one extra
    stability poll is done before accepting the result.
#>
param([Parameter(Mandatory=$true)][string]$FilePath)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$DEBUG   = ($env:EXCEL_WORKER_DEBUG -eq 'true')
$OPEN_MS = 2500   # ms after Workbooks.Open before any ExecuteMso
$POLL_MS = 900    # UIAutomation polling interval
$MAX_MS  = 16000  # max time to wait for result-level content in pane

# --------------------------------------------------------------------------
# Assembly loading
# --------------------------------------------------------------------------
try {
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
} catch {
    [ordered]@{ ok = $false; error = "UIAutomationClient not available: $($_.Exception.Message)" } |
        ConvertTo-Json -Compress
    exit 0
}

# Win32 helper: get process ID from window handle (more reliable than GetProcessesByName)
try {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class WinApiHelperXl {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
}
'@
} catch {}  # harmless if already defined in this session

# --------------------------------------------------------------------------
# UIAutomation text collection
# --------------------------------------------------------------------------
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

function Collect-UiaTexts([System.Windows.Automation.PropertyCondition]$PidCond) {
    $script:uiaTexts.Clear()
    $r = [System.Windows.Automation.AutomationElement]::RootElement
    $wins = $r.FindAll([System.Windows.Automation.TreeScope]::Children, $PidCond)
    if ($null -ne $wins) { foreach ($w in $wins) { Add-WindowTexts $w } }
}

# --------------------------------------------------------------------------
# Pane state detection
# --------------------------------------------------------------------------
function Test-StalePane {
    # True when a (possibly stale, from prior workbook) pane title is visible
    foreach ($t in $script:uiaTexts) {
        if ($t -match 'Accessibility Checker|Accessibility Assistant|Inspection Results') {
            return $true
        }
    }
    return $false
}

function Test-PaneReady {
    # True ONLY when result-category or issue-label level text is visible.
    # "Accessibility Checker" alone is NOT enough - that just means the pane title appeared.
    $readySentinels = @(
        'Color and Contrast',
        'Tables',
        'Document Structure',
        'Document Access',
        'Missing alt text',
        'Hard-to-read text',
        'Default sheet tab name',
        'Missing table header',
        'Merged or split cells',
        'Unclear hyperlink',
        'Looks good! No issues found.',
        'No issues found'
    )
    foreach ($t in $script:uiaTexts) {
        foreach ($s in $readySentinels) {
            if ($t -match [regex]::Escape($s)) { return $true }
        }
    }
    return $false
}

# --------------------------------------------------------------------------
# Parser with optional trace
# --------------------------------------------------------------------------
function Parse-CheckerTexts([bool]$Trace = $false) {
    $counts     = [ordered]@{}
    $statuses   = [ordered]@{}
    $parseTrace = [System.Collections.Generic.List[object]]::new()

    # Keys are the final officeLikeSummary key names
    $ruleMap = [ordered]@{
        missingAltText       = 'Missing Alt(?:ernative)? [Tt]ext|Missing alternative text|Alt Text.*Missing|Missing.*alt text'
        contrast             = 'Hard.to.Read(?:\s+\w+)*|Hard-to-Read|Hard to Read|Color Contrast|Insufficient.*[Cc]ontrast'
        tableHeader          = 'Missing [Tt]able [Hh]eader|Table [Hh]eader [Rr]ow [Mm]issing|Header [Rr]ow.*[Mm]issing'
        mergedCells          = 'Merged (?:or Split )?[Cc]ells|Split [Cc]ells|Use of [Mm]erged'
        unclearHyperlinkText = 'Unclear [Hh]yperlink [Tt]ext|Unclear [Ll]ink [Tt]ext|Unclear [Hh]yperlink'
        sheetName            = 'Default [Ss]heet [Tt]ab [Nn]ame|[Ii]naccessible [Ss]heet [Nn]ame|[Ss]heet [Tt]ab.*[Nn]ame'
        restrictedAccess     = 'Restricted [Aa]ccess|Document [Rr]estrictions?'
    }

    $ignCase = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    $textArr = $script:uiaTexts.ToArray()
    $textLen = $textArr.Length

    foreach ($entry in $ruleMap.GetEnumerator()) {
        $key = $entry.Key; $pat = $entry.Value
        $found = $false

        # -- Pass 1: Adjacent label + count/checkmark (highest priority) ----------
        # The Accessibility Assistant renders each check as a digit-free label
        # immediately followed by a pure-digit count or a checkmark ([OK]/[PASS]).
        # This represents whole-workbook totals and is always preferred.
        for ($i = 0; $i -lt ($textLen - 1) -and -not $found; $i++) {
            $t = $textArr[$i]
            if ([regex]::IsMatch($t, '\d')) { continue }    # skip lines with digits
            if (-not [regex]::IsMatch($t, "(?:$pat)", $ignCase)) { continue }
            $next = $textArr[$i + 1]
            if ($next -match '^\d+$') {
                $n = [int]$next
                if ($n -gt 0) { $counts[$key] = $n }
                $found = $true
                if ($Trace) {
                    $parseTrace.Add([ordered]@{
                        key = $key; pass = 1; sourceFormat = 'adjacent-label-count'
                        labelIndex = $i; labelText = $t; nextText = $next; result = $n; selected = $true
                    })
                }
            } elseif ($next.Length -ge 1 -and $next.Length -le 2) {
                $cp = [int][char]$next[0]
                if ($cp -eq 0x2714 -or $cp -eq 0x2713) {   # checkmark = passed
                    $found = $true
                    if ($Trace) {
                        $parseTrace.Add([ordered]@{
                            key = $key; pass = 1; sourceFormat = 'adjacent-label-checkmark'
                            labelIndex = $i; labelText = $t; nextText = $next; result = 0; selected = $true
                        })
                    }
                }
            }
        }

        # -- Pass 2: Parens / dash summary lines (fallback only) ------------------
        if (-not $found) {
            for ($i = 0; $i -lt $textLen -and -not $found; $i++) {
                $t = $textArr[$i]
                $m = [regex]::Match($t, "(?:$pat)\s*\((\d+)\)", $ignCase)
                if ($m.Success) {
                    $n = [int]$m.Groups[1].Value
                    if ($n -gt 0) { $counts[$key] = $n }
                    $found = $true
                    if ($Trace) {
                        $parseTrace.Add([ordered]@{
                            key = $key; pass = 2; sourceFormat = 'parens'
                            sourceIndex = $i; text = $t; result = $n; selected = $true
                        })
                    }
                    break
                }
                $m2 = [regex]::Match($t, '^(.*\S)\s+-\s+(\d+)\s*$')
                if ($m2.Success -and [regex]::IsMatch($m2.Groups[1].Value, "(?:$pat)", $ignCase)) {
                    $n = [int]$m2.Groups[2].Value
                    if ($n -gt 0) { $counts[$key] = $n }
                    $found = $true
                    if ($Trace) {
                        $parseTrace.Add([ordered]@{
                            key = $key; pass = 2; sourceFormat = 'dash'
                            sourceIndex = $i; text = $t; result = $n; selected = $true
                        })
                    }
                    break
                }
            }
        }
    }

    # Build statuses from counts
    $autoStatusKeys = @('missingAltText', 'contrast', 'tableHeader', 'mergedCells', 'unclearHyperlinkText', 'sheetName')
    foreach ($sk in $autoStatusKeys) {
        $statuses[$sk] = if ($counts.Contains($sk) -and $counts[$sk] -gt 0) { 'failed' } else { 'passed' }
    }
    $statuses['restrictedAccess'] = 'manual'   # always manual - not exposed through this pane path

    return [PSCustomObject]@{ counts = $counts; statuses = $statuses; parseTrace = $parseTrace }
}

# --------------------------------------------------------------------------
# Input validation
# --------------------------------------------------------------------------
if (-not (Test-Path $FilePath)) {
    [ordered]@{ ok = $false; error = "File not found: $FilePath" } | ConvertTo-Json -Compress
    exit 0
}

$absPath    = (Resolve-Path $FilePath).Path
$xl         = $null
$wb         = $null
$xlPid      = 0

try {
    # -- Force a NEW EXCEL.EXE instance ----------------------------------------
    # [Activator]::CreateInstance calls CoCreateInstance, which always launches
    # a new local-server EXE, unlike New-Object -ComObject which calls
    # GetActiveObject first and reuses any running Excel instance.
    # This is critical: reusing an existing instance can leave a stale
    # Accessibility pane open from a previous workbook.
    if ($DEBUG) { [Console]::Error.WriteLine("[EXCEL DEBUG] Creating new Excel.Application via Activator::CreateInstance") }
    $xlType = [Type]::GetTypeFromProgID('Excel.Application')
    if ($null -eq $xlType) { throw "Excel.Application ProgID not registered - is Microsoft Excel installed?" }
    $xl = [Activator]::CreateInstance($xlType)

    $xl.Visible       = $true
    $xl.DisplayAlerts = $false
    $xl.AutomationSecurity = 3   # msoAutomationSecurityForceDisable

    # -- Open workbook read-only -----------------------------------------------
    if ($DEBUG) { [Console]::Error.WriteLine("[EXCEL DEBUG] Opening: $absPath") }
    # Workbooks.Open(Filename, UpdateLinks, ReadOnly)
    $wb     = $xl.Workbooks.Open($absPath, 0, $true)
    $wbName = $wb.Name
    try { $xl.Windows(1).Activate() } catch {}   # bring our workbook window to front

    $xlVersion  = $xl.Version
    $sheetCount = $wb.Sheets.Count
    if ($DEBUG) { [Console]::Error.WriteLine("[EXCEL DEBUG] Opened: version=$xlVersion wbName=$wbName sheets=$sheetCount") }

    # -- Get PID from HWND (reliable even when multiple EXCEL instances exist) --
    try {
        $hwnd = [IntPtr]$xl.Hwnd
        [WinApiHelperXl]::GetWindowThreadProcessId($hwnd, [ref]$xlPid) | Out-Null
        if ($DEBUG) { [Console]::Error.WriteLine("[EXCEL DEBUG] PID from HWND: $xlPid") }
    } catch {
        # Fallback: pick the most recently started EXCEL process
        $procs = [System.Diagnostics.Process]::GetProcessesByName('EXCEL')
        if ($procs.Count -gt 0) {
            $xlPid = ($procs | Sort-Object StartTime -Descending | Select-Object -First 1).Id
        }
        if ($DEBUG) { [Console]::Error.WriteLine("[EXCEL DEBUG] PID via fallback: $xlPid (HWND method failed: $($_.Exception.Message))") }
    }

    if ($xlPid -eq 0) { throw "Could not determine Excel process ID" }

    $root    = [System.Windows.Automation.AutomationElement]::RootElement
    $pidCond = [System.Windows.Automation.PropertyCondition]::new(
                   [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $xlPid)

    Start-Sleep -Milliseconds $OPEN_MS

    # -- Dismiss any stale Accessibility pane ----------------------------------
    # Even with a new Excel instance, if Excel inherits task-pane state from a
    # previous session (e.g. roaming profile), the pane may already be open.
    # Toggle it off first so the fresh trigger runs against the current workbook.
    Collect-UiaTexts $pidCond
    if (Test-StalePane) {
        if ($DEBUG) { [Console]::Error.WriteLine("[EXCEL DEBUG] Stale accessibility pane detected - dismissing before fresh trigger") }
        foreach ($id in @('ReviewCheckAccessibility', 'CheckAccessibility', 'AccessibilityChecker')) {
            try { $xl.CommandBars.ExecuteMso($id); Start-Sleep -Milliseconds 600; break } catch {}
        }
    }

    # -- Trigger Accessibility Checker pane ------------------------------------
    $executedMso = $null
    $msoAttempts = @('ReviewCheckAccessibility', 'CheckAccessibility', 'AccessibilityChecker', 'ReviewAccessibilityChecker')
    foreach ($id in $msoAttempts) {
        try {
            $xl.CommandBars.ExecuteMso($id)
            $executedMso = $id
            if ($DEBUG) { [Console]::Error.WriteLine("[EXCEL DEBUG] ExecuteMso '$id' succeeded") }
            break
        } catch {
            if ($DEBUG) { [Console]::Error.WriteLine("[EXCEL DEBUG] ExecuteMso '$id' failed: $($_.Exception.Message)") }
        }
    }

    if ($null -eq $executedMso) {
        $tried = $msoAttempts -join ', '
        throw "Could not open Accessibility Checker pane. All ExecuteMso IDs failed: $tried. Ensure Excel 2016+ is installed and the Review tab is accessible."
    }

    # -- Poll until RESULT-LEVEL content appears --------------------------------
    # "Accessibility Checker" or "Accessibility Assistant" appearing is NOT enough
    # - we need actual result categories or issue labels to be visible.
    $elapsed   = 0
    $pollCount = 0
    $loaded    = $false

    while ($elapsed -lt $MAX_MS) {
        Start-Sleep -Milliseconds $POLL_MS
        $elapsed   += $POLL_MS
        $pollCount++

        Collect-UiaTexts $pidCond

        if ($DEBUG) {
            $sample = ($script:uiaTexts | Select-Object -First 10) -join ' | '
            [Console]::Error.WriteLine("[EXCEL DEBUG] Poll $pollCount (${elapsed}ms): nodes=$($script:uiaTexts.Count) sample=[$sample]")
        }

        if (Test-PaneReady) { $loaded = $true; break }
    }

    # -- Stability re-poll when "Looks good" appears ---------------------------
    # "Looks good! No issues found." may appear before Excel has finished
    # running the checker on the current workbook. Do one extra poll after
    # a pause to ensure the result is stable.
    if ($loaded) {
        $hasLooksGood = $script:uiaTexts | Where-Object { $_ -match 'Looks good|No issues found' }
        if ($hasLooksGood) {
            if ($DEBUG) { [Console]::Error.WriteLine("[EXCEL DEBUG] 'Looks good' detected - stability re-poll in 1500ms") }
            Start-Sleep -Milliseconds 1500
            Collect-UiaTexts $pidCond
            $pollCount++
            if ($DEBUG) {
                $sample = ($script:uiaTexts | Select-Object -First 10) -join ' | '
                [Console]::Error.WriteLine("[EXCEL DEBUG] Stability poll: nodes=$($script:uiaTexts.Count) sample=[$sample]")
            }
        }
    }

    # -- Current workbook window verification ----------------------------------
    # Confirm that the UIAutomation text we collected belongs to an Excel window
    # for the workbook we opened, not some other window.
    $wbStem = [System.IO.Path]::GetFileNameWithoutExtension($wbName)
    $verifiedWindowTitle = $null
    foreach ($t in $script:uiaTexts) {
        if ($wbStem.Length -gt 4 -and $t -match ([regex]::Escape($wbStem))) {
            $verifiedWindowTitle = $t
            break
        }
    }
    if ($DEBUG) {
        if ($verifiedWindowTitle) {
            [Console]::Error.WriteLine("[EXCEL DEBUG] Workbook verified in UIAutomation: '$verifiedWindowTitle'")
        } else {
            [Console]::Error.WriteLine("[EXCEL DEBUG] WARNING: Workbook name '$wbName' not found in UIAutomation. Check window isolation.")
        }
    }

    # -- Build output ----------------------------------------------------------
    if (-not $loaded) {
        $errMsg = "Accessibility Checker result content not found after ${MAX_MS}ms. Possible causes: pane in WebView2 (UIAutomation cannot read it), Excel did not finish analysis, or ExecuteMso toggle closed pane. Enable EXCEL_WORKER_DEBUG=true for raw UIAutomation text."
        $errOut = [ordered]@{
            ok           = $false
            error        = $errMsg
            excelVersion = $xlVersion
            executedMso  = $executedMso
            pollCount    = $pollCount
        }
        if ($DEBUG) {
            $cap = [Math]::Min(199, $script:uiaTexts.Count - 1)
            $errOut['rawOfficeText'] = if ($script:uiaTexts.Count -gt 0) { $script:uiaTexts.ToArray()[0..$cap] } else { @() }
        }
        $errOut | ConvertTo-Json -Depth 3 -Compress

    } else {
        $parsed = Parse-CheckerTexts -Trace:$DEBUG

        # Debug warning: images/charts present but zero alt-text issues reported
        if ($DEBUG) {
            $imgHints = @($script:uiaTexts | Where-Object { $_ -match '^Picture\s+\d+$|^Image\s+\d+$|^Chart\s+\d+$' })
            if ($imgHints.Count -gt 0 -and ((-not $parsed.counts.Contains('missingAltText')) -or $parsed.counts['missingAltText'] -eq 0)) {
                [Console]::Error.WriteLine("[EXCEL DEBUG] WARNING: Accessibility Checker reports 0 missing alt text, but $($imgHints.Count) picture/image/chart UIAutomation element(s) detected. The pane result may be stale.")
            }
        }

        $cap = [Math]::Min(199, $script:uiaTexts.Count - 1)
        $raw = if ($script:uiaTexts.Count -gt 0) { $script:uiaTexts.ToArray()[0..$cap] } else { @() }

        $out = [ordered]@{
            ok            = $true
            engine        = 'excel-ui-automation'
            excelVersion  = $xlVersion
            sheetCount    = $sheetCount
            executedMso   = $executedMso
            counts        = $parsed.counts
            statuses      = $parsed.statuses
            rawOfficeText = $raw
        }

        if ($DEBUG) {
            $out['openedFilePath']      = $absPath
            $out['openedFileName']      = $wbName
            $out['verifiedWindowTitle'] = $verifiedWindowTitle
            $out['pollCount']           = $pollCount
            $out['parseTrace']          = $parsed.parseTrace.ToArray()
        }

        $out | ConvertTo-Json -Depth 6 -Compress
    }

} catch {
    [ordered]@{
        ok    = $false
        error = "Excel Accessibility Assistant UI could not be read: $($_.Exception.Message)"
    } | ConvertTo-Json -Compress
} finally {
    if ($null -ne $wb) {
        try { $wb.Close($false) } catch {}    # close without saving
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } catch {}
        $wb = $null
    }
    if ($null -ne $xl) {
        try { $xl.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null } catch {}
        $xl = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
