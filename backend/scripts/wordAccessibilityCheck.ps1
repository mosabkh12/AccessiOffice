<#
.SYNOPSIS
  Opens a DOCX file in a NEW Word instance (visible), triggers the Accessibility
  Checker pane via ExecuteMso, then reads the pane content through UIAutomation.
  Outputs one JSON object to stdout. Debug messages go to stderr.

.PARAMETER FilePath
  Absolute path to the .docx file to scan.

.NOTES
  Env vars read at runtime:
    WORD_WORKER_DEBUG=true  - verbose stderr logs + extra debug fields in JSON output
  Requires: Windows, Microsoft Word installed (desktop edition, 2016+).

  Key isolation strategy:
  - [Activator]::CreateInstance forces a NEW WINWORD.EXE via CoCreateInstance,
    rather than reusing a running instance (GetActiveObject). This prevents
    reading a stale Accessibility pane from a prior scan session.
  - UIAutomation is filtered to that process's PID, obtained from the window
    handle (HWND) via GetWindowThreadProcessId.
  - Any pre-existing Accessibility pane is dismissed before re-triggering,
    ensuring the checker runs fresh for the current document.
  - The pane is not considered "ready" until result-category-level text
    appears (not just the pane title). If "Looks good" appears, one extra
    stability poll is done before accepting the result.
#>
param([Parameter(Mandatory=$true)][string]$FilePath)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$DEBUG   = ($env:WORD_WORKER_DEBUG -eq 'true')
$OPEN_MS = 2500   # ms after Documents.Open before any ExecuteMso
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
public static class WinApiHelper {
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
    # True when a (possibly stale, from prior doc) pane title is visible
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
        'Media and Illustrations',
        'Document Structure',
        'Document Access',
        'Missing alt text',
        'Hard-to-read text',
        'Missing table header',
        'Merged or split cells',
        'Unclear hyperlink',
        'Document title',
        'No headings in document',
        'No headings',
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
    $counts         = [ordered]@{}
    $statuses       = [ordered]@{}
    $labelPositions = [ordered]@{}
    $parseTrace     = [System.Collections.Generic.List[object]]::new()

    # Keys are the final officeLikeSummary key names (no translation layer needed for Word)
    $ruleMap = [ordered]@{
        missingAltText       = 'Missing Alt(?:ernative)? [Tt]ext|Missing alternative text|Alt Text.*Missing|Missing.*alt text'
        contrast             = 'Hard.to.Read(?:\s+\w+)*|Hard-to-Read|Hard to Read|Color Contrast|Insufficient.*[Cc]ontrast'
        tableHeader          = 'Missing [Tt]able [Hh]eader|Table [Hh]eader [Rr]ow [Mm]issing|Header [Rr]ow.*[Mm]issing'
        mergedCells          = 'Merged (?:or Split )?[Cc]ells|Split [Cc]ells|Use of [Mm]erged'
        unclearHyperlinkText = 'Unclear [Hh]yperlink [Tt]ext|Unclear [Ll]ink [Tt]ext|Unclear [Hh]yperlink'
        documentTitle        = '(?:Missing )?[Dd]ocument [Tt]itle|No [Dd]ocument [Tt]itle'
        noHeadings           = 'No [Hh]eadings?(?: in [Dd]ocument)?|Missing [Hh]eadings?'
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
        # This represents whole-document totals and is always preferred.
        for ($i = 0; $i -lt ($textLen - 1) -and -not $found; $i++) {
            $t = $textArr[$i]
            if ([regex]::IsMatch($t, '\d')) { continue }    # skip lines with digits
            if (-not [regex]::IsMatch($t, "(?:$pat)", $ignCase)) { continue }
            $next = $textArr[$i + 1]
            if ($next -match '^\d+$') {
                $n = [int]$next
                if ($n -gt 0) { $counts[$key] = $n }
                $found = $true
                if (-not $labelPositions.Contains($key)) { $labelPositions[$key] = $i }
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
    $autoStatusKeys = @('missingAltText', 'contrast', 'tableHeader', 'mergedCells', 'unclearHyperlinkText', 'documentTitle', 'noHeadings')
    foreach ($sk in $autoStatusKeys) {
        $statuses[$sk] = if ($counts.Contains($sk) -and $counts[$sk] -gt 0) { 'failed' } else { 'passed' }
    }
    $statuses['restrictedAccess'] = 'manual'   # always manual - not exposed through this pane path

    # -- Pass 3: Per-occurrence item extraction -----------------------------------
    # Items appear immediately after their category label+count in the UIA text.
    # Collects up to 15 items per category; stops at the next category label.
    $occurrences    = [ordered]@{}
    $sortedLabelKeys = @($labelPositions.Keys | Sort-Object { $labelPositions[$_] })

    for ($ki = 0; $ki -lt $sortedLabelKeys.Count; $ki++) {
        $key    = $sortedLabelKeys[$ki]
        $catPos = $labelPositions[$key]
        $count  = if ($counts.Contains($key)) { [int]$counts[$key] } else { 0 }
        if ($count -le 0) { continue }

        $rangeEnd = $textLen
        if ($ki + 1 -lt $sortedLabelKeys.Count) {
            $nk = $sortedLabelKeys[$ki + 1]
            $np = $labelPositions[$nk]
            if ($np -gt $catPos) { $rangeEnd = $np }
        }

        $maxItems = [Math]::Min($count, 15)
        $itemList = [System.Collections.Generic.List[string]]::new()

        for ($j = ($catPos + 2); $j -lt $rangeEnd -and $itemList.Count -lt $maxItems; $j++) {
            if ($j -ge $textLen) { break }
            $t = $textArr[$j]
            if (-not $t -or $t.Length -lt 2 -or $t.Length -gt 150) { continue }
            if ($t -match '^\d+$') { continue }
            if ($t.Length -le 3 -and [int][char]$t[0] -ge 0x2600) { continue }
            if ($t -match '^(?:Errors?|Warnings?|Tips?|Inspection Results|Accessibility Checker|Accessibility Assistant|Looks good|No issues|Check)') { continue }
            $hitCat = $false
            foreach ($re in $ruleMap.GetEnumerator()) {
                if (-not [regex]::IsMatch($t, '\d') -and [regex]::IsMatch($t, "(?:$($re.Value))", $ignCase)) {
                    $hitCat = $true; break
                }
            }
            if ($hitCat) { break }
            [void]$itemList.Add($t)
        }

        if ($itemList.Count -gt 0) { $occurrences[$key] = $itemList.ToArray() }
    }

    return [PSCustomObject]@{ counts = $counts; statuses = $statuses; parseTrace = $parseTrace; occurrences = $occurrences }
}

# --------------------------------------------------------------------------
# --------------------------------------------------------------------------
# COM-based occurrence enumeration for Word
# --------------------------------------------------------------------------
function Get-Word-COMOccurrences {
    param([object]$Doc, $Counts, [bool]$Debug)

    $occ = [ordered]@{}
    try {   # top-level guard — any COM exception returns empty occ instead of crashing the worker

    # ── Missing alt text (inline shapes + floating shapes) ────────────────────
    if ($Counts.Contains('missingAltText') -and [int]$Counts['missingAltText'] -gt 0) {
        $target = [int]$Counts['missingAltText']
        $items  = [System.Collections.Generic.List[object]]::new()

        $inlines = try { $Doc.InlineShapes } catch { $null }
        if ($null -ne $inlines) {
            for ($i = 1; $i -le $inlines.Count -and $items.Count -lt $target; $i++) {
                try {
                    $sh  = $inlines.Item($i)
                    $alt = try { $sh.AlternativeText } catch { '' }
                    if ($alt -and $alt.Trim()) { continue }
                    $pn  = try { [int]$sh.Range.Information(1) } catch { 0 }  # wdActiveEndPageNumber=1
                    $nm  = "Picture $i"
                    [void]$items.Add([ordered]@{
                        index=($items.Count+1); key='missingAltText'
                        pageNumber=$pn; objectName=$nm; location="Page $pn - $nm"
                        source='Microsoft Word Accessibility Checker'
                    })
                } catch {}
            }
        }
        $floats = try { $Doc.Shapes } catch { $null }
        if ($null -ne $floats) {
            for ($i = 1; $i -le $floats.Count -and $items.Count -lt $target; $i++) {
                try {
                    $sh  = $floats.Item($i)
                    $alt = try { $sh.AlternativeText } catch { '' }
                    $ttl = try { $sh.Title }           catch { '' }
                    if (($alt -and $alt.Trim()) -or ($ttl -and $ttl.Trim())) { continue }
                    $pn  = try { [int]$sh.Anchor.Information(1) } catch { 0 }
                    $nm  = try { $sh.Name } catch { "Shape $i" }
                    [void]$items.Add([ordered]@{
                        index=($items.Count+1); key='missingAltText'
                        pageNumber=$pn; objectName=$nm; location="Page $pn - $nm"
                        source='Microsoft Word Accessibility Checker'
                    })
                } catch {}
            }
        }
        if ($items.Count -gt 0) { $occ['missingAltText'] = $items.ToArray() }
        if ($Debug) { [Console]::Error.WriteLine("[WORD COM] missingAltText: $($items.Count) of $target") }
    }

    # ── Merged cells ──────────────────────────────────────────────────────────
    if ($Counts.Contains('mergedCells') -and [int]$Counts['mergedCells'] -gt 0) {
        $target = [int]$Counts['mergedCells']
        $items  = [System.Collections.Generic.List[object]]::new()
        $tables = try { $Doc.Tables } catch { $null }
        $tNum   = 0
        if ($null -ne $tables) {
            for ($t = 1; $t -le $tables.Count -and $items.Count -lt $target; $t++) {
                try {
                    $tbl  = $tables.Item($t); $tNum++
                    $pn   = try { [int]$tbl.Range.Information(1) } catch { 0 }
                    for ($r = 1; $r -le $tbl.Rows.Count -and $items.Count -lt $target; $r++) {
                        for ($c = 1; $c -le $tbl.Columns.Count -and $items.Count -lt $target; $c++) {
                            try {
                                $cell = $tbl.Cell($r, $c)
                                if ($cell.MergeCells) {
                                    [void]$items.Add([ordered]@{
                                        index=($items.Count+1); key='mergedCells'
                                        pageNumber=$pn; objectName="Table $tNum"
                                        location="Page $pn - Table $tNum (Row $r Col $c)"
                                        source='Microsoft Word Accessibility Checker'
                                    })
                                }
                            } catch {}
                        }
                    }
                } catch {}
            }
        }
        if ($items.Count -gt 0) { $occ['mergedCells'] = $items.ToArray() }
        if ($Debug) { [Console]::Error.WriteLine("[WORD COM] mergedCells: $($items.Count) of $target") }
    }

    # ── No headings ───────────────────────────────────────────────────────────
    if ($Counts.Contains('noHeadings') -and [int]$Counts['noHeadings'] -gt 0) {
        # Document-level issue: the whole document has no heading styles
        $occ['noHeadings'] = @([ordered]@{
            index=1; key='noHeadings'; pageNumber=0; objectName='Entire document'
            location='Entire document - No headings defined'
            source='Microsoft Word Accessibility Checker'
        })
    }

    # ── Unclear hyperlink text ────────────────────────────────────────────────
    if ($Counts.Contains('unclearHyperlinkText') -and [int]$Counts['unclearHyperlinkText'] -gt 0) {
        $target  = [int]$Counts['unclearHyperlinkText']
        $items   = [System.Collections.Generic.List[object]]::new()
        $badPat  = '^(?:click here|here|link|this|read more|more)$'
        $hypers  = try { $Doc.Hyperlinks } catch { $null }
        if ($null -ne $hypers) {
            foreach ($hl in $hypers) {
                if ($items.Count -ge $target) { break }
                try {
                    $disp = try { $hl.TextToDisplay.Trim() } catch { '' }
                    if (-not $disp -or -not [regex]::IsMatch($disp, $badPat, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) { continue }
                    $pn   = try { [int]$hl.Range.Information(1) } catch { 0 }
                    [void]$items.Add([ordered]@{
                        index=($items.Count+1); key='unclearHyperlinkText'
                        pageNumber=$pn; objectName=$disp
                        location="Page $pn - Hyperlink"
                        source='Microsoft Word Accessibility Checker'
                    })
                } catch {}
            }
        }
        if ($items.Count -gt 0) { $occ['unclearHyperlinkText'] = $items.ToArray() }
    }

    } catch {
        if ($Debug) { [Console]::Error.WriteLine("[WORD COM] Unhandled exception in Get-Word-COMOccurrences: $($_.Exception.Message)") }
    }

    return $occ
}

# Input validation
# --------------------------------------------------------------------------
if (-not (Test-Path $FilePath)) {
    [ordered]@{ ok = $false; error = "File not found: $FilePath" } | ConvertTo-Json -Compress
    exit 0
}

$absPath = (Resolve-Path $FilePath).Path
$word    = $null
$doc     = $null
$wordPid = 0

try {
    # -- Force a NEW WINWORD.EXE instance --------------------------------------
    # [Activator]::CreateInstance calls CoCreateInstance, which always launches
    # a new local-server EXE, unlike New-Object -ComObject which calls
    # GetActiveObject first and reuses any running Word instance.
    # This is critical: reusing an existing instance can leave a stale
    # Accessibility pane open from a previous document.
    if ($DEBUG) { [Console]::Error.WriteLine("[WORD DEBUG] Creating new Word.Application via Activator::CreateInstance") }
    $wordType = [Type]::GetTypeFromProgID('Word.Application')
    if ($null -eq $wordType) { throw "Word.Application ProgID not registered - is Microsoft Word installed?" }
    $word = [Activator]::CreateInstance($wordType)

    $word.Visible          = $true
    $word.DisplayAlerts    = 0     # wdAlertsNone
    $word.AutomationSecurity = 3   # msoAutomationSecurityForceDisable

    # -- Open document read-only -----------------------------------------------
    if ($DEBUG) { [Console]::Error.WriteLine("[WORD DEBUG] Opening: $absPath") }
    # Documents.Open(FileName, ConfirmConversions, ReadOnly)
    $doc     = $word.Documents.Open($absPath, $false, $true)
    $docName = $doc.Name
    try { $doc.Activate() } catch {}   # bring our document to front

    $wordVersion = $word.Version
    if ($DEBUG) { [Console]::Error.WriteLine("[WORD DEBUG] Opened: version=$wordVersion docName=$docName") }

    # -- Get PID from HWND (reliable even when multiple WINWORD instances exist) -
    try {
        $hwnd = [IntPtr]$word.Hwnd
        [WinApiHelper]::GetWindowThreadProcessId($hwnd, [ref]$wordPid) | Out-Null
        if ($DEBUG) { [Console]::Error.WriteLine("[WORD DEBUG] PID from HWND: $wordPid") }
    } catch {
        # Fallback: pick the most recently started WINWORD process
        $procs = [System.Diagnostics.Process]::GetProcessesByName('WINWORD')
        if ($procs.Count -gt 0) {
            $wordPid = ($procs | Sort-Object StartTime -Descending | Select-Object -First 1).Id
        }
        if ($DEBUG) { [Console]::Error.WriteLine("[WORD DEBUG] PID via fallback: $wordPid (HWND method failed: $($_.Exception.Message))") }
    }

    if ($wordPid -eq 0) { throw "Could not determine Word process ID" }

    $root    = [System.Windows.Automation.AutomationElement]::RootElement
    $pidCond = [System.Windows.Automation.PropertyCondition]::new(
                   [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $wordPid)

    Start-Sleep -Milliseconds $OPEN_MS

    # -- Dismiss any stale Accessibility pane ----------------------------------
    # Even with a new Word instance, if Word inherits task-pane state from a
    # previous session (e.g. roaming profile), the pane may already be open.
    # Toggle it off first so the fresh trigger runs against the current doc.
    Collect-UiaTexts $pidCond
    if (Test-StalePane) {
        if ($DEBUG) { [Console]::Error.WriteLine("[WORD DEBUG] Stale accessibility pane detected - dismissing before fresh trigger") }
        foreach ($id in @('ReviewCheckAccessibility', 'CheckAccessibility', 'AccessibilityChecker')) {
            try { $word.CommandBars.ExecuteMso($id); Start-Sleep -Milliseconds 600; break } catch {}
        }
    }

    # -- Trigger Accessibility Checker pane ------------------------------------
    $executedMso = $null
    $msoAttempts = @('ReviewCheckAccessibility', 'CheckAccessibility', 'AccessibilityChecker', 'ReviewAccessibilityChecker')
    foreach ($id in $msoAttempts) {
        try {
            $word.CommandBars.ExecuteMso($id)
            $executedMso = $id
            if ($DEBUG) { [Console]::Error.WriteLine("[WORD DEBUG] ExecuteMso '$id' succeeded") }
            break
        } catch {
            if ($DEBUG) { [Console]::Error.WriteLine("[WORD DEBUG] ExecuteMso '$id' failed: $($_.Exception.Message)") }
        }
    }

    if ($null -eq $executedMso) {
        $tried = $msoAttempts -join ', '
        throw "Could not open Accessibility Checker pane. All ExecuteMso IDs failed: $tried. Ensure Word 2016+ is installed and the Review tab is accessible."
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
            [Console]::Error.WriteLine("[WORD DEBUG] Poll $pollCount (${elapsed}ms): nodes=$($script:uiaTexts.Count) sample=[$sample]")
        }

        if (Test-PaneReady) { $loaded = $true; break }
    }

    # -- Stability re-poll when "Looks good" appears ---------------------------
    # "Looks good! No issues found." may appear before Word has finished
    # running the checker on the current document. Do one extra poll after
    # a pause to ensure the result is stable.
    if ($loaded) {
        $hasLooksGood = $script:uiaTexts | Where-Object { $_ -match 'Looks good|No issues found' }
        if ($hasLooksGood) {
            if ($DEBUG) { [Console]::Error.WriteLine("[WORD DEBUG] 'Looks good' detected - stability re-poll in 1500ms") }
            Start-Sleep -Milliseconds 1500
            Collect-UiaTexts $pidCond
            $pollCount++
            if ($DEBUG) {
                $sample = ($script:uiaTexts | Select-Object -First 10) -join ' | '
                [Console]::Error.WriteLine("[WORD DEBUG] Stability poll: nodes=$($script:uiaTexts.Count) sample=[$sample]")
            }
        }
    }

    # -- Current document window verification ----------------------------------
    # Confirm that the UIAutomation text we collected belongs to a Word window
    # for the document we opened, not some other window.
    $docStem = [System.IO.Path]::GetFileNameWithoutExtension($docName)
    $verifiedWindowTitle = $null
    foreach ($t in $script:uiaTexts) {
        if ($docStem.Length -gt 4 -and $t -match ([regex]::Escape($docStem))) {
            $verifiedWindowTitle = $t
            break
        }
    }
    if ($DEBUG) {
        if ($verifiedWindowTitle) {
            [Console]::Error.WriteLine("[WORD DEBUG] Document verified in UIAutomation: '$verifiedWindowTitle'")
        } else {
            [Console]::Error.WriteLine("[WORD DEBUG] WARNING: Document name '$docName' not found in UIAutomation. Check window isolation.")
        }
    }

    # -- Get page count (optional, best-effort) --------------------------------
    $pageCount = $null
    try { $pageCount = $doc.ComputeStatistics(2) } catch {}   # 2 = wdStatisticPages

    # -- Build output ----------------------------------------------------------
    if (-not $loaded) {
        $errMsg = "Accessibility Checker result content not found after ${MAX_MS}ms. Possible causes: pane in WebView2 (UIAutomation cannot read it), Word did not finish analysis, or ExecuteMso toggle closed pane. Enable WORD_WORKER_DEBUG=true for raw UIAutomation text."
        $errOut = [ordered]@{
            ok          = $false
            error       = $errMsg
            wordVersion = $wordVersion
            executedMso = $executedMso
            pollCount   = $pollCount
        }
        if ($DEBUG) {
            $cap = [Math]::Min(199, $script:uiaTexts.Count - 1)
            $errOut['rawOfficeText'] = if ($script:uiaTexts.Count -gt 0) { $script:uiaTexts.ToArray()[0..$cap] } else { @() }
        }
        $errOut | ConvertTo-Json -Depth 3 -Compress

    } else {
        $parsed = Parse-CheckerTexts -Trace:$DEBUG

        # Debug warning: pictures present but zero alt-text issues reported
        if ($DEBUG) {
            $picHints = @($script:uiaTexts | Where-Object { $_ -match '^Picture\s+\d+$|^Image\s+\d+$' })
            if ($picHints.Count -gt 0 -and ((-not $parsed.counts.Contains('missingAltText')) -or $parsed.counts['missingAltText'] -eq 0)) {
                [Console]::Error.WriteLine("[WORD DEBUG] WARNING: Accessibility Checker reports 0 missing alt text, but $($picHints.Count) picture/image UIAutomation element(s) detected. The pane result may be stale.")
            }
        }

        $cap = [Math]::Min(199, $script:uiaTexts.Count - 1)
        $raw = if ($script:uiaTexts.Count -gt 0) { $script:uiaTexts.ToArray()[0..$cap] } else { @() }

        # COM-based occurrence extraction (reads from $doc directly — works with all Office versions)
        $comOcc   = Get-Word-COMOccurrences -Doc $doc -Counts $parsed.counts -Debug:$DEBUG
        $finalOcc = [ordered]@{}
        # COM results take priority; text-scan (Pass 3) fills gaps
        foreach ($k in $comOcc.Keys) {
            $arr = @($comOcc[$k])
            if ($arr.Count -gt 0) { $finalOcc[$k] = $arr }
        }
        foreach ($k in $parsed.occurrences.Keys) {
            if (-not $finalOcc.Contains($k) -or @($finalOcc[$k]).Count -eq 0) {
                $arr = @($parsed.occurrences[$k])
                if ($arr.Count -gt 0) { $finalOcc[$k] = $arr }
            }
        }

        $out = [ordered]@{
            ok            = $true
            engine        = 'word-ui-automation'
            wordVersion   = $wordVersion
            pageCount     = $pageCount
            executedMso   = $executedMso
            counts        = $parsed.counts
            statuses      = $parsed.statuses
            occurrences   = $finalOcc
            rawOfficeText = $raw
        }

        if ($DEBUG) {
            $out['openedFilePath']      = $absPath
            $out['openedFileName']      = $docName
            $out['verifiedWindowTitle'] = $verifiedWindowTitle
            $out['pollCount']           = $pollCount
            $out['parseTrace']          = $parsed.parseTrace.ToArray()
        }

        $out | ConvertTo-Json -Depth 8 -Compress
    }

} catch {
    [ordered]@{
        ok    = $false
        error = "Word Accessibility Assistant UI could not be read: $($_.Exception.Message)"
    } | ConvertTo-Json -Compress
} finally {
    if ($null -ne $doc) {
        try { $doc.Close(0) } catch {}    # 0 = wdDoNotSaveChanges
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } catch {}
        $doc = $null
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
        $word = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
