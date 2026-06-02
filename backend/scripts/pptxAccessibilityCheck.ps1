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
    $counts         = [ordered]@{}
    $statuses       = [ordered]@{}
    $labelPositions = [ordered]@{}

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
                if (-not $labelPositions.Contains($key)) { $labelPositions[$key] = $i }
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

    # -- Pass 3: Per-occurrence item extraction -----------------------------------
    $occurrences    = [ordered]@{}
    $sortedLabelKeys = @($labelPositions.Keys | Sort-Object { $labelPositions[$_] })

    for ($ki = 0; $ki -lt $sortedLabelKeys.Count; $ki++) {
        $key    = $sortedLabelKeys[$ki]
        $catPos = $labelPositions[$key]
        $count  = if ($counts.Contains($key)) { [int]$counts[$key] } else { 0 }
        if ($count -le 0) { continue }

        $rangeEnd = $textLen
        if ($ki + 1 -lt $sortedLabelKeys.Count) {
            $nk = $sortedLabelKeys[$ki + 1]; $np = $labelPositions[$nk]
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
                if (-not [regex]::IsMatch($t, '\d') -and [regex]::IsMatch($t, "(?:$($re.Value))", $ignCase)) { $hitCat = $true; break }
            }
            if ($hitCat) { break }
            [void]$itemList.Add($t)
        }

        if ($itemList.Count -gt 0) { $occurrences[$key] = $itemList.ToArray() }
    }

    return [PSCustomObject]@{ counts = $counts; statuses = $statuses; occurrences = $occurrences }
}

# --------------------------------------------------------------------------
# Scan pane UIAutomation tree for occurrence elements belonging to each category
# --------------------------------------------------------------------------
# Returns: ordered hashtable { key -> List[AutomationElement] }
# BFS from pane windows; when a node matches a category pattern, all invokable
# descendants (up to their own children) are collected under that key.
function Scan-PaneForCategories {
    param(
        [System.Windows.Automation.AutomationElement[]]$Wins,
        $PatternMap,
        [System.Text.RegularExpressions.RegexOptions]$IgnCase,
        [int]$MaxNodes = 6000
    )
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $result = [ordered]@{}
    foreach ($k in $PatternMap.Keys) {
        $result[$k] = [System.Collections.Generic.List[System.Windows.Automation.AutomationElement]]::new()
    }

    $queue = [System.Collections.Generic.Queue[object]]::new()
    foreach ($w in $Wins) { $queue.Enqueue([PSCustomObject]@{ El=$w; D=0; CK=$null }) }

    $n = 0
    while ($queue.Count -gt 0 -and $n -lt $MaxNodes) {
        $fr = $queue.Dequeue(); $n++
        $el = $fr.El; $d = $fr.D; $ck = $fr.CK
        if ($d -gt 10) { continue }

        $nm = try { $el.Current.Name } catch { '' }

        if ($null -ne $ck) {
            # Inside a category subtree — collect invokable elements (occurrence items)
            $inv = $false
            try { $el.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern); $inv = $true } catch {}
            if (-not $inv) { try { $el.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern); $inv = $true } catch {} }

            if ($inv -and $nm -and $nm.Length -gt 2) {
                [void]$result[$ck].Add($el)
                # Occurrence items are leaves — do not recurse into their children
            } else {
                # Non-invokable container inside category — keep descending
                $ch = try { $walker.GetFirstChild($el) } catch { $null }
                while ($null -ne $ch) {
                    $queue.Enqueue([PSCustomObject]@{ El=$ch; D=$d+1; CK=$ck })
                    $ch = try { $walker.GetNextSibling($ch) } catch { $null }
                }
            }
        } else {
            # Outside any category — check if this element IS a category header
            $matchKey = $null
            foreach ($entry in $PatternMap.GetEnumerator()) {
                if ($nm -and [regex]::IsMatch($nm, "(?:$($entry.Value))", $IgnCase)) {
                    $matchKey = $entry.Key; break
                }
            }
            $ch = try { $walker.GetFirstChild($el) } catch { $null }
            while ($null -ne $ch) {
                $queue.Enqueue([PSCustomObject]@{ El=$ch; D=$d+1; CK=$matchKey })
                $ch = try { $walker.GetNextSibling($ch) } catch { $null }
            }
        }
    }
    return $result
}

# --------------------------------------------------------------------------
# Invoke each occurrence element, read PowerPoint COM state, return structured
# occurrence objects  { index, key, slideNumber, objectName, shapeId, location, source }
# --------------------------------------------------------------------------
function Extract-PPT-Occurrences {
    param(
        [System.Windows.Automation.AutomationElement[]]$Wins,
        $Counts,
        [object]$PptApp,
        [bool]$Debug,
        [int]$WaitMs   = 320,
        [int]$BudgetMs = 30000   # No per-key cap — try to extract all Office-reported occurrences
    )
    $ignCase = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase

    # Patterns for categories whose occurrences we want to extract
    $extractPats = [ordered]@{
        missingAltText      = 'Missing Alt(?:ernative)? [Tt]ext|Missing alternative text'
        missingSlideTitle   = 'Missing Slide Title|Slide title.*missing|Missing.*slide title'
        duplicateSlideTitle = 'Duplicate Slide Title|Duplicate.*slide title'
        hardToReadText      = 'Hard.to.Read(?:\s+\w+)*|Hard-to-Read|Hard to Read|Color Contrast|Insufficient.*[Cc]ontrast'
        missingTableHeader  = 'Missing Table Header|Table.*Header Row|Missing table header'
        mergedCells         = 'Merged(?:.*)[Cc]ells|Split Cells|Use of merged or split cells'
        mediaCaptions       = 'Missing.*(?:Audio|Video|Subtitles|Captions)|(?:Audio|Video).*[Cc]aptions?.*[Mm]issing'
        readingOrder        = '(?:Check )?Reading Order'
    }

    # Only keep patterns for keys where Office reported a failure
    $activePats = [ordered]@{}
    foreach ($entry in $extractPats.GetEnumerator()) {
        if ($Counts.Contains($entry.Key) -and [int]$Counts[$entry.Key] -gt 0) {
            $activePats[$entry.Key] = $entry.Value
        }
    }

    if ($activePats.Count -eq 0) {
        return [ordered]@{ occurrences=[ordered]@{}; occurrencesExtracted=$false; occurrencesNote='No active failures to extract.' }
    }

    if ($Debug) { [Console]::Error.WriteLine("[PPT OCC] Starting extraction for keys: $($activePats.Keys -join ', ')") }

    # Single BFS scan: find all occurrence elements per category
    $catElements = Scan-PaneForCategories -Wins $Wins -PatternMap $activePats -IgnCase $ignCase

    $resultOcc         = [ordered]@{}
    $anyOk             = $false
    $allFullyExtracted = $true    # false when any key ends with fewer than its Office count
    $globalBudgetHit   = $false
    $partialNotes      = [System.Collections.Generic.List[string]]::new()
    $trace             = [System.Collections.Generic.List[object]]::new()
    $tStart            = [DateTime]::UtcNow

    foreach ($key in $activePats.Keys) {
        $elapsed = ([DateTime]::UtcNow - $tStart).TotalMilliseconds
        if ($elapsed -gt $BudgetMs) {
            $globalBudgetHit   = $true
            $allFullyExtracted = $false
            [void]$partialNotes.Add("$($key): budget exhausted before processing")
            if ($Debug) { [Console]::Error.WriteLine("[PPT OCC] Budget exhausted (${elapsed}ms) before key=$key") }
            break
        }

        $count      = [int]$Counts[$key]
        $occEls     = $catElements[$key]
        # targetCount = Office-reported count — no artificial cap
        $maxExtract = $count

        if ($Debug) { [Console]::Error.WriteLine("[PPT OCC] key=$key target=$count candidateEls=$($occEls.Count)") }

        $extracted    = [System.Collections.Generic.List[object]]::new()
        $seen         = [System.Collections.Generic.HashSet[string]]::new()
        $keyBudgetHit = $false

        foreach ($occEl in $occEls) {
            if ($extracted.Count -ge $maxExtract) { break }   # target reached
            $timeNow = ([DateTime]::UtcNow - $tStart).TotalMilliseconds
            if ($timeNow -gt $BudgetMs) { $keyBudgetHit = $true; $globalBudgetHit = $true; break }

            $elName  = try { $occEl.Current.Name } catch { '' }
            $invoked = $false

            # Try InvokePattern first (preferred — navigates to the object)
            try {
                $ip = $occEl.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                $ip.Invoke()
                Start-Sleep -Milliseconds $WaitMs
                $invoked = $true
            } catch {
                # Fall back to SelectionItemPattern
                try {
                    $sp = $occEl.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
                    $sp.Select()
                    Start-Sleep -Milliseconds $WaitMs
                    $invoked = $true
                } catch {}
            }

            if (-not $invoked) {
                if ($Debug) { [Console]::Error.WriteLine("[PPT OCC] Could not invoke element: $elName") }
                continue
            }

            # Read COM state after PowerPoint navigated to the selected object
            $slideNum = 0; $shapeName = ''; $shapeId = 0; $shapeType = ''

            try { $slideNum = [int]$PptApp.ActiveWindow.View.Slide.SlideIndex } catch {}
            try {
                $selType = $PptApp.ActiveWindow.Selection.Type
                if ($selType -eq 2) {   # ppSelectionShapes = 2
                    $sh        = $PptApp.ActiveWindow.Selection.ShapeRange.Item(1)
                    $shapeName = try { $sh.Name }           catch { '' }
                    $shapeId   = try { [int]$sh.Id }        catch { 0 }
                    $shapeType = try { $sh.Type.ToString() } catch { '' }
                }
            } catch {}

            # Deduplicate by slide + shape identity
            $dk = "$key|$slideNum|$(if($shapeId -gt 0){$shapeId}else{$shapeName})"
            if ($seen.Contains($dk)) { continue }
            [void]$seen.Add($dk)

            # Build ASCII-safe location string: "Slide N - ShapeName"
            $loc = if ($slideNum -gt 0) { "Slide $slideNum" } else { $elName }
            if ($shapeName) { $loc = if ($slideNum -gt 0) { "Slide $slideNum - $shapeName" } else { $shapeName } }

            [void]$extracted.Add([ordered]@{
                index       = $extracted.Count + 1
                key         = $key
                slideNumber = $slideNum
                objectName  = $shapeName
                shapeId     = $shapeId
                shapeType   = $shapeType
                location    = $loc
                source      = 'Microsoft PowerPoint Accessibility Checker'
            })

            if ($Debug) {
                [void]$trace.Add([ordered]@{
                    key=$key; elName=$elName; invoked=$invoked
                    slideNum=$slideNum; shapeName=$shapeName; shapeId=$shapeId; loc=$loc
                })
                [Console]::Error.WriteLine("[PPT OCC] Extracted: key=$key loc=$loc slideNum=$slideNum shapeName=$shapeName")
            }
        }

        if ($extracted.Count -gt 0) { $resultOcc[$key] = $extracted.ToArray(); $anyOk = $true }

        # Track whether this key was fully extracted
        if ($extracted.Count -lt $count) {
            $allFullyExtracted = $false
            if ($keyBudgetHit) {
                [void]$partialNotes.Add("$($key): $($extracted.Count) of $count (timeout)")
            } else {
                # No more invokable elements in pane for this category
                [void]$partialNotes.Add("$($key): $($extracted.Count) of $count (no more pane elements)")
            }
        }

        if ($Debug) { [Console]::Error.WriteLine("[PPT OCC] key=$key extracted=$($extracted.Count) of $count budgetHit=$keyBudgetHit") }
    }

    $note = $null
    if (-not $anyOk) {
        $note = 'Microsoft Office Accessibility Checker occurrence elements could not be invoked. The pane may use WebView2 (Office 365) or a version of Office where UIAutomation cannot reach individual occurrence items.'
    } elseif (-not $allFullyExtracted) {
        $reason = if ($globalBudgetHit) { 'timeout' } else { 'not all items visible in pane' }
        $detail = $partialNotes -join '; '
        $note   = "Partial extraction ($reason): $detail. Open the file in Microsoft Office for the complete list."
    }
    # else: note = null — complete extraction, all keys reached their Office count target

    $ret = [ordered]@{
        occurrences          = $resultOcc
        occurrencesExtracted = $anyOk
        occurrencesNote      = $note
    }
    if ($Debug) { $ret['occurrenceTrace'] = $trace.ToArray() }
    return $ret
}

# --------------------------------------------------------------------------
# COM-based occurrence enumeration
# Reads directly from the PowerPoint presentation object (no UIAutomation pane
# needed).  Used as the primary source / fallback when pane clicking returns 0.
# Returns structured occurrence objects identical to Extract-PPT-Occurrences.
# --------------------------------------------------------------------------
function Get-PPTX-COMOccurrences {
    param([object]$Pres, $Counts, [bool]$Debug)

    $occ = [ordered]@{}
    try {   # top-level guard

    # ── Missing alt text ──────────────────────────────────────────────────────
    if ($Counts.Contains('missingAltText') -and [int]$Counts['missingAltText'] -gt 0) {
        $target = [int]$Counts['missingAltText']
        $items  = [System.Collections.Generic.List[object]]::new()
        $seen   = [System.Collections.Generic.HashSet[string]]::new()

        foreach ($slide in $Pres.Slides) {
            if ($items.Count -ge $target) { break }
            $sn = try { [int]$slide.SlideIndex } catch { 0 }
            foreach ($shape in $slide.Shapes) {
                if ($items.Count -ge $target) { break }
                try {
                    $alt = try { $shape.AlternativeText } catch { '' }
                    $ttl = try { $shape.Title }           catch { '' }
                    if (($alt -and $alt.Trim()) -or ($ttl -and $ttl.Trim())) { continue }
                    $st  = try { [int]$shape.Type } catch { -1 }
                    if ($st -eq 17) { continue }   # msoTextBox — text-only, Office usually skips
                    $nm  = try { $shape.Name } catch { "Shape" }
                    $sid = try { [int]$shape.Id } catch { 0 }
                    $dk  = "$sn|$sid"
                    if ($seen.Contains($dk)) { continue }
                    [void]$seen.Add($dk)
                    [void]$items.Add([ordered]@{
                        index=($items.Count+1); key='missingAltText'; slideNumber=$sn
                        objectName=$nm; shapeId=$sid; location="Slide $sn - $nm"
                        source='Microsoft PowerPoint Accessibility Checker'
                    })
                } catch {}
            }
        }
        if ($items.Count -gt 0) { $occ['missingAltText'] = $items.ToArray() }
        if ($Debug) { [Console]::Error.WriteLine("[PPT COM] missingAltText: $($items.Count) of $target") }
    }

    # ── Missing slide title ───────────────────────────────────────────────────
    if ($Counts.Contains('missingSlideTitle') -and [int]$Counts['missingSlideTitle'] -gt 0) {
        $target = [int]$Counts['missingSlideTitle']
        $items  = [System.Collections.Generic.List[object]]::new()
        foreach ($slide in $Pres.Slides) {
            if ($items.Count -ge $target) { break }
            $sn     = try { [int]$slide.SlideIndex } catch { 0 }
            $hasTtl = $false
            foreach ($shape in $slide.Shapes) {
                try {
                    $ph = try { $shape.PlaceholderFormat } catch { $null }
                    if ($null -eq $ph) { continue }
                    $pt = try { [int]$ph.Type } catch { -1 }
                    if ($pt -eq 2 -or $pt -eq 3) {   # ppPlaceholderTitle=2 / ppPlaceholderCenterTitle=3
                        $txt = try { $shape.TextFrame.TextRange.Text.Trim() } catch { '' }
                        if ($txt.Length -gt 0) { $hasTtl = $true; break }
                    }
                } catch {}
            }
            if (-not $hasTtl) {
                [void]$items.Add([ordered]@{
                    index=($items.Count+1); key='missingSlideTitle'; slideNumber=$sn
                    objectName=''; shapeId=0; location="Slide $sn"
                    source='Microsoft PowerPoint Accessibility Checker'
                })
            }
        }
        if ($items.Count -gt 0) { $occ['missingSlideTitle'] = $items.ToArray() }
    }

    # ── Duplicate slide title ─────────────────────────────────────────────────
    if ($Counts.Contains('duplicateSlideTitle') -and [int]$Counts['duplicateSlideTitle'] -gt 0) {
        $target   = [int]$Counts['duplicateSlideTitle']
        $items    = [System.Collections.Generic.List[object]]::new()
        $titleMap = @{}
        foreach ($slide in $Pres.Slides) {
            $sn  = try { [int]$slide.SlideIndex } catch { 0 }
            $txt = ''
            foreach ($shape in $slide.Shapes) {
                try {
                    $ph = try { $shape.PlaceholderFormat } catch { $null }
                    if ($null -eq $ph) { continue }
                    $pt = try { [int]$ph.Type } catch { -1 }
                    if ($pt -eq 2 -or $pt -eq 3) { $txt = try { $shape.TextFrame.TextRange.Text.Trim() } catch { '' }; break }
                } catch {}
            }
            if ($txt.Length -gt 0) {
                if (-not $titleMap.ContainsKey($txt)) { $titleMap[$txt] = [System.Collections.Generic.List[int]]::new() }
                [void]$titleMap[$txt].Add($sn)
            }
        }
        foreach ($kv in $titleMap.GetEnumerator()) {
            if ($items.Count -ge $target) { break }
            if ($kv.Value.Count -gt 1) {
                foreach ($sn in $kv.Value) {
                    if ($items.Count -ge $target) { break }
                    $safe = ($kv.Key -replace '"', "'")
                    [void]$items.Add([ordered]@{
                        index=($items.Count+1); key='duplicateSlideTitle'; slideNumber=$sn
                        objectName=$kv.Key; shapeId=0; location="Slide $sn - '$safe'"
                        source='Microsoft PowerPoint Accessibility Checker'
                    })
                }
            }
        }
        if ($items.Count -gt 0) { $occ['duplicateSlideTitle'] = $items.ToArray() }
    }

    } catch {
        if ($Debug) { [Console]::Error.WriteLine("[PPT COM] Unhandled exception in Get-PPTX-COMOccurrences: $($_.Exception.Message)") }
    }

    return $occ
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

        # -- Occurrence extraction: UIAutomation pane clicking (primary) ----------
        $occWins   = @($root.FindAll([System.Windows.Automation.TreeScope]::Children, $pidCond))
        $occResult = Extract-PPT-Occurrences -Wins $occWins -Counts $parsed.counts -PptApp $ppt -Debug:$DEBUG

        # -- COM-based fallback for keys UIAutomation could not extract ---------
        # Reads shapes directly from the PowerPoint COM object. Works even when the
        # Accessibility Assistant pane uses WebView2 and blocks UIAutomation clicking.
        $comOcc = Get-PPTX-COMOccurrences -Pres $pres -Counts $parsed.counts -Debug:$DEBUG

        $finalOcc      = [ordered]@{}
        $anyExtracted  = $false

        # Merge: prefer UIAutomation results; fall back to COM for missing keys
        foreach ($k in $occResult.occurrences.Keys) {
            $arr = @($occResult.occurrences[$k])
            if ($arr.Count -gt 0) { $finalOcc[$k] = $arr; $anyExtracted = $true }
        }
        foreach ($k in $comOcc.Keys) {
            if (-not $finalOcc.Contains($k) -or @($finalOcc[$k]).Count -eq 0) {
                $arr = @($comOcc[$k])
                if ($arr.Count -gt 0) { $finalOcc[$k] = $arr; $anyExtracted = $true }
            }
        }

        $finalNote = if ($anyExtracted) { $null } else { $occResult.occurrencesNote }

        $out = [ordered]@{
            ok                   = $true
            engine               = 'powerpoint-ui-automation'
            pptVersion           = $pptVersion
            slideCount           = $slideCount
            executedMso          = $executedMso
            counts               = $parsed.counts
            statuses             = $parsed.statuses
            occurrences          = $finalOcc
            occurrencesExtracted = $anyExtracted
            occurrencesNote      = $finalNote
            rawOfficeText        = $raw
        }
        if ($DEBUG -and $occResult.occurrenceTrace) { $out['occurrenceTrace'] = $occResult.occurrenceTrace }
        $out | ConvertTo-Json -Depth 8 -Compress
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
