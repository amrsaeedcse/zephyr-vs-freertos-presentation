param([string]$TargetPptx = (Join-Path (Split-Path $PSScriptRoot -Parent) "Zephyr_RTOS_Embedded_Systems.pptx"))
$ErrorActionPreference = "Stop"

# Verified PowerPoint COM constants
$TR_FADE   = 3849   # ppEffectFadeSmoothly
$TR_PUSH_L = 3853   # ppEffectPushLeft
$TR_REVEAL = 3894   # ppEffectRevealSmoothLeft

$AN_FADE   = 10     # msoAnimEffectFade
$AN_WIPE   = 22     # msoAnimEffectWipe
$AN_ASCEND = 39     # msoAnimEffectAscend
$AN_FZOOM  = 48     # msoAnimEffectFadedZoom

$TRIG_WITH = 2      # msoAnimTriggerWithPrevious
$DIR_UP    = 1      # msoAnimDirectionUp
$DIR_RIGHT = 2      # msoAnimDirectionRight (wipe left -> right)

function Add-Anim {
    param($slide, $shape, [int]$effect, [float]$duration, [float]$delay, [int]$dir = 0)
    try {
        $eff = $slide.TimeLine.MainSequence.AddEffect($shape, $effect, 0, $TRIG_WITH)
        $eff.Timing.Duration = $duration
        $eff.Timing.TriggerDelayTime = $delay
        $eff.Timing.SmoothEnd = -1 # msoTrue
        if ($dir -gt 0) {
            try { $eff.EffectParameters.Direction = $dir } catch {}
        }
        return $eff
    } catch {
        return $null
    }
}

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Open($TargetPptx, $false, $false, $false)

$totalEffects = 0
$slideCount = $pres.Slides.Count

for ($sn = 1; $sn -le $slideCount; $sn++) {
    $sl = $pres.Slides.Item($sn)

    while ($sl.TimeLine.MainSequence.Count -gt 0) {
        $sl.TimeLine.MainSequence.Item(1).Delete()
    }

    # 1. Pro Slide Transition Choreography (32 slides)
    $tr = $sl.SlideShowTransition
    if ($sn -eq 1) {
        $tr.EntryEffect = $TR_FADE; $tr.Duration = 0.6
    } elseif ($sn -in @(3, 14, 28, 32)) {
        # Major narrative pivots: Embedded Reality, Enter Zephyr v4.5, Course Code Showdown, Final Verdict
        $tr.EntryEffect = $TR_PUSH_L; $tr.Duration = 0.72
    } elseif ($sn -in @(2, 7, 15, 17, 29, 30)) {
        # Structural reveals: Roadmap, Course Map, OS Canvas, Footprint Chart, Matrix, native_sim Lab
        $tr.EntryEffect = $TR_REVEAL; $tr.Duration = 0.60
    } else {
        $tr.EntryEffect = $TR_FADE; $tr.Duration = 0.48
    }

    $items = @()
    for ($i = 1; $i -le $sl.Shapes.Count; $i++) {
        $sh = $sl.Shapes.Item($i)
        $x = [float]$sh.Left / 72.0
        $y = [float]$sh.Top / 72.0
        $w = [float]$sh.Width / 72.0
        $h = [float]$sh.Height / 72.0
        $isChart = ($sh.HasChart -eq -1)
        $isTable = ($sh.HasTable -eq -1)
        $isPic   = ($sh.Type -eq 13) -or ($sh.Type -eq 14 -and $sh.HasTextFrame -eq 0 -and $sh.HasTable -eq 0 -and $sh.HasChart -eq 0)
        $txt = ""
        if ($sh.HasTextFrame -eq -1) {
            try { $txt = $sh.TextFrame.TextRange.Text.Trim() } catch {}
        }
        $isTopTitle   = ($sn -notin @(1, 14)) -and ($y -lt 0.92) -and ($w -gt 4.0) -and ($txt.Length -gt 0) -and (-not $isPic)
        $isLeadSub    = ($sn -notin @(1, 14)) -and ($y -ge 0.85) -and ($y -le 1.15) -and ($w -gt 7.0) -and ($h -lt 0.5)
        $isBottomPill = ($y -gt 4.40) -and ($w -gt 4.5) -and ($h -lt 0.85) -and (-not $isPic)

        $items += [PSCustomObject]@{
            Idx = $i; Shape = $sh; X = $x; Y = $y; W = $w; H = $h;
            IsChart = $isChart; IsTable = $isTable; IsPic = $isPic;
            IsTopTitle = $isTopTitle; IsLeadSub = $isLeadSub; IsBottomPill = $isBottomPill;
            Txt = $txt
        }
    }

    # 2. Slide-Specific Motion Choreography
    if ($sn -eq 1) {
        # HERO COVER: Smooth Ascend Title + Gentle Fade on Subtitle & Course/Team Metadata
        $titleSh = $items | Where-Object { $_.Y -lt 2.8 } | Select-Object -First 1
        $others  = $items | Where-Object { $_.Y -ge 2.8 } | Sort-Object Y
        if ($titleSh) { Add-Anim $sl $titleSh.Shape $AN_ASCEND 0.72 0.15 | Out-Null }
        $d = 0.42
        foreach ($it in $others) {
            Add-Anim $sl $it.Shape $AN_FADE 0.55 $d | Out-Null
            $d += 0.18
        }
    }
    elseif ($sn -eq 2) {
        # ROADMAP: 6 numbered columns cascade cleanly 01 -> 02 -> 03 -> 04 -> 05 -> 06
        $body = $items | Where-Object { -not $_.IsTopTitle }
        foreach ($it in $body) {
            $row = if ($it.Y -lt 3.05) { 0 } else { 1 }
            $col = if ($it.X -lt 3.4) { 0 } elseif ($it.X -lt 6.2) { 1 } else { 2 }
            $step = $row * 3 + $col
            $d = 0.12 + $step * 0.11
            $eff = if ($it.W -lt 0.4 -and $it.H -lt 0.3) { $AN_WIPE } else { $AN_FADE }
            $dir = if ($eff -eq $AN_WIPE) { $DIR_RIGHT } else { 0 }
            Add-Anim $sl $it.Shape $eff 0.38 $d $dir | Out-Null
        }
    }
    elseif ($sn -in @(3, 14, 30)) {
        # HERO HARDWARE PHOTO SLIDES (MCU Silicon, Zephyr Boards Pivot, native_sim Lab)
        $picIt = $items | Where-Object { $_.IsPic } | Select-Object -First 1
        if ($picIt) { Add-Anim $sl $picIt.Shape $AN_FZOOM 0.78 0.10 | Out-Null }
        $others = $items | Where-Object { -not $_.IsPic -and -not $_.IsTopTitle } | Sort-Object Y, X
        $d = 0.24
        foreach ($it in $others) {
            if ($it.IsBottomPill) {
                Add-Anim $sl $it.Shape $AN_WIPE 0.50 0.72 $DIR_RIGHT | Out-Null
            } else {
                Add-Anim $sl $it.Shape $AN_FADE 0.44 $d | Out-Null
                $d = [Math]::Min(0.68, $d + 0.08)
            }
        }
    }
    elseif ($sn -in @(10, 12, 18, 19)) {
        # LEFT-TO-RIGHT MULTI-STAGE PIPELINES (Context Switch, Mars Pathfinder, Unified Drivers, SYS_INIT)
        $body = $items | Where-Object { -not $_.IsTopTitle }
        foreach ($it in $body) {
            if ($it.IsLeadSub) {
                Add-Anim $sl $it.Shape $AN_FADE 0.32 0.08 | Out-Null
            } elseif ($it.IsBottomPill) {
                Add-Anim $sl $it.Shape $AN_WIPE 0.48 0.90 $DIR_RIGHT | Out-Null
            } elseif ($it.Y -lt 3.20) {
                $div = if ($sn -eq 19) { 1.70 } elseif ($sn -eq 18) { 2.75 } else { 2.10 }
                $step = [Math]::Min(4, [Math]::Max(0, [int][Math]::Floor(($it.X - 0.6) / $div)))
                Add-Anim $sl $it.Shape $AN_FADE 0.36 (0.14 + $step * 0.12) | Out-Null
            } else {
                $d = if ($it.X -lt 4.8) { 0.68 } else { 0.80 }
                Add-Anim $sl $it.Shape $AN_ASCEND 0.44 $d | Out-Null
            }
        }
    }
    elseif ($sn -in @(5, 22, 27)) {
        # CONCENTRIC RINGS + 3 RIGHT ICON CALLOUTS (Determinism, Interrupts IPSR, Automatic Power Management)
        $lefts  = $items | Where-Object { -not $_.IsTopTitle -and -not $_.IsLeadSub -and $_.X -lt 4.35 } | Sort-Object W -Descending
        $rights = $items | Where-Object { -not $_.IsTopTitle -and -not $_.IsLeadSub -and $_.X -ge 4.35 }
        foreach ($it in $items | Where-Object { $_.IsLeadSub }) {
            Add-Anim $sl $it.Shape $AN_FADE 0.32 0.08 | Out-Null
        }
        foreach ($it in $lefts) {
            $ring = if ($it.W -gt 2.8 -or $it.Y -lt 2.0) { 0 } elseif ($it.W -gt 1.9 -or $it.Y -lt 3.1) { 1 } else { 2 }
            Add-Anim $sl $it.Shape $AN_FZOOM 0.48 (0.12 + $ring * 0.15) | Out-Null
        }
        foreach ($it in $rights) {
            $band = if ($it.Y -lt 2.45) { 0 } elseif ($it.Y -lt 3.65) { 1 } else { 2 }
            Add-Anim $sl $it.Shape $AN_ASCEND 0.44 (0.45 + $band * 0.16) | Out-Null
        }
    }
    elseif ($sn -eq 16) {
        # HIERARCHICAL TREE (Devicetree 3-Layer Cascade)
        $body = $items | Where-Object { -not $_.IsTopTitle }
        foreach ($it in $body) {
            if ($it.IsBottomPill) {
                Add-Anim $sl $it.Shape $AN_WIPE 0.50 0.88 $DIR_RIGHT | Out-Null
            } else {
                $lvl = if ($it.Y -lt 1.85) { 0 } elseif ($it.Y -lt 2.85) { 1 } elseif ($it.Y -lt 3.85) { 2 } else { 3 }
                Add-Anim $sl $it.Shape $AN_FADE 0.40 (0.12 + $lvl * 0.17) | Out-Null
            }
        }
    }
    elseif ($sn -eq 17) {
        # FOOTPRINT CHART SLIDE
        $body = $items | Where-Object { -not $_.IsTopTitle }
        foreach ($it in $body) {
            if ($it.IsChart) {
                Add-Anim $sl $it.Shape $AN_WIPE 0.72 0.48 $DIR_UP | Out-Null
            } elseif ($it.X -lt 3.95) {
                $band = if ($it.Y -lt 2.1) { 0 } elseif ($it.Y -lt 2.95) { 1 } elseif ($it.Y -lt 3.8) { 2 } else { 3 }
                Add-Anim $sl $it.Shape $AN_ASCEND 0.38 (0.10 + $band * 0.12) | Out-Null
            } else {
                Add-Anim $sl $it.Shape $AN_FADE 0.40 0.82 | Out-Null
            }
        }
    }
    elseif ($sn -in @(28, 32)) {
        # DUAL HERO SHOWDOWN (Course HVAC Upgrade & Final Engineering Verdict + Q&A)
        $body = $items | Where-Object { -not $_.IsTopTitle }
        foreach ($it in $body) {
            if ($it.IsBottomPill) {
                Add-Anim $sl $it.Shape $AN_WIPE 0.52 0.76 $DIR_RIGHT | Out-Null
            } elseif ($it.X -lt 4.75) {
                Add-Anim $sl $it.Shape $AN_ASCEND 0.52 0.14 | Out-Null
            } else {
                Add-Anim $sl $it.Shape $AN_ASCEND 0.52 0.42 | Out-Null
            }
        }
    }
    else {
        # ALL REMAINING CONTENT SLIDES
        $body = $items | Where-Object { -not $_.IsTopTitle }
        foreach ($it in $body) {
            if ($it.IsLeadSub) {
                Add-Anim $sl $it.Shape $AN_FADE 0.30 0.08 | Out-Null
            } elseif ($it.IsBottomPill) {
                Add-Anim $sl $it.Shape $AN_WIPE 0.50 0.74 $DIR_RIGHT | Out-Null
            } elseif ($it.IsTable) {
                Add-Anim $sl $it.Shape $AN_FADE 0.48 0.14 | Out-Null
            } else {
                $col = if ($it.X -lt 3.2) { 0 } elseif ($it.X -lt 6.0) { 1 } else { 2 }
                $row = if ($it.Y -lt 2.8) { 0 } else { 1 }
                $step = $row * 3 + $col
                Add-Anim $sl $it.Shape $AN_FADE 0.40 ([Math]::Min(0.65, 0.14 + $step * 0.09)) | Out-Null
            }
        }
    }

    $totalEffects += $sl.TimeLine.MainSequence.Count
}

$pres.Save()
$sc = $pres.Slides.Count
$pres.Close()
if ($pp.Presentations.Count -eq 0) {
    $pp.Quit()
}
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pp) | Out-Null

Write-Host "ANIMATED_OK: FILE=$TargetPptx SLIDES=$sc TOTAL_ANIM_EFFECTS=$totalEffects SIZE=$((Get-Item $TargetPptx).Length)"
