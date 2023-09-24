param (
    [string]$Drive = $(Read-Host -Prompt 'Drive'),
    [int]$KeepLast = 4,
    [int]$KeepHourly = 24,
    [int]$KeepDaily = 7,
    [int]$KeepWeekly = 4,
    [int]$KeepMonthly = 12,
    [int]$KeepYearly = 10,
    [switch]$WhatIf
)

function Main {
    $Drive = Format-DriveLetter $Drive
    $shadows = Get-VssAdminShadows $Drive

    # Keep last x copies
    for ($x = - $KeepLast; $x -lt 0; $x++) {
        if ($shadows[$x]) {
            $shadows[$x].delete = $false
            $shadows[$x].why.Add("Last") | Out-Null
        }
    }

    # Keep last x hourly copies
    for ($x = 0; $x -gt - $KeepHourly; $x--) {
        $now = Get-Date
        $from = (Get-Date -Year $now.Year -Month $now.Month -Day $now.Day -Hour $now.Hour -Minute 0 -Second 0).AddHours($x)
        $to = $from.AddHours(1)
        foreach ($s in $shadows) {
            if ($from -le $s.datetime -and $s.datetime -lt $to) {
                $s.delete = $false
                $s.why.Add("Hourly") | Out-Null
                break
            }
        }
    }

    # Keep last x daily copies
    for ($x = 0; $x -gt - $KeepDaily; $x--) {
        $now = Get-Date
        $from = (Get-Date -Year $now.Year -Month $now.Month -Day $now.Day -Hour 0 -Minute 0 -Second 0).AddDays($x)
        $to = $from.AddDays(1)
        foreach ($s in $shadows) {
            if ($from -le $s.datetime -and $s.datetime -lt $to) {
                $s.delete = $false
                $s.why.Add("Daily") | Out-Null
                break
            }
        }
    }

    # Keep last x weekly copies
    for ($x = 0; $x -gt - $KeepWeekly; $x--) {
        $now = Get-Date
        $from = (Get-Date -Year $now.Year -Month $now.Month -Day $now.Day -Hour 0 -Minute 0 -Second 0).AddDays(-$now.DayOfWeek.value__ + 1).AddDays($x * 7)
        $to = $from.AddDays(7)
        foreach ($s in $shadows) {
            if ($from -le $s.datetime -and $s.datetime -lt $to) {
                $s.delete = $false
                $s.why.Add("Weekly") | Out-Null
                break
            }
        }
    }

    # Keep last x monthly copies
    for ($x = 0; $x -gt - $KeepMonthly; $x--) {
        $now = Get-Date
        $from = (Get-Date -Year $now.Year -Month $now.Month -Day 1 -Hour 0 -Minute 0 -Second 0).AddMonths($x)
        $to = $from.AddMonths(1)
        foreach ($s in $shadows) {
            if ($from -le $s.datetime -and $s.datetime -lt $to) {
                $s.delete = $false
                $s.why.Add("Monthly") | Out-Null
                break
            }
        }
    }

    # Keep last x yearly copies
    for ($x = 0; $x -gt - $KeepYearly; $x--) {
        $now = Get-Date
        $from = (Get-Date -Year $now.Year -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0).AddYears($x)
        $to = $from.AddYears(1)
        foreach ($s in $shadows) {
            if ($from -le $s.datetime -and $s.datetime -lt $to) {
                $s.delete = $false
                $s.why.Add("Yearly") | Out-Null
                break
            }
        }
    }

    # What if
    $deleted = 0
    if ($WhatIf) {
        foreach ($s in $shadows) {
            if ($s.why) { $why = "[$($s.why -join ", ")]" } else { $why = "" }
            if ($s.delete) { $action = "Delete" } else { $action = "Keep $why" }
            "$Drive @ $(DTFormat $s.datetime) -> $action"
        }
    }

    # Delete shadow copies
    else {
        foreach ($s in $shadows) {
            if ($s.delete) {
                Log "Deleting shadow copy from $(DTFormat)" Yellow
                vssadmin delete shadows /shadow=$($s.id) /quiet | Out-Null
                $deleted++
            }
        }
    }

    # Show info
    $storage = Get-VssAdminStorage $Drive
    Log ("Drive {0} has {1} shadow copies using {2} ({3}) of space" -f $Drive, ($shadows.Count - $deleted), $storage.Size, $storage.Percentage) Green
}

function Format-DriveLetter ([string]$Drive) {
    if ($Drive.Length -eq 1) { $Drive = $Drive + ":" }
    $Drive = $Drive.ToUpper()
    if ($Drive -cnotmatch "^[A-Z]:$") { throw "Invalid drive `"$Drive`"" }
    return $Drive
}

function Get-VssAdminShadows ([string]$Drive) {
    [array]$output = vssadmin list shadows /for=$Drive
    if (-not $?) {
        Log $output[3].Split(".")[0] Red
        exit
    }

    $shadows = [System.Collections.ArrayList]@()
    $shadow = @{}
    foreach ($line in $output) {
        $line = $line.Trim()
        if ($line -clike "*shadow copies at creation time:*") {
            $datetime = [DateTime]::ParseExact($line.Split()[7..8] -join " ", "d-M-yyyy HH:mm:ss", $null)
        }
        elseif ($line -clike "Shadow Copy ID:*") {
            if ($shadow.id) {
                $shadows.Add($shadow) | Out-Null
                $shadow = @{}
            }
            $id = $line.Split()[3]
            $shadow.id = $id
            $shadow.datetime = $datetime
            $shadow.delete = $true
            $shadow.why = [System.Collections.ArrayList]@()
        }
    }
    if ($shadow.id) {
        $shadows.Add($shadow) | Out-Null
    }

    return $shadows
}

function Get-VssAdminStorage ([string]$Drive) {
    [array]$output = vssadmin list shadowstorage /on=$Drive
    $line = ($output | Where-Object { $_ -like "*Allocated Shadow Copy Storage space:*" }).Trim()
    $size = ($line.Split()[5..6] -join " ").replace(",", ".")
    $percentage = $line.Split()[7].replace("(", "").replace(")", "")
    return @{"Size" = $size; "Percentage" = $percentage }
}

function Log ($Message, [System.ConsoleColor]$ForegroundColor = 7 ) {
    Write-Host $Message -ForegroundColor $ForegroundColor
    "$(DTFormat) $Message" | Out-File $PSCommandPath.Replace(".ps1", ".log") -Append
}

function DTFormat ([datetime]$datetime = $(Get-Date)) {
    Get-Date -Date $datetime -Format "yyyy-MM-dd HH:mm"
}

Main
