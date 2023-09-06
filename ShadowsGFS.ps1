param (
    [String]$Drive = $(Read-Host -Prompt 'Drive'),
    [Int]$KeepLast = 4,
    [Int]$KeepHourly = 24,
    [Int]$KeepDaily = 7,
    [Int]$KeepWeekly = 4,
    [Int]$KeepMonthly = 12,
    [Int]$KeepYearly = 10,
    [switch]$WhatIf
)

if ($Drive.Length -eq 1) { $Drive = $Drive + ":" }
if ($Drive -notmatch "^[a-z]:$") { throw "Invalid drive `"$Drive`"" }

# Build list of shadow copies
[array] $output = vssadmin list shadows /for=$Drive
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
    }
}
if ($shadow.id) {
    $shadows.Add($shadow) | Out-Null
}

# Keep last x copies
for ($x = -$KeepLast; $x -lt 0; $x++) {
    if ($shadows[$x]) {
        $shadows[$x].delete = $false
        $shadows[$x].why += "L"
    }
}

# Keep last x hourly copies
for ($x = 0; $x -gt -$KeepHourly; $x--) {
    $now = Get-Date
    $from = (Get-Date -Year $now.Year -Month $now.Month -Day $now.Day -Hour $now.Hour -Minute 0 -Second 0).AddHours($x)
    $to = $from.AddHours(1)
    foreach ($s in $shadows) {
        if ($from -le $s.datetime -and $s.datetime -lt $to) {
            $s.delete = $false
            $s.why += "H"
            break
        }
    }
}

# Keep last x daily copies
for ($x = 0; $x -gt -$KeepDaily; $x--) {
    $now = Get-Date
    $from = (Get-Date -Year $now.Year -Month $now.Month -Day $now.Day -Hour 0 -Minute 0 -Second 0).AddDays($x)
    $to = $from.AddDays(1)
    foreach ($s in $shadows) {
        if ($from -le $s.datetime -and $s.datetime -lt $to) {
            $s.delete = $false
            $s.why += "D"
            break
        }
    }
}

# Keep last x weekly copies
for ($x = 0; $x -gt -$KeepWeekly; $x--) {
    $now = Get-Date
    $from = (Get-Date -Year $now.Year -Month $now.Month -Day $now.Day -Hour 0 -Minute 0 -Second 0).AddDays(-$now.DayOfWeek.value__+1).AddDays($x*7)
    $to = $from.AddDays(7)
    foreach ($s in $shadows) {
        if ($from -le $s.datetime -and $s.datetime -lt $to) {
            $s.delete = $false
            $s.why += "W"
            break
        }
    }
}

# Keep last x monthly copies
for ($x = 0; $x -gt -$KeepMonthly; $x--) {
    $now = Get-Date
    $from = (Get-Date -Year $now.Year -Month $now.Month -Day 1 -Hour 0 -Minute 0 -Second 0).AddMonths($x)
    $to = $from.AddMonths(1)
    foreach ($s in $shadows) {
        if ($from -le $s.datetime -and $s.datetime -lt $to) {
            $s.delete = $false
            $s.why += "M"
            break
        }
    }
}

# Keep last x yearly copies
for ($x = 0; $x -gt -$KeepYearly; $x--) {
    $now = Get-Date
    $from = (Get-Date -Year $now.Year -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0).AddYears($x)
    $to = $from.AddYears(1)
    foreach ($s in $shadows) {
        if ($from -le $s.datetime -and $s.datetime -lt $to) {
            $s.delete = $false
            $s.why += "Y"
            break
        }
    }
}

# Delete shadows
if ($WhatIf) {
    foreach ($s in $shadows) {
        "$($s.id) @ $($s.datetime.ToShortDateString()) $($s.datetime.ToShorttimeString()) -> Delete: $($s.delete) ($($s.why))"
    }
}
else {
    foreach ($s in $shadows) {
        if ($s.delete) {
            "Deleting shadow copy from $($s.datetime.ToShortDateString()) $($s.datetime.ToShorttimeString())"
            vssadmin delete shadows /shadow=$($s.id) /quiet | Out-Null
        }
    }
}
