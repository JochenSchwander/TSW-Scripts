Param(
    [Parameter(Mandatory = $true)][string]$accountName,
    [Parameter(Mandatory = $true)][string]$loreFile
)

function Get-InstanceCounter($content) {
    [int]$instanceCounter = 0
    try {
        [int]$instanceCounter = [convert]::ToInt32(((Select-String -Pattern "instance=`"(\d+)`"" -InputObject $content -AllMatches).matches[-1].groups[1].value), 10)
    } catch {
        [int]$instanceCounter = 0
    }
    return ($instanceCounter + 1)
}

function Escape-XMLTags($text) {
    $text = $text -replace "&", "&amp;"
    $text = $text -replace "`"", "&quot;"
    $text = $text -replace "'", "&apos;"
    $text = $text -replace "<", "&lt;"
    $text = $text -replace ">", "&gt;"
    return $text
}

function Write-Waypoint($content, $title, $tooltip, $x, $y) {
    [int]$instanceCounter = Get-InstanceCounter $content
    $title = Escape-XMLTags $title
    $tooltip = Escape-XMLTags $tooltip
    $waypoint = "    <Waypoint type=`"custom`" instance=`"$instanceCounter`" name=`"$title`" tooltip_title=`"$title`" tooltip_text=`"$tooltip`" position=`"Point($x,$y)`" distributed_value=`"ShowCustomWaypoints`" />`r`n</Root>"
    return $content -replace "</Root>", $waypoint
}

function Create-Waypoint($zone, $title, $tooltip, $x, $y) {
    switch -regex ($zone) {
        "^London$" {
            $global:london = Write-Waypoint $global:london $title $tooltip $x $y
        }
        "^New York$" {
            $global:newYork = Write-Waypoint $global:newYork $title $tooltip $x $y
        }
        "^Seoul$" {
            $global:seoul = Write-Waypoint $global:seoul $title $tooltip $x $y
        }
        "^Kingsmouth$" {
            $global:kingsmouth = Write-Waypoint $global:kingsmouth $title $tooltip $x $y
        }
        "^(?:The )?Savage Coast$" {
            $global:savageCoast = Write-Waypoint $global:savageCoast $title $tooltip $x $y
        }
        "^(?:The )?Blue Mountains$" {
            $global:blueMountains = Write-Waypoint $global:blueMountains $title $tooltip $x $y
        }
        "^Kaidan$" {
            $global:kaidan = Write-Waypoint $global:kaidan $title $tooltip $x $y
        }
        "^(?:The )?Scorched Desert$" {
            $global:scorchedDesert = Write-Waypoint $global:scorchedDesert $title $tooltip $x $y
        }
        "^(?:The )?City of the Sun God$" {
            $global:cityOfTheSunGod = Write-Waypoint $global:cityOfTheSunGod $title $tooltip $x $y
        }
        "^(?:The )?Besieged Farmlands$" {
            $global:besiegedFarmlands = Write-Waypoint $global:besiegedFarmlands $title $tooltip $x $y
        }
        "^(?:The )?Shadowy Forest$" {
            $global:shadowyForest = Write-Waypoint $global:shadowyForest $title $tooltip $x $y
        }
        "^(?:The )?Carpathian Fangs$" {
            $global:carpathianFangs = Write-Waypoint $global:carpathianFangs $title $tooltip $x $y
        }
        "^(?:The )?Slaughterhouse$" {
            $global:slaughterhouse = Write-Waypoint $global:slaughterhouse $title $tooltip $x $y
        }
        "^(?:The )?Polaris$" {
            $global:polaris = Write-Waypoint $global:polaris $title $tooltip $x $y
        }
        "^Agartha$" {
            $global:agartha = Write-Waypoint $global:Agartha $title $tooltip $x $y
        }
        "^(?:The )?Ankh$" {
            $global:ankh = Write-Waypoint $global:ankh $title $tooltip $x $y
        }
        "^Hell Raised$" {
            $global:hellRaised = Write-Waypoint $global:hellRaised $title $tooltip $x $y
        }
        "^Hell Fallen$" {
            $global:hellFallen = Write-Waypoint $global:hellFallen $title $tooltip $x $y
        }
        "^Hell Eternal$" {
            $global:hellEternal = Write-Waypoint $global:hellEternal $title $tooltip $x $y
        }
        "^(?:The )?Darkness Wars$" {
            $global:darknessWars = Write-Waypoint $global:darknessWars $title $tooltip $x $y
        }
        "^(?:The )?Facility$" {
            $global:facility = Write-Waypoint $global:facility $title $tooltip $x $y
        }
        "^(?:The )?Manhattan Exclusion Zone$" {
            $global:manhattanExclusionZone = Write-Waypoint $global:manhattanExclusionZone $title $tooltip $x $y
        }
        "^N'gha-Pei the Corpse-Island$" {
            $global:corpseIsland = Write-Waypoint $global:corpseIsland $title $tooltip $x $y
        }
        "^Agartha Defiled$" {
            $global:agarthaDefiled = Write-Waypoint $global:agarthaDefiled $title $tooltip $x $y
        }
        "^(?:The )?Manufactory$" {
            $global:manufactory = Write-Waypoint $global:manufactory $title $tooltip $x $y
        }
        "^(?:The )?Manufactory:? Breached$" {
            $global:manufactoryBreached = Write-Waypoint $global:manufactoryBreached $title $tooltip $x $y
        }
        default {
            Write-Host "No map marker for lore entry `"$title`" created, unknown zone `"$zone`"."
        }
    }
}

function Get-WaypointFileContent($fileName) {
    [string]$content = ""
    try {
        $content = Get-Content $fileName -Raw -ErrorAction Stop
    } catch {
        $content = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`" ?>`r`n<Root>`r`n</Root>`r`n"
    }
    $content = $content -replace "<Root />", "<Root>`r`n</Root>"
    return $content
}

function Write-WaypointFile($fileName, $content) {
    $content = $content -replace "(?:`r`n){2,100}", ""
    $content = $content -replace "<Root>`r`n</Root>", "<Root />"
    [IO.File]::WriteAllLines($fileName, $content)
}

# try to read lore file
try {
    $loreFileContent = Get-Content $loreFile -Raw -ErrorAction Stop
} catch {
    Write-Host "Invalid lore file `"$loreFile`"."
    return
}

# TSW waypoints folder
$path = $env:LOCALAPPDATA + "\Funcom\TSW\Prefs\" + $accountName + "\Waypoints\"
if(-not(Test-Path $path)) {
    Write-Host "Invalid account name `"$accountName`""
    return
}

# Waypoint file names
$londonFileName = $path + "PF1000.xml"
$newYorkFileName = $path + "PF1100.xml"
$seoulFileName = $path + "PF1200.xml"
$kingsmouthFileName = $path + "PF3030.xml"
$savageCoastFileName = $path + "PF3040.xml"
$blueMountainsFileName = $path + "PF3050.xml"
$kaidanFileName = $path + "PF3070.xml"
$scorchedDesertFileName = $path + "PF3090.xml"
$cityOfTheSunGodFileName = $path + "PF3100.xml"
$besiegedFarmlandsFileName = $path + "PF3120.xml"
$shadowyForestFileName = $path + "PF3130.xml"
$carpathianFangsFileName = $path + "PF3140.xml"
$slaughterhouseFileName = $path + "PF5000.xml"
$polarisFileName = $path + "PF5040.xml"
$agarthaFileName = $path + "PF5060.xml"
$ankhFileName = $path + "PF5080.xml"
$hellRaisedFileName = $path + "PF5140.xml"
$hellFallenFileName = $path + "PF5150.xml"
$hellEternalFileName = $path + "PF5160.xml"
$darknessWarsFileName = $path + "PF5170.xml"
$facilityFileName = $path + "PF5190.xml"
$manhattanExclusionZoneName = $path + "PF5710.xml"
$corpseIslandFileName = $path + "PF5720.xml"
$agarthaDefiledFileName = $path + "PF5730.xml"
$manufactoryFileName = $path + "PF6900.xml"
$manufactoryBreachedFileName = $path + "PF6910.xml"

# Waypoint files
$global:london = Get-WaypointFileContent $londonFileName
$global:newYork = Get-WaypointFileContent $newYorkFileName
$global:seoul = Get-WaypointFileContent $seoulFileName
$global:kingsmouth = Get-WaypointFileContent $kingsmouthFileName
$global:savageCoast = Get-WaypointFileContent $savageCoastFileName
$global:blueMountains = Get-WaypointFileContent $blueMountainsFileName
$global:kaidan = Get-WaypointFileContent $kaidanFileName
$global:scorchedDesert = Get-WaypointFileContent $scorchedDesertFileName
$global:cityOfTheSunGod = Get-WaypointFileContent $cityOfTheSunGodFileName
$global:besiegedFarmlands = Get-WaypointFileContent $besiegedFarmlandsFileName
$global:shadowyForest = Get-WaypointFileContent $shadowyForestFileName
$global:carpathianFangs = Get-WaypointFileContent $carpathianFangsFileName
$global:slaughterhouse = Get-WaypointFileContent $slaughterhouseFileName
$global:polaris = Get-WaypointFileContent $polarisFileName
$global:agartha = Get-WaypointFileContent $agarthaFileName
$global:ankh = Get-WaypointFileContent $ankhFileName
$global:hellRaised = Get-WaypointFileContent $hellRaisedFileName
$global:hellFallen = Get-WaypointFileContent $hellFallenFileName
$global:hellEternal = Get-WaypointFileContent $hellEternalFileName
$global:darknessWars = Get-WaypointFileContent $darknessWarsFileName
$global:facility = Get-WaypointFileContent $facilityFileName
$global:manhattanExclusionZone = Get-WaypointFileContent $manhattanExclusionZoneName
$global:corpseIsland = Get-WaypointFileContent $corpseIslandFileName
$global:agarthaDefiled = Get-WaypointFileContent $agarthaDefiledFileName
$global:manufactory = Get-WaypointFileContent $manufactoryFileName
$global:manufactoryBreached = Get-WaypointFileContent $manufactoryBreachedFileName

# lore category, lore number, zone, x, y, comment
$lores = (Select-String -Pattern "(.+?),([\d\s]+?),(.+?),`"?([\d\s]+),([\d\s]+)`"?,?(.*)" -InputObject $loreFileContent -AllMatches).matches
ForEach($lore in $lores) {
    Create-Waypoint $lore.groups[3].value.Trim() ($lore.groups[1].value.Trim() + " " + $lore.groups[2].value.Trim()) $lore.groups[6].value.Trim() $lore.groups[4].value.Trim() $lore.groups[5].value.Trim()
}

Write-WaypointFile $londonFileName $global:london
Write-WaypointFile $newYorkFileName $global:newYork
Write-WaypointFile $seoulFileName $global:seoul
Write-WaypointFile $kingsmouthFileName $global:kingsmouth
Write-WaypointFile $savageCoastFileName $global:savageCoast
Write-WaypointFile $blueMountainsFileName $global:blueMountains
Write-WaypointFile $kaidanFileName $global:kaidan
Write-WaypointFile $scorchedDesertFileName $global:scorchedDesert
Write-WaypointFile $cityOfTheSunGodFileName $global:cityOfTheSunGod
Write-WaypointFile $besiegedFarmlandsFileName $global:besiegedFarmlands
Write-WaypointFile $shadowyForestFileName $global:shadowyForest
Write-WaypointFile $carpathianFangsFileName $global:carpathianFangs
Write-WaypointFile $slaughterhouseFileName $global:slaughterhouse
Write-WaypointFile $polarisFileName $global:polaris
Write-WaypointFile $agarthaFileName $global:agartha
Write-WaypointFile $ankhFileName $global:ankh
Write-WaypointFile $hellRaisedFileName $global:hellRaised
Write-WaypointFile $hellFallenFileName $global:hellFallen
Write-WaypointFile $hellEternalFileName $global:hellEternal
Write-WaypointFile $darknessWarsFileName $global:darknessWars
Write-WaypointFile $facilityFileName $global:facility
Write-WaypointFile $manhattanExclusionZoneName $global:manhattanExclusionZone
Write-WaypointFile $corpseIslandFileName $global:corpseIsland
Write-WaypointFile $agarthaDefiledFileName $global:agarthaDefiled
Write-WaypointFile $manufactoryFileName $global:manufactory
Write-WaypointFile $manufactoryBreachedFileName $global:manufactoryBreached

$failedLines = $loreFileContent -replace ".+?,[\d\s]+?,.+?,`"?[\d\s]+,[\d\s]+`"?,?.*\n?", ""
if (-not(($failedLines -replace "`r`n", "") -eq "")) {
    Write-Host "Could not parse the following lines:`r`n$failedLines"
}
