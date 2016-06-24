$path = $env:LOCALAPPDATA + "\Funcom\TSW"
$backupIntervalDays = 2
$backupCount = 10

$currentBackups = Get-ChildItem $path | Where-Object Name -match '^Prefs \d{8}$'

# parse date of last backup
if (($currentBackups.length -gt 0) -and ($currentBackups[-1].Name -match '^Prefs (\d{8})$')) {
	$lastBackupDate = New-Object DateTime
	[void] ([DateTime]::TryParseExact(
		$matches[1],
		"yyyyMMdd",
		[System.Globalization.CultureInfo]::InvariantCulture,
		[System.Globalization.DateTimeStyles]::None,
		[ref]$lastBackupDate))
} else {
	$lastBackupDate = (Get-Date 2000-01-01)
}

# create new backup
if ($lastBackupDate.AddDays($backupIntervalDays) -lt (Get-Date)) {
	Copy-Item -Path ($path + "\Prefs") -Destination ($path + "\Prefs " + (Get-Date -format yyyyMMdd)) -Recurse
	$currentBackups = Get-ChildItem $path | Where-Object Name -match '^Prefs \d{8}$'
}

# delete old backups
while ($currentBackups.length -gt $backupCount) {
	$currentBackups[0] | Remove-Item -Recurse
	$currentBackups = Get-ChildItem $path | Where-Object Name -match '^Prefs \d{8}$'
}
