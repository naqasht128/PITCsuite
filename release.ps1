# PowerShell helper for release.py
param (
    [string]$Version = "",
    [string]$Message = ""
)

$argsList = @()
if ($Version) { $argsList += $Version }
if ($Message) { $argsList += "-m", $Message }

python release.py @argsList
