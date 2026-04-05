# 빌드: .\build.ps1  |  정식: .\build.ps1 -Formal
param([switch]$Formal)
$ErrorActionPreference = "Stop"
& "$PSScriptRoot\scripts\build.ps1" -Formal:$Formal
