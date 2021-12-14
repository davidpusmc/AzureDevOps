$errorActionPreference="Stop"
Import-Module -Name .\Invoke-StopOnError.psm1 -Force

Invoke-StopOnError git --version
try { Invoke-ThrowOnError git fail } catch { Write-Warning $_ }
try { Invoke-ThrowOnError NotACommand } catch { Write-Warning $_ }
