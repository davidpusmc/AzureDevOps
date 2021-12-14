# Load modules
Import-Module -Force "$PSScriptRoot\Modules\Common\Common.psd1"
Import-Module "$PSScriptRoot\SendReleaseAgenda.psm1" -Force -Verbose

# Invoke function.
SendReleaseAgenda
