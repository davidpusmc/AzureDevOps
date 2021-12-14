$ErrorActionPreference='Stop'

# Get public and private function definition files.
$Public  = @( Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue )
$Private = @( Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue )

# Dot source the files.
Foreach($import in @($Public + $Private)) {
  Try {
    . $import.fullname
  }
  Catch {
    Write-Error -Message "Failed to import function $($import.fullname): $_"
  }
}


# Load configuration file.
# ...

# Set variables visible to the module and its functions only
# ...

# Exporting "$Public.Basename" works when the module is imported, but causes PowerShell 
# not to automatically import the module when exports are first used.
# This requires functions names to be listed in the manifest ".psd1" file to be auto loaded.
# To work without a manifest, functions can be exported by writing the name explicitly here, 
# for example: "Export-ModuleMember -Function 'Verb-Noun', 'Verb-Noun'"
Export-ModuleMember -Function $Public.Basename
