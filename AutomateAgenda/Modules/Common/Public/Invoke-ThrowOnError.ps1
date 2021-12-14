function Invoke-ThrowOnError {
    <#
    .SYNOPSIS
    Invokes an external utility, ensuring successful execution.

    .DESCRIPTION
    Invokes an external utility (program) and, if the utility indicates failure by 
    way of a nonzero exit code, throws a script-terminating error.

    * Pass the command the way you would execute the command directly.
    * Do NOT use & as the first argument if the executable name is not a literal.

    This was implementation found at:
    https://stackoverflow.com/questions/48864988/powershell-with-git-command-error-handling-automatically-abort-on-non-zero-exi

    .EXAMPLE
    Invoke-Utility git push

    Executes `git push` and throws a script-terminating error if the exit code
    is nonzero.
    #>

    # Set ErrorActionPreference to 'Continue' to prevent a bug where stderr redirections "2>" applied to calls to this function from accidentally triggering a terminating error.
    # See bug reports at:
    # https://github.com/PowerShell/PowerShell/issues/4002
    # https://github.com/PowerShell/PowerShell/issues/3996
    $ErrorActionPreference = 'Continue'
    $exe, $argsForExe = $Args
    & $exe $argsForExe
    if ($LASTEXITCODE) { Throw "$exe indicated failure (exit code $LASTEXITCODE; full command: $Args)." }
}
