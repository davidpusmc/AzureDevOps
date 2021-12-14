function Convert-PSObjectToHashtable
{
    <#
    .SYNOPSIS
    Converts a PSObject to a Hash Table.

    .DESCRIPTION
    This was implementation found at:
    https://stackoverflow.com/questions/3740128/pscustomobject-to-hashtable
    https://stackoverflow.com/a/34383413

    .EXAMPLE
    $myHashTable = $myPSObject | Convert-PSObjectToHashtable
    #>
    param (
        [Parameter(ValueFromPipeline)]
        $InputObject
    )

    process
    {
        if ($null -eq $InputObject) { return Write-Output $null }

        if ($InputObject -is [System.Collections.IEnumerable] -and $InputObject -isnot [string])
        {
            $collection = @(
                foreach ($object in $InputObject) { Convert-PSObjectToHashtable $object }
            )

            Write-Output -NoEnumerate $collection
        }
        elseif ($InputObject -is [psobject])
        {
            $hash = @{}

            foreach ($property in $InputObject.PSObject.Properties)
            {
                $hash[$property.Name] = Convert-PSObjectToHashtable $property.Value
            }

            Write-Output $hash
        }
        else
        {
            Write-Output $InputObject
        }
    }
}
