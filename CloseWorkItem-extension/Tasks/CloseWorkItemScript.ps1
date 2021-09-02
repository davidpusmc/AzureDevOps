
#These variables are defined by the variables within the environment.
$AzureDevOpsPAT = $Env:SYSTEM_ACCESSTOKEN
$BuildID = $Env:BUILD_BUILDID
$ProjectName = $Env:BUILD_PROJECTNAME
$UriOrga = $Env:SYSTEM_TEAMFOUNDATIONCOLLECTIONURI
#####################################Process###################################################################

$AzureDevOpsAuthenicationHeader = @{Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($AzureDevOpsPAT)")) }

$BuildUri = $UriOrga + "$ProjectName/_apis/build/builds/$BuildID/workitems?$top=%7B$top%7D&api-version=6.0" #URI for RequestBuildWorkItems

#RequestBuildWorkItems used to request work items from build
$RequestBuildWorkItems = (Invoke-RestMethod -Uri $BuildUri -Method Get -Headers $AzureDevOpsAuthenicationHeader -ContentType "application/json-patch+json")

#iterates through each work Item within the build where $id is the WorkItem ID within RequestBuildWorkItems.value.id
foreach($id in $RequestBuildWorkItems.value.id)
{
    $ItemUpdateURL = "$UriOrga/_apis/wit/workItems/$($id)?api-version=6.0" #uri for workItems
    #request specific workItem to be modified
    $StateChange = (Invoke-RestMethod -Uri $ItemUpdateURL -Method Get -Headers $AzureDevOpsAuthenicationHeader -ContentType "application/json-patch+json")
    
    #Checks that each work item requested has a type of bug or Product Backlog Item and that the from state is correct before updating with new state
    if (($StateChange.fields.'System.WorkItemType' -eq "Bug" -or $StateChange.fields.'System.WorkItemType' -eq "Product Backlog Item"))#must match System.State value exactly
    {
        if($StateChange.fields.'System.State' -eq "Committed")
        {
            $newValue = "Done"
        }
        elseif($StateChange.fields.'System.State' -eq "Resolved")
        {
            $newValue = "Closed"
        }
        $jsonData = ConvertTo-Json @(@{ 
            op = "add"
            path = "/fields/System.State"
            value = $newValue
        })#which key/s is being updated and with what value
        #sends patch after parameters are met
        $request = Invoke-RestMethod -Uri $ItemUpdateURL -Method Patch -Body ($jsonData) -Headers $AzureDevOpsAuthenicationHeader -ContentType "application/json-patch+json"
    }
    #else statement handles any instance other than specific parameters
    else
    {
        Write-Warning "`r`nThe type of the work Items must be either 'Bug' or 'Product Backlog Item' and`r`n
        The previous state must be either 'Committed' or 'Resolved' to process this request."
    }
}



