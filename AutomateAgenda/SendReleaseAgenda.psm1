#Requires -Modules Common
# using namespace System.IO

# Define ErrorActionPreference for all functions in this module.
$ErrorActionPreference="Stop"

$QueryDM=@"
SELECT * FROM WorkItems
WHERE
   [System.WorkItemType] In ( `"User Story`",`"Product Backlog Item`",`"Bug`" )
   AND (
      [Microsoft.VSTS.Common.DD]=`"{{ReleaseDate}}`"
      OR [Microsoft.VSTS.Common.ImpDate]=`"{{ReleaseDate}}`"
   )
   AND (
      [System.AssignedTo] In Group `"[Data_Management]\Big_Efforts`"
      OR [System.AssignedTo] In Group `"[Data_Management]\Implementations`"
   )
   AND [System.AreaPath] Not Under `"AS400\Firefighters Kanban`"
"@
$QueryRelease=@"
SELECT * FROM WorkItems
WHERE
   [System.WorkItemType]=`"Release`"
   AND (
      [Microsoft.VSTS.Common.DD]=`"{{ReleaseDate}}`"
      OR [Microsoft.VSTS.Common.ImpDate]=`"{{ReleaseDate}}`"
   )
"@
$QueryAS400=@"
SELECT * FROM WorkItems
WHERE
   [System.WorkItemType] in (`"Product Backlog Item`", `"Bug`")
   AND (
      [Microsoft.VSTS.Common.DD]= `"{{ReleaseDate}}`"
      OR [Microsoft.VSTS.Common.ImpDate]= `"{{ReleaseDate}}`"
      )
   AND 
      [SWBC.ApplicationType] in (`"Auto`", `"Gap`", `"Mortgage`", `"Stone Eagle`", `"Accounting`")
   AND
      [System.State]= `"Ready for Deployment`"
   AND
      [SWBC.AS400Task]!= `"None`"
   AND 
      [System.Tags] Not Contains `"ETL`"
"@

function SendReleaseAgenda() {
   [cmdletbinding(PositionalBinding=$False, SupportsShouldProcess)]
   Param(
      [string]
      $Org = $(
         if (Get-Variable SYSTEM_TEAMFOUNDATIONCOLLECTIONURI -ErrorAction 'Continue' 2> $null) {$SYSTEM_TEAMFOUNDATIONCOLLECTIONURI}
         elseif (Test-Path Env:SYSTEM_TEAMFOUNDATIONCOLLECTIONURI) {$Env:SYSTEM_TEAMFOUNDATIONCOLLECTIONURI}
         else {Throw 'No value provided for "Org".'}
      ),
      [string]
      $Auth = $(
         if (Get-Variable SYSTEM_ACCESSTOKEN -ErrorAction 'Continue' 2> $null) {$SYSTEM_ACCESSTOKEN}
         elseif (Test-Path Env:SYSTEM_ACCESSTOKEN) {$Env:SYSTEM_ACCESSTOKEN}
         else {Throw 'No value provided for "Auth".'}
      ),
      [DateTime]
      $ReleaseDate = $(
         if (Get-Variable ReleaseDate -ErrorAction 'Continue' 2> $null) {$ReleaseDate}
         elseif (Test-Path Env:ReleaseDate) {[System.DateTime]::Parse($Env:ReleaseDate, [System.Globalization.CultureInfo]::InvariantCulture)}
         else {[DateTime]::now}
      ),
      [string]
      $Directory = $(
         if (Get-Variable AGENT_TEMPDIRECTORY -ErrorAction 'Continue' 2> $null) {$AGENT_TEMPDIRECTORY}
         elseif (Test-Path Env:AGENT_TEMPDIRECTORY) {$Env:AGENT_TEMPDIRECTORY}
         else {Throw 'No value provided for "Directory".'}
      ),
      [string]
      $EmailTo = $(
         if (Get-Variable EmailTo -ErrorAction 'Continue' 2> $null) {$EmailTo}
         elseif (Test-Path Env:EmailTo) {$Env:EmailTo}
         else {"FIGDevOps@swbc.com"}
      ),
      [string]
      $EmailFrom = "FIGDevOps@swbc.com",
      [string]
      $SmtpServer = "smtp.swbc.local",
      [string]
      $SmtpPort = "25",
      [string]
      $ChromeExe
   )

   try {
      # Load all dlls in folder.
      Get-ChildItem -Path "$PSScriptRoot\dlls" -Filter '*.dll' | ForEach-Object { [System.Reflection.Assembly]::LoadFrom($_.FullName) > $null }
      # https://github.com/PowerShell/PowerShell/issues/5818
      if ($PSVersionTable.PSVersion.Major -ge 6) {
         $PSDefaultParameterValues["Invoke-RestMethod:SkipHeaderValidation"] = $true
         $PSDefaultParameterValues["Invoke-WebRequest:SkipHeaderValidation"] = $true
      }

      $AzureDevopsPat = @{authorization = "Basic" + [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$Auth"))}
      $AzureUri = $Org + "_apis/wit/wiql?api-version=6.0"
      $BatchUri = $Org + "_apis/wit/workitemsbatch?`$expand=all&api-version=6.0"

      $hasWorkitem = $true

      #DataManagement
      $BatchDataManagement = @()
      $BatchDataManagement += Invoke-DM|Convert-DM

      # Release Items
      Write-Host "Call Azure DevOps API for Release Items."
      $BatchReleases = @()
      $BatchReleases += Invoke-Release|Convert-Release
      $BatchReleasesGrouped = @()
      $BatchReleasesGrouped += $BatchReleases|GroupSort-ReleaseItem

      # AS400
      $BatchAS400 = @()
      $BatchAS400 += Invoke-AS400|Convert-AS400
      $BatchAS400Grouped = @()
      $BatchAS400Grouped += $BatchAS400|GroupSort-ReleaseItem

      $myTemplateRenderedInline = Render-ReleaseAgenda

      if($hasWorkitem -eq $true){
         $Signature = Get-Content -Raw "$PSScriptRoot\Signature.html"
         Write-Verbose 'Building data object.'
         $myData = ([PSCustomObject]@{
            Subject = "Deployment Agenda - $($ReleaseDate.ToString('dddd, yyyy-MM-dd'))";
            DayHumanFormat = $ReleaseDate.ToString('dddd');
            DateHumanFormat = $ReleaseDate.ToString('yyyy-MM-dd');
            DepartmentReleaseSet = @($BatchDataManagement) + @($BatchReleasesGrouped) + @($BatchAS400Grouped);
            stylesDoNotInline = (Get-Content -Raw "$PSScriptRoot\styles-doNotInline.css");
            Signature = $Signature
         })
         $myDataTech = ([PSCustomObject]@{
            Subject = "Deployments Starting - $($ReleaseDate.ToString('dddd, yyyy-MM-dd'))";
            DayHumanFormat = $ReleaseDate.ToString('dddd');
            DateHumanFormat = $ReleaseDate.ToString('yyyy-MM-dd');
            Releases = @(($BatchReleases + $BatchAS400) | Sort-Object -Property {$_.DD});
            stylesDoNotInline = (Get-Content -Raw "$PSScriptRoot\styles-doNotInline.css");
            Signature = $Signature
         })
         Write-Verbose "`$myData (Excluding stylesDoNotInline):`n$(ConvertTo-Xml ($myData | Select-Object -Property * -ExcludeProperty stylesDoNotInline ) -As String -Depth ([Int32]::MaxValue) )"
         Write-Verbose "`$myDataTech (Excluding stylesDoNotInline):`n$(ConvertTo-Xml ($myDataTech | Select-Object -Property * -ExcludeProperty stylesDoNotInline ) -As String -Depth ([Int32]::MaxValue) )"
         
            $tempResPath = "$PSScriptRoot/Agenda.html"
            Write-Host $tempResPath
            $myTemplate = Get-Content -Raw $tempResPath
            $style = Get-Content -Raw "$PSScriptRoot/styles.css"

            # # Register handlebars helper.
            # # [Action[Void]]$action = {
            # [Action]$action = {
            #    Param($conditional, $options)
            #    if(($conditional % 2) -eq 0) {
            #       return $options.fn($this);
            #     } else {
            #       return $options.inverse($this);
            #     }
            # }
            # [HandlebarsDotNet.Handlebars]::registerHelpers('if_even', $action)

            Write-Verbose 'Rendering templates.'
            $myFunc = [HandlebarsDotNet.Handlebars]::Compile( $myTemplate )
            $myTemplateRendered = $myFunc.Invoke($myData)
            # Inline CSS using PreMailer.
            # Method signature found at https://github.com/milkshakesoftware/PreMailer.Net/blob/v2.2.0/PreMailer.Net/PreMailer.Net/PreMailer.cs
            $myTemplateRenderedInline=[PreMailer.Net.PreMailer]::MoveCssInline(
               # string html,
               $myTemplateRendered,
               # bool removeStyleElements = false
               $true,
               # string ignoreElements = null
               ".doNotInline",
               # string css = null
               $style,
               # bool stripIdAndClassAttributes = false
               $false,
               # bool removeComments = false
               $true,
               # IMarkupFormatter customFormatter = null
               $null
            ).Html
            $myResultFile = [System.IO.Path]::Combine("$Directory", "Release-$($ReleaseDate.ToString('yyyy-MM-dd'))-Agenda.html")
            Write-Host "Write email body to file: `"$myResultFile`"."
            Set-Content -Path $myResultFile -Value $myTemplateRenderedInline -WhatIf:$False
            Write-Host "##vso[task.setvariable variable=AgendaFile]$myResultFile"
            ###########################Tech Notify Multiple######################
            $TechPath = "$PSScriptRoot\TechNotify.html"
            #TechNotify not being found using Get-Content so used [System.IO.File] and it worked.
            $myTech = [System.IO.File]::ReadAllText($TechPath)
            # $myTech = Get-Content -Raw $TechPath
            $myFuncTech = [HandlebarsDotNet.Handlebars]::Compile( $myTech )
            $myTechRendered = $myFuncTech.Invoke($myDataTech)
            # Inline CSS using PreMailer.
            # Method signature found at https://github.com/milkshakesoftware/PreMailer.Net/blob/v2.2.0/PreMailer.Net/PreMailer.Net/PreMailer.cs
            $myTechRenderedInline=[PreMailer.Net.PreMailer]::MoveCssInline(
               # string html,
               $myTechRendered,
               # bool removeStyleElements = false
               $true,
               # string ignoreElements = null
               ".doNotInline",
               # string css = null
               $style,
               # bool stripIdAndClassAttributes = false
               $false,
               # bool removeComments = false
               $true,
               # IMarkupFormatter customFormatter = null
               $null
            ).Html
            $myResultTech = [System.IO.Path]::Combine("$Directory", "Release-$($ReleaseDate.ToString('yyyy-MM-dd'))-TechNotify.html")
            Set-Content -Path $myResultTech -Value $myTechRenderedInline -WhatIf:$False
      

         if ($WhatIfPreference) {
            @($myResultFile, $myResultTech) | Show-FileInChrome
         } else {
            Write-Host "Send email to `"$($EmailTo)`"."
            Send-MailMessage-RM -Subject $myData.Subject -Body $myTemplateRenderedInline
            # EmailTechTable $TechmessageOut
            Send-MailMessage-RM -Subject $myDataTech.Subject -Body $myTechRenderedInline
         }
      }
      else {
         Write-Error "No Work Items returned by this query.
         Export to Excel spreadsheet unsuccessful.
         The form you're trying to export is blank. `r`n"
      }
   }
   catch{
      Write-Error $_ #logs error for error control
      throw
   }#error is present
}
 #Invoke-DM|Transform
function Invoke-DM {
   process {
      $ResultsDM = Invoke-RestMethod `
         -Uri $AzureUri `
         -Method Post `
         -ContentType "application/json" `
         -Headers $AzureDevopsPat `
         -Body (ConvertTo-Json -Compress @{
            query = ([Regex]::Replace($QueryDM, "{{ReleaseDate}}", $ReleaseDate.ToString('yyyy-MM-dd')))
         }) `
         -DisableKeepAlive `
         -Verbose:$false
      Write-Verbose "`$ResultsDM.WorkItems.Count: $($ResultsDM.WorkItems.Count)"
      if($ResultsDM.WorkItems.Count -gt 0){
         #takes Result ID's and puts them into a hashtable so they can be used in workitemsbatch api request
         #Requests a batch of work items so other properties such as state, dates, titles, etc. can be accessed
         Write-Host "Call Azure DevOps API for Data Management User Story Content."
         $BatchDataManagement = Invoke-RestMethod `
            -Uri $BatchUri `
            -Method Post `
            -ContentType "application/json" `
            -Headers $AzureDevopsPat `
            -Body (ConvertTo-Json -Compress @{
               ids = @($ResultsDM.workItems.id)
            }) `
            -DisableKeepAlive `
            -Verbose:$false
      }
      else{$BatchDataManagement = @()}
      Write-Output $BatchDataManagement
   }
}
function Convert-DM {
   param
   (
     [Parameter(Mandatory,ValueFromPipeline)]
     [object]
     $BatchDataManagement
   )
   process {
      if ($BatchDataManagement.Count -gt 0){
         # Map API results to data for processing.
         $BatchDataManagement = @(
            @($BatchDataManagement.value) `
            | ForEach-Object {
               Write-Output ([PSCustomObject]@{
               Type = $(switch ($_.fields.'System.WorkItemType') {
                     "Product Backlog Item" {'PBI'}
                     "User Story" {'US'}
                     "Bug" {'B'}
                     Default {'Item'}
                  });
               DDHumanFormat="N/A";
               HasRelease = $true;
               #Release = "_";
               #ReleaseId = "_";
               ReleaseName = "N/A";
               WID = $_.id;
               Description = if($null -ne $_.fields.'System.Title') {$_.fields.'System.Title'}
                  else {"No Description"};
               Risk = if($null -eq $_.fields.'Microsoft.VSTS.Common.Risk') {"<missing>"}
                  else {[System.Text.RegularExpressions.Regex]::Replace($_.fields.'Microsoft.VSTS.Common.Risk', "^\d - ", "")};
               Application = "N/A";
               Department = "Data Management";
               System = if($null -ne $_.fields.'System.TeamProject') {$_.fields.'System.TeamProject'}
                  else {"<missing>"};
               RelatedTasks = "N/A";
               RMAttendance = "N/A";
               })
            }
         )
            # Log data at this stage.
            # if ($VerbosePreference -ne 'SilentlyContinue') {
            #    $BatchDataManagement | %{[PSCustomObject]$_} `
            #    | Select-Object WID, ReleaseName, Application, Description, Risk, DDHumanFormat, RelatedTasks, Department, DD `
            #    | Sort-Object -Property Department, DD `
            #    | Format-Table -Wrap -GroupBy Department
            # }
            # Group and sort.
            $BatchDataManagement = @(
               $BatchDataManagement | Group-Object -Property {$_.Department} -AsHashTable -AsString `
               | ForEach-Object {$_.GetEnumerator()} `
               | ForEach-Object {
                  Write-Output ([PSCustomObject]@{
                     DepartmentName = $_.Key;
                     Releases = @($_.Value);
                  })
               }
            )
      }
      else {
         $BatchDataManagement = @()
      }
      if ($Env:LogBatchObjects)
         { Write-Host -ForegroundColor Green "`$BatchDataManagement:`n$(ConvertTo-Xml $BatchDataManagement -As String -Depth ([Int32]::MaxValue) )" }
      Write-Output $BatchDataManagement
   } 
}
function Invoke-Release {
   process {
      $ResultsRelease = Invoke-RestMethod `
         -Uri $AzureUri `
         -Method Post `
         -ContentType "application/json" `
         -Headers $AzureDevopsPat `
         -Body (ConvertTo-Json -Compress @{
            query = ([Regex]::Replace($QueryRelease, "{{ReleaseDate}}", $ReleaseDate.ToString('yyyy-MM-dd')))
         }) `
         -DisableKeepAlive `
         -Verbose:$false
      Write-Verbose "`$ResultsRelease.WorkItems.Count: $($ResultsRelease.WorkItems.Count)"
      if($ResultsRelease.WorkItems.Count -gt 0){
         #takes Result ID's and puts them into a hashtable so they can be used in workitemsbatch api request
         
         #Requests a batch of work items so other properties such as state, dates, titles, etc. can be accessed
         Write-Host "Call Azure DevOps API for Release Item Content."
         $BatchReleases = Invoke-RestMethod -Uri $BatchUri `
            -Method Post `
            -ContentType "application/json" `
            -Headers $AzureDevopsPat `
            -Body (ConvertTo-Json -Compress @{
               ids = @($ResultsRelease.workItems.id)
               fields = @(
                  "System.id",
                  "SWBC.Department",
                  "SWBC.Application",
                  "System.Title",
                  "System.WorkItemType",
                  "System.TeamProject",
                  "SWBC.Risk",
                  "Microsoft.VSTS.Common.DD",
                  "SWBC.RelatedTasks",
                  "SWBC.DevMan",
                  "SWBC.Developer",
                  "SWBC.Tester",
                  "Microsoft.VSTS.Build.FoundIn",
                  "SWBC.AutoStart"
               )
            }) `
            -DisableKeepAlive `
            -Verbose:$false 
         }
      else{$BatchReleases = @()}
      Write-Output $BatchReleases
   }
}
function Convert-Release {
   param
   (
     [Parameter(Mandatory,ValueFromPipeline)]
     [object]
     $BatchReleases
   )
   process {
      if ($BatchReleases.Count -gt 0){
         # Map API results to data for processing.
         $BatchReleases = @(
            @($BatchReleases.value.fields) `
            | ForEach-Object {
               $DD = [System.DateTime]::Parse($_.'Microsoft.VSTS.Common.DD', [System.Globalization.CultureInfo]::InvariantCulture)

               $teamProject = $_.'System.TeamProject'
               $releaseId = $null
               $releaseName = $null
               $HasRelease = $false
               if ($null -ne $_.'Microsoft.VSTS.Build.FoundIn')
               {
                  Write-Host "Call Azure DevOps API for Deployments."
                  $ResultsReleaseFound = Invoke-RestMethod -Uri "https://vsrm.dev.azure.com/SWBC-FigWebDev/$teamProject/_apis/release/deployments?queryOrder=decending&api-version=6.0" `
                     -Method Get `
                     -Headers $AzureDevopsPat `
                     -ContentType "application/json" `
                     -Verbose:$false
                  foreach ($value in $ResultsReleaseFound.value) {
                     if ($Env:LogFoundIn) {Write-Verbose "foundin: $($_.'Microsoft.VSTS.Build.FoundIn') artifact: $($value.release.artifacts[0].definitionReference.version.id) releaseid: $($value.release.id)"}
                     if ($value.release.artifacts[0].definitionReference.version.id -eq $_.'Microsoft.VSTS.Build.FoundIn')
                     {
                        $HasRelease = $true
                        $releaseId = $value.release.id
                        $releaseName = $value.release.name
                        break
                     }
                  }
               } else {
                  Write-Verbose 'Property "Microsoft.VSTS.Build.FoundIn" missing.'
               }

               Write-Output ([PSCustomObject]@{
                  Type = 'R'; # Hard code "R" because the query that gets $BatchReleases only matches "Release" items.
                  DD = $DD;
                  DDHumanFormat=$DD.ToString("hh:mm tt");
                  DevMan = if($null -eq $_.'SWBC.DevMan'.'displayName'){"N/A"} else {$_.'SWBC.DevMan'.'displayName'};
                  Developer = if($null -eq $_.'SWBC.Developer'.'displayName'){"N/A"} else {$_.'SWBC.Developer'.'displayName'};
                  Tester = if($null -eq $_.'SWBC.Tester'.'displayName'){"N/A"} else {$_.'SWBC.Tester'.'displayName'};
                  HasRelease = $HasRelease;
                  Release = $(if($HasRelease) {"https://dev.azure.com/SWBC-FigWebDev/RM.Test/_releaseProgress?_a=release-pipeline-progress&releaseId=$releaseId"} else {$null});
                  ReleaseId = $releaseId;
                  ReleaseName = $releaseName;
                  WID = $_.'System.id';
                  Description = $_.'System.Title';
                  Risk = [System.Text.RegularExpressions.Regex]::Replace($_.'SWBC.Risk', "^\d - ", "");
                  Application = if ($null -eq $_.'SWBC.Application') {"<missing>"} else {$_.'SWBC.Application'};
                  Department = if ($null -eq $_.'SWBC.Department') {"<Department Missing>"} else {$_.'SWBC.Department'};
                  System = $_.'System.TeamProject';
                  RelatedTasks = $_.'SWBC.RelatedTasks';
                  RMAttendance = $_.'SWBC.AutoStart';
               })
            }
         )
         # # Log data at this stage.
         # if ($VerbosePreference -ne 'SilentlyContinue') {
         #    $BatchReleases | %{[PSCustomObject]$_} `
         #    | Select-Object WID, ReleaseName, Application, Description, Risk, DDHumanFormat, RelatedTasks, Department, DD `
         #    | Sort-Object -Property Department, DD `
         #    | Format-Table -Wrap -GroupBy Department
         # }
      }
      else {
         $BatchReleases = @()
      }
      if ($Env:LogBatchObjects)
         { Write-Host -ForegroundColor Green "`$BatchReleases:`n$(ConvertTo-Xml $BatchReleases -As String -Depth ([Int32]::MaxValue) )" }
      Write-Output $BatchReleases
   } 
}
function GroupSort-ReleaseItem{
   end{
      Write-Host $input
      if ($null -ne $input){
         $ObjectGrouped = @( 
            $input | Group-Object -Property {$_.Department} -AsHashTable -AsString `
            | ForEach-Object {$_.GetEnumerator()} `
            | ForEach-Object {
               Write-Output ([PSCustomObject]@{
                  DepartmentName = $_.Key;
                  Releases = @($_.Value | Sort-Object -Property {$_.DD} );
               })
            } `
            | Sort-Object -Property {$_.DepartmentName}
         )
      }
      else {
         $ObjectGrouped = @()
      }
      if ($Env:LogBatchObjects)
         { Write-Host -ForegroundColor Green "`$ObjectGrouped:`n$(ConvertTo-Xml $ObjectGrouped -As String -Depth ([Int32]::MaxValue) )" }
     Write-Output $ObjectGrouped
   } 
  
}
function Invoke-AS400 {
   process{
      Write-Host "Call Azure DevOps API for AS400 PBIs."
      $ResultsAS400 = Invoke-RestMethod `
         -Uri $AzureUri `
         -Method Post `
         -ContentType "application/json" `
         -Headers $AzureDevopsPat `
         -Body (ConvertTo-Json -Compress @{
            query = ([Regex]::Replace($QueryAS400, "{{ReleaseDate}}", $ReleaseDate.ToString('yyyy-MM-dd')))
         }) `
         -DisableKeepAlive `
         -Verbose:$false
      Write-Host $ResultsAS400.workitems
      Write-Verbose "`$ResultsAS400.WorkItems.Count: $($ResultsAS400.WorkItems.Count)"
      if($ResultsAS400.workItems.Count -gt 0) {
         #takes Result ID's and puts them into a hashtable so they can be used in workitemsbatch api request
         #Requests a batch of work items so other properties such as state, dates, titles, etc. can be accessed
         Write-Host "Call Azure DevOps API for AS400 PBI Content."
         $BatchAS400 = Invoke-RestMethod -Uri $BatchUri `
         -Method Post `
         -ContentType "application/json" `
         -Headers $AzureDevopsPat `
         -Body (ConvertTo-Json  -Compress @{
            ids = @($ResultsAS400.workItems.id)
            # fields = @(
            #    "System.id",
            #    "System.Title",
            #    "System.WorkItemType",
            #    "System.TeamProject",
            #    "SWBC.Risk",
            #    "Microsoft.VSTS.Common.ImpDate",
            #    "SWBC.ReleaseUrl",
            #    "SWBC.ReleaseName",
            #    "SWBC.RelatedTasks",
            #    "SWBC.DevMan",
            #    "SWBC.Developer",
            #    "SWBC.Tester",
            #    "Microsoft.VSTS.Build.FoundIn"
            # )
         }) `
         -DisableKeepAlive `
         -Verbose:$false
      }
      else {
         $BatchAS400 = @()
      }
      Write-Output $BatchAS400
   }
}
function Convert-AS400 {
   param
   (
     [Parameter(Mandatory,ValueFromPipeline)]
     [object]
     $BatchAS400
   )
   process {
      if ($BatchAS400.Count -gt 0){
         # Map API results to data for processing.
         $BatchAS400 = @(
            @($BatchAS400.value) `
            | ForEach-Object {
               $DD = [System.DateTime]::Parse($_.fields.'Microsoft.VSTS.Common.ImpDate', [System.Globalization.CultureInfo]::InvariantCulture)
                  
               Write-Output ([PSCustomObject]@{
                  Type = 'PBI'; # Hard code "PBI" because the query that gets $BatchAS400 only matches "PBI" items.
                  DD = $DD;
                  HasRelease = $true
                  ReleaseName = "<Not Implemented>"
                  DDHumanFormat=$DD.ToString("hh:mm tt");
                  DevMan = if($null -eq $_.'SWBC.DevMan'.'displayName'){"<missing>"} else {$_.'SWBC.DevMan'.'displayName'};
                  Developer = if($null -eq $_.'SWBC.Developer'.'displayName'){"<missing>"} else {$_.'SWBC.Developer'.'displayName'};
                  Tester = if($null -eq $_.'SWBC.Tester'.'displayName'){"<missing>"} else {$_.'SWBC.Tester'.'displayName'};
                  WID = $_.id;
                  Description = $_.fields.'System.Title';
                  Risk = [System.Text.RegularExpressions.Regex]::Replace($_.fields.'Microsoft.VSTS.Common.Risk', "^\d - ", "");
                  Application = $_.fields.'SWBC.ApplicationType';
                  Department = if ($_.fields.'SWBC.ApplicationType' -in ('Stone Eagle', 'Accounting')) {'AS400 Accounting'}
                  elseif($_.fields.'SWBC.ApplicationType' -in ('Mortgage', 'Auto','Gap')) {'AS400 Development'}
                  else {'AS400 Other'};
                  System = $_.fields.'System.TeamProject';
                  RelatedTasks = $_.fields.'SWBC.AS400Task';
                  RMAttendance = "N/A";
               })
            }
            Write-Host $BatchAS400
         ) 

         # # Log data at this stage.
         # if ($VerbosePreference -ne 'SilentlyContinue') {
         # $BatchAS400 | %{[PSCustomObject]$_} `
         # | Select-Object WID, ReleaseName, Application, Description, Risk, DDHumanFormat, RelatedTasks, Department, DD `
         # | Sort-Object -Property Department, DD `
         # | Format-Table -Wrap -GroupBy Department
         # }
      }
      else {
         $BatchAS400 = @()
      }
      if ($Env:LogBatchObjects)
         { Write-Host -ForegroundColor Green "`$BatchAS400:`n$(ConvertTo-Xml $BatchAS400 -As String -Depth ([Int32]::MaxValue) )" }
      Write-Output $BatchAS400
   }
   
}
function Render-ReleaseAgenda {
   param
   (
      [DateTime]
      $ReleaseDate = $(
         if (Get-Variable ReleaseDate -ErrorAction 'Continue' 2> $null) {$ReleaseDate}
         elseif (Test-Path Env:ReleaseDate) {[System.DateTime]::Parse($Env:ReleaseDate, [System.Globalization.CultureInfo]::InvariantCulture)}
         else {[DateTime]::now}
      ),
      [string]
      $Directory = $(
         if (Get-Variable Directory -ErrorAction 'Continue' 2> $null) {$Directory}
         else {Throw 'No value provided for "Directory".'}
      ),
      [string]
      $BatchDataManagement = $(
         if (Get-Variable BatchDataManagement -ErrorAction 'Continue' 2> $null) {$BatchDataManagement}
         else {Throw 'No value provided for "BatchDataManagement".'}
      ),
      [string]
      $BatchReleasesGrouped = $(
         if (Get-Variable BatchReleasesGrouped -ErrorAction 'Continue' 2> $null) {$BatchReleasesGrouped}
         else {Throw 'No value provided for "BatchReleasesGrouped".'}
      ),
      [string]
      $BatchAS400Grouped = $(
         if (Get-Variable BatchAS400Grouped -ErrorAction 'Continue' 2> $null) {$BatchAS400Grouped}
         else {Throw 'No value provided for "BatchAS400Grouped".'}
      )
   )
   # TODO: Implement Render-ReleaseAgenda.
}
function Show-FileInChrome {
   param
   (
      [Parameter(Mandatory,ValueFromPipeline)]
      [object]
      $myResultFile
   )
   end {
      if ( $ChromeExe ) {
         Write-Host 'Open Chrome.'
         if ($PSVersionTable.PSVersion.Major -gt 5) { $WhatIfStartProcess = @{WhatIf=$false} }
         else {$WhatIfStartProcess = @{}}
         Start-Process -FilePath $ChromeExe -ArgumentList @($input + '--new-window') @WhatIfStartProcess
      }
   }
}
function Send-MailMessage-RM {
   Param(
      [string]$Subject,
      [string]$Body
   )
   Send-MailMessage `
      -SmtpServer $smtpServer `
      -Port $SmtpPort `
      -To $EmailTo `
      -From $EmailFrom  `
      -Subject $Subject  `
      -Body $Body `
      -BodyAsHtml `
      -Attachments "$PSScriptRoot\SWBC_Corporate_RGB.small.png"
}

Export-ModuleMember -Function SendReleaseAgenda, `
Invoke-DM, `
Convert-DM, `
Invoke-Release, `
Convert-Release, `
Invoke-AS400, `
Convert-AS400, `
GroupSort-ReleaseItem