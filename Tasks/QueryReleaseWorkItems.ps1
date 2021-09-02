
#######################################Environment Variable#################################################
$Org = $Env:SYSTEM_TEAMFOUNDATIONCOLLECTIONURI
$Auth = $Env:SYSTEM_ACCESSTOKEN
$ReleaseDate = get-date -date "01/02/2099" -Format "yyyy-MM-dd" #-date "07/15/2021"
$Directory = $Env:AGENT_TEMPDIRECTORY
$EmailToFunc = "FIGDevOps@swbc.com"

##########################################Function############################################################
function EmailTable($messageOut) {
   #structure for email body remove -date for live testing
   $agendaDate = $(Get-Date -date "01/02/2099" -Format "dddd MM/dd/yyyy") #-date "07/15/2021" 
   $agendaMessage = @"
   <p><b> Agenda for $($agendaDate)<p>
"@
   #fetch updated html file for use
   $tableBody= Get-Content `
   -Path "$Directory\Release-$ReleaseDate-Template.html"
   #add agenda Message and  table body together to be loaded into body of email
   $emailMessage = $agendaMessage + $tableBody

   $smtpServer = "smtp.swbc.local"
   $smtpPort = "25"
   $Esubject = "$(Get-Date -date "01/02/2099" -Format "dddd MM/dd/yyyy") Cab Agenda" #remove -date for live tests -date "07/15/2021"
   $EmailFrom = "FIGDevOps@swbc.com"
   Send-MailMessage `
   -To $EmailToFunc `
   -From $EmailFrom  `
   -Subject $Esubject  `
   -Body $emailMessage `
   -BodyAsHtml `
   -SmtpServer $smtpServer `
   -Port $smtpPort
}
function EmailTechTable($TechmessageOut) {
   #structure for email body remove -date for live testing
   $agendaDate = $(Get-Date -date "01/02/2099" -Format "dddd MM/dd/yyyy") #-date "07/15/2021" 
   $agendaMessage = @"
   <p><b> Please Reply for the following tasks to be scheduled.<p>
"@
   #fetch updated html file for use
   $tableBody= Get-Content `
   -Path "$Directory\Release-$ReleaseDate-TechNotify.html"
   #add agenda Message and  table body together to be loaded into body of email
   $emailMessage = $agendaMessage + $tableBody

   $smtpServer = "smtp.swbc.local"
   $smtpPort = "25"
   $Esubject = "Deployment Coordination - $(Get-Date -date "01/02/2099" -Format "dddd MM/dd/yyyy")" #remove -date for live tests -date "07/15/2021"
   $EmailFrom = "FIGDevOps@swbc.com"
   Send-MailMessage `
   -To $EmailToFunc `
   -From $EmailFrom  `
   -Subject $Esubject  `
   -Body $emailMessage `
   -BodyAsHtml `
   -SmtpServer $smtpServer `
   -Port $smtpPort
}
##########################################Process##############################################################
try {
   $AzureDevopsPat = @{authorization = "Basic" + [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(":$Auth"))}
   $AzureUri = $Org + "_apis/wit/wiql?api-version=6.0"
   $BatchUri = $Org + "_apis/wit/workitemsbatch?`$expand=all&api-version=6.0"
   #gets current date when run [-date $(get-date).adddays(-30)] use when necessary to see days before. Remove -date when live testing for current date -date "07/15/2021"
   $ImpDate = get-date `
   -date "01/02/2099" `
   -Format "yyyy-MM-dd"

   $hasWorkitem = $true

   #######query body in Wiql##################################################################################################################################################################
   $PostBody = "{
      'query' : 'SELECT * FROM WorkItems WHERE [System.WorkItemType]=`"Release`" AND ([Microsoft.VSTS.Common.DD]= `"$ImpDate`" OR [Microsoft.VSTS.Common.ImpDate]= `"$ImpDate`")'
   }"
   $PostBody2 = "{
      'query' : 'SELECT * FROM WorkItems WHERE [System.WorkItemType]=`"Product Backlog Item`" AND ([Microsoft.VSTS.Common.DD]= `"$ImpDate`" OR [Microsoft.VSTS.Common.ImpDate]= `"$ImpDate`")'
   }"
     ##########################################################################################################################################################################################
   #Stores request in $Results
   $Results = Invoke-RestMethod  -Uri $AzureUri -Method Post -ContentType "application/json" -Headers $AzureDevopsPat -Body $PostBody -DisableKeepAlive
   $Results2 = Invoke-RestMethod  -Uri $AzureUri -Method Post -ContentType "application/json" -Headers $AzureDevopsPat -Body $PostBody2 -DisableKeepAlive
   if($Results.WorkItems.Count -gt 0){
      #takes Result ID's and puts them into a hashtable so they can be used in workitemsbatch api request
      $PostBodyBatch = ConvertTo-Json @{
         ids = @($Results.workItems.id)
         fields = "System.id", "System.Title", "System.WorkItemType", "System.TeamProject", "SWBC.Risk", "Microsoft.VSTS.Common.DD", "SWBC.ReleaseUrl", "SWBC.ReleaseName", "SWBC.RelatedTasks", "SWBC.DevMan", "SWBC.Developer", "SWBC.Tester", "Microsoft.VSTS.Build.FoundIn", "SWBC.Application"
      }
      #Requests a batch of work items so other properties such as state, dates, titles, etc. can be accessed
      $Batch = Invoke-RestMethod -Uri $BatchUri `
      -Method Post `
      -ContentType "application/json" `
      -Headers $AzureDevopsPat `
      -Body $PostBodyBatch `
      -DisableKeepAlive

      # Removes every value except name from specified fields.
      foreach ($item in $Batch) {
         if($item.value.fields.'System.WorkItemType' -eq 'Release')
         {
            $item.value.fields.'System.WorkItemType' = "R"
         }
         
         $item.fields.'Microsoft.VSTS.Build.FoundIn'
         $DD = [System.DateTime]::Parse($item.value.fields.'Microsoft.VSTS.Common.DD').ToString("yyyy-MM-dd HH:mm tt")
         $item.value.fields.'Microsoft.VSTS.Common.DD' = $DD
         $DevMan = $item.value.fields.'SWBC.DevMan'.'displayName'
         $item.value.fields.'SWBC.DevMan' = $DevMan
         $Developer = $item.value.fields.'SWBC.Developer'.'displayName'
         $item.value.fields.'SWBC.Developer' = $Developer
         $Tester = $item.value.fields.'SWBC.Tester'.'displayName'
         $item.value.fields.'SWBC.Tester' = $Tester
         $tP = $item.value.fields.'System.TeamProject'
         $Results4 = Invoke-RestMethod -Uri "https://vsrm.dev.azure.com/SWBC-FigWebDev/$tp/_apis/release/deployments?queryOrder=decending&api-version=6.0" `
         -Method Get `
         -Headers $AzureDevopsPat `
         -ContentType "application/json"
         $releaseId = $null
         $releaseName = $null
         foreach ($value in $Results4.value) {
            if ($value.release.artifacts[0].'definitionReference.version.id' -eq $item.fields.'Microsoft.VSTS.Build.FoundIn')
            {
               $releaseId = $value.release.id
               $releaseName = $value.release.name
               break
            }
         }
         $releaseUrl = "https://dev.azure.com/SWBC-FigWebDev/RM.Test/_releaseProgress?_a=release-pipeline-progress&releaseId=$releaseId"
         $item.value.fields.'SWBC.ReleaseUrl' = $releaseUrl
         $item.value.fields.'SWBC.ReleaseName' = $releaseName
      }


   }
   if($Results2.WorkItems.Count -gt 0)
   {
      #takes Result ID's and puts them into a hashtable so they can be used in workitemsbatch api request
      $PostBodyBatch2 = ConvertTo-Json @{
         ids = @($Results2.workItems.id)
         fields = "System.id", "System.Title", "System.WorkItemType", "System.TeamProject", "SWBC.Risk", "Microsoft.VSTS.Common.ImpDate", "SWBC.ReleaseUrl", "SWBC.ReleaseName", "SWBC.RelatedTasks", "SWBC.DevMan", "SWBC.Developer", "SWBC.Tester", "Microsoft.VSTS.Build.FoundIn"
      }
      #Requests a batch of work items so other properties such as state, dates, titles, etc. can be accessed
      $Batch2 = Invoke-RestMethod -Uri $BatchUri `
      -Method Post `
      -ContentType "application/json" `
      -Headers $AzureDevopsPat `
      -Body $PostBodyBatch2 `
      -DisableKeepAlive
      #disableKeepAlive imediately stops connection after request is sent so the server isn't expecting a constant connection
      #without it, the server throws timeout errors
      # changes value that is accessible by workitem field
      foreach ($item2 in $Batch2) {
         if($item2.value.fields.'System.WorkItemType' -eq 'Product Backlog Item')
         {
            $item2.value.fields.'System.WorkItemType' = "PBI"
         }
         $Imp = [System.DateTime]::Parse($item2.value.fields.'Microsoft.VSTS.Common.ImpDate').ToString("yyyy-MM-dd HH:mm tt")
         $item2.value.fields.'Microsoft.VSTS.Common.ImpDate' = $Imp
         $DevMan2 = $item2.value.fields.'SWBC.DevMan'.'displayName'
         $item2.value.fields.'SWBC.DevMan' = $DevMan2
         $Developer2 = $item2.value.fields.'SWBC.Developer'.'displayName'
         $item2.value.fields.'SWBC.Developer' = $Developer2
         $Tester2 = $item2.value.fields.'SWBC.Tester'.'displayName'
         $item2.value.fields.'SWBC.Tester' = $Tester2
         $tP2 = $item2.value.fields.'System.TeamProject'
         $Results5 = Invoke-RestMethod -Uri "https://vsrm.dev.azure.com/SWBC-FigWebDev/$tp/_apis/release/deployments?queryOrder=decending&api-version=6.0" `
         -Method Get `
         -Headers $AzureDevopsPat `
         -ContentType "application/json"
         $releaseId2 = $null
         $releaseName2 = $null
         foreach ($value2 in $Results5.value) {
            if ($value2.release.artifacts[0].'definitionReference.version.id' -eq $item2.fields.'Microsoft.VSTS.Build.FoundIn')
            {
               $releaseId2 = $value2.release.id
               $releaseName2 = $value2.release.name
               break
            }
         }
         $releaseUrl2 = "https://dev.azure.com/SWBC-FigWebDev/$tp2/_releaseProgress?_a=release-pipeline-progress&releaseId=$releaseId2"
         $item2.value.fields.'SWBC.ReleaseUrl' = $releaseUrl2
         $item2.value.fields.'SWBC.ReleaseName' = $releaseName2
      }
   }
   elseif(($Results.WorkItems.Count -lt 1) -and ($Results2.WorkItems.Count -lt 1))
   {
      $hasWorkitem = $False
   }


   if($hasWorkitem -eq $true){

      #exports to an excel file
      
      $BatchEx1 = $Batch.value.fields
      $BatchEx1| Export-Csv `
      -Path "$Directory\Release1-$ReleaseDate.csv" `
      -NoTypeInformation
      $ChangeHeader1 = Get-Content `
      -Path "$Directory\Release1-$ReleaseDate.csv"
      $ChangeHeader1 = $ChangeHeader1[1..($ChangeHeader1.Count - 1)]
      $ChangeHeader1 | Out-File -FilePath "$Directory\Release1-$ReleaseDate.csv"
      $Header = 'WID', 'System', 'Type', 'Description', 'Build', 'DD', 'Risk', 'Release', 'DevMan', 'Developer', 'Tester', 'ReleaseName', 'RelatedTasks', 'Application'
      $importToEmail1 = @( Import-Csv -Path "$Directory\Release1-$ReleaseDate.csv" -Header $Header )
      $importToEmail1|Export-Csv `
      -Path "$Directory\Release1-$ReleaseDate.csv" `
      -NoTypeInformation

      # $BatchEx2 = $Batch2.value.fields
      # $BatchEx2|Export-Csv `
      # -Path "$Directory\Release2-$ReleaseDate.csv" `
      # -NoTypeInformation

      # $ChangeHeader2 = Get-Content `
      # -Path "$Directory\Release2-$ReleaseDate.csv"
      # $ChangeHeader2 = $ChangeHeader2[1..($ChangeHeader2.Count - 1)]
      # $ChangeHeader2 | Out-File -FilePath "$Directory\Release2-$ReleaseDate.csv"
      # $Header2 = 'WID', 'System', 'Type', 'Description', 'DD', 'Risk', 'Release', 'ReleaseName', 'RelatedTasks', 'DevMan', 'Tester', 'Developer', 'Build', 'Application'
      # $importToEmail1 = Import-Csv -Path "$Directory\Release2-$ReleaseDate.csv" -Header $Header2
      # $importToEmail1|Export-Csv `
      # -Path "$Directory\Release2-$ReleaseDate.csv" `
      # -NoTypeInformation
      # # combines both exported csv's into one if they have the same headers. 
      # Get-ChildItem  "$Directory/*.csv" | Select-Object -ExpandProperty FullName | Import-Csv | Export-Csv -Path "$Directory\Release-$ReleaseDate.csv" -NoTypeInformation

      Write-Host "Export Successful follow this path '$Directory\Release1-$ReleaseDate.csv' to locate release file."
      #augments Column headers set previously due to workitem fields.
      # $Header = 'WID', 'System', 'Type', 'Description', 'Build', 'DD', 'Risk', 'Release', 'DevMan', 'Tester', 'Developer', 'ReleaseName', 'RelatedTasks', 'Application'
      # $ChangeHeader3 = Get-Content `
      # -Path "$Directory\Release-$ReleaseDate.csv"
      # $ChangeHeader3 = $ChangeHeader3[1..($ChangeHeader3.Count - 1)]
      # $ChangeHeader3 | Out-File -FilePath "$Directory\Release-$ReleaseDate.csv"
      # #imports csv back to send in email to Release team
      # $importToEmail = Import-Csv -Path "$Directory\Release-$ReleaseDate.csv" -Header $Header #| ConvertTo-Html -Fragment
      #put csv into format usable by html and then append to html file Templat.html
      #appends template loaded with handlebars syntax, and replaces with content of imported csv
#change data back to $importToEmail once PBI work item fields have been updated
      [System.Reflection.Assembly]::LoadFrom("Dependencies\Handlebars.dll") > $null
      $myData = @{
         templateName = "myTemplate";
         data = $importToEmail1
      }
      if($Results.WorkItems.Count -gt 0)
      {
         $tempResPath = Resolve-Path "template.html.template.html"
         $myTemplate = Get-Content $tempResPath

         $myFunc = [HandlebarsDotNet.Handlebars]::Compile( $myTemplate )
         $myTemplateRendered = $myFunc.Invoke($myData)
         $myResultFile = "$Directory\Release-$ReleaseDate-Template.html" 
         Set-Content -Path $myResultFile -Value $myTemplateRendered
         ###########################Tech Notify Multiple######################
         $tempTechPath = Resolve-Path "TempTechNotify.html"
         $myTempTech = Get-Content $tempTechPath

         $myFuncTech = [HandlebarsDotNet.Handlebars]::Compile( $myTempTech )
         $myTempTechRendered = $myFuncTech.Invoke($myData)
         $myResultTech = "$Directory\Release-$ReleaseDate-TechNotify.html" 
         Set-Content -Path $myResultTech -Value $myTempTechRendered
      }
      #send email.
      EmailTable $messageOut
      EmailTechTable $TechmessageOut
      Write-Host "Email sent containing work Items for release, to $($EmailToFunc)"
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