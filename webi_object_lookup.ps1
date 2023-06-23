
<#
 .Synopsis
  Retrieves X-SAP-Logontoken.

 .Description
  Retrieves X-SAP-Logontoken that is required to authenticate all subsequent REST API calls.

 .Parameter BaseUrl
  Base url of the RESTful Web Service that you want to connect to in format 'http://<host>:<port>/biprws'.

 .Parameter Cred
  Credentials used to generate the logon token passed as PSCredential object.

 .Parameter AuthMode
  Authentication mode to be used eg.
  - secWinAD
  - secLDAP
  - secEnterprise

 .Example
  Get-SapBILogonToken -BaseUrl "http://sapbi.samplehost.net:463/biprws" -AuthMode "secWinAD" -Cred "JohnDoe"
#>
function Get-SapBILogonToken {
    [CmdletBinding()]
    param (
        # REST Web Service base url
        [Parameter(Mandatory)]
        [string]$BaseUrl,

        # Login credentials
        [Parameter(Mandatory)]
        [pscredential]$Cred,        

        # Authentication mode 
        [Parameter()]
        [string]$AuthMode = "secWinAD"
    )
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"  
    $headers.Add("Accept","application/json")  
    $headers.Add("Content-Type","application/json")

    $logonObj = [PSCustomObject]@{
        userName = $Cred.UserName
        password = $Cred.GetNetworkCredential().Password
        auth = $AuthMode
    }
    
    $logonJson = ConvertTo-Json $logonObj
    $logonUrl = "${BaseUrl}/logon/long"
    try {
        $response = Invoke-RestMethod -Uri $logonURL -Method Post -Headers $headers -Body $logonJson    
    }
    catch {
        Write-Host "An error occurred:"
        Write-Host $_
    }
    
    return "`"" + $response.logonToken + "`""
}

<#
 .Synopsis
  Invalidates X-SAP-Logontoken.

 .Description
  Invalidates X-SAP-Logontoken so that it can't be used to authenticate REST API calls.

 .Parameter BaseUrl
  Base url of the RESTful Web Service that you want to connect to in format 'http://<host>:<port>/biprws'.

 .Parameter Token
  Token to be invalidated.

 .Example
  Remove-SapBILogonToken -BaseUrl "http://sapbi.samplehost.net:463/biprws" -Token "AAAAASFDGDFGDF.....DDFsd"
#>
function Remove-SapBILogonToken {
    [CmdletBinding()]
    param (
        # REST Web Service base url
        [Parameter(Mandatory)]
        [string]$BaseUrl,

        # Sap Logon Token
        [Parameter(Mandatory)]
        [string]$Token
    )

    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"  
    $headers.Add("Accept","application/json")  
    $headers.Add("Content-Type","application/json")
    $headers.Add("X-SAP-Logontoken",$Token)
    $logoffUrl = "${BaseUrl}/logoff"
 
    try {
        Invoke-RestMethod -Uri $logoffURL -Method Post -Headers $headers   
    }
    catch {
        Write-Host "An error occurred:"
        Write-Host $_
    }
}

<#
 .Synopsis
  Retrieves a list occurrences of objects in ObjectsArray in all the data provider of a Webi report.

 .Description
  Function returns a list with all occurrences of the object given in the ObjectsArray parameter.
  The list contains report path, report name, data provider name and object name.

 .Parameter BaseUrl
  Base url of the RESTful Web Service that you want to connect to in format 'http://<host>:<port>/biprws'.

 .Parameter ReportId
  Report id for the report which you want to check for the existence of objects in the data providers.

 .Parameter Token
  X-SAP-Logontoken required to authenticate all subsequent REST API calls.

 .Parameter ObjectsArray
  Array of object names that you want to check.

 .Example
  ProcessWebiReport -BaseUrl "http://sapbi.samplehost.net:463/biprws/raylight/v1" -ReportId "123453" -Token "AAAAASFDGDFGDF.....DDFsd" -ObjectsArray $("Object 1","Object 2")
#>
function ProcessWebiReport {
    [CmdletBinding()]
    param (
        # REST Web Service base url
        [Parameter(Mandatory)]
        [string]$BaseUrl,

        # Report id (Remark: Id is an integer number, it is not the Cuid which is a string of alphanumeric character)
        [Parameter(Mandatory)]
        [string]$ReportId,

        # Sap Logon Token
        [Parameter(Mandatory)]
        [string]$Token,

        # Array of object names to lookup
        [Parameter(Mandatory)]
        [string[]]$ObjectsArray
    )
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"  
    $headers.Add("Accept","application/json")  
    $headers.Add("Content-Type","application/json")
    $headers.Add("X-SAP-Logontoken",$Token)

    $webiUrl = "$BaseUrl/raylight/v1"
    $resultObjects = @()
    $report = Invoke-RestMethod -Uri "${webiUrl}/documents/${ReportId}" -Method Get -Headers $headers

    $dataProviders = Invoke-RestMethod -Uri "${webiUrl}/documents/${ReportId}/dataproviders" -Method Get -Headers $headers

    $headers["Accept"]="text/xml"
    $headers["Content-Type"]="text/xml"
    foreach ($dataProvider in $dataProviders.dataproviders.dataprovider) {
        $querySpecification = Invoke-RestMethod -Uri "${webiUrl}/documents/${ReportId}/dataproviders/$($dataProvider.id)/specification" -Method Get -Headers $headers
        foreach ($objectName in $ObjectsArray) {
            $foundObjects = $querySpecification.QuerySpec.queriesTree.children.bOQuery.resultObjects | Where-Object -Property "name" -EQ -Value $objectName
            foreach ($foundObject in $foundObjects) {
                $resultObject = New-Object -TypeName PSObject
                $o = [ordered]@{ ReportPath=$report.document.path; ReportName=$report.document.name; DataProvider=$dataProvider.name; ObjectName=$foundObject.name}
                $resultObject | Add-Member -NotePropertyMembers $o -TypeName resultObject
                $resultObjects += $resultObject
            }
        }
    }

    return $resultObjects
}

<#
 .Synopsis
  Retrieves a list occurrences of objects in ObjectsArray in all the data provider of a Webi reports from a given folder.

 .Description
  Function returns a list with all occurrences of the objects given in the ObjectsArray parameter for all the report from the given folder.
  The list contains report path, report name, data provider name and object name.

 .Parameter BaseUrl
  Base url of the RESTful Web Service that you want to connect to in format 'http://<host>:<port>/biprws'.

 .Parameter FolderId
  Folder id for the folder containing report which you want to check for the existence of objects in the data providers.

 .Parameter Token
  X-SAP-Logontoken required to authenticate all subsequent REST API calls.

 .Parameter ObjectsArray
  Array of object names that you want to check.

 .Example
  ProcessReportFolder -BaseUrl "http://sapbi.samplehost.net:463/biprws" -Folder "123456" -Token "AAAAASFDGDFGDF.....DDFsd" -ObjectsArray $("Object 1","Object 2")
#>
function ProcessReportFolder {
    [CmdletBinding()]
    param (
        # Web Service base url
        [Parameter(Mandatory)]
        [string]$BaseUrl,

        # Folder id (Remark: Id is an integer number, it is not the Cuid which is a string of alphanumeric character)
        [Parameter(Mandatory)]
        [string]$FolderId,

        # Sap Logon Token
        [Parameter(Mandatory)]
        [string]$Token,

        # Array of object names to lookup
        [Parameter(Mandatory)]
        [string[]]$ObjectsArray
    )
    
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"  
    $headers.Add("Accept","application/json")  
    $headers.Add("Content-Type","application/json")
    $headers.Add("X-SAP-Logontoken",$Token)

    $webiUrl   = "$BaseUrl/raylight/v1"

    $folderContent = Invoke-RestMethod -Uri "${BaseUrl}/infostore/${FolderId}/children" -Method Get -Headers $headers
    $reports = $folderContent.entries | Where-Object { $_.type -eq "Webi" }

    $resultObjects = @()
    foreach ($report in $reports) {
        $resultObjects += (ProcessWebiReport -WebiUrl $webiUrl -ReportId $report.id -Token $headers["X-SAP-Logontoken"] -ObjectsArray $ObjectsArray)
    }

    return $resultObjects
}

# Sample script for checking two objects in reports from one specific folder

$baseUrl   = "http://sapbi.samplehost.net:463/biprws"
$objectNames = @("Object one","Object two")
$folderId = "123456"
$token = Get-SapBILogonToken -BaseUrl $baseURL -AuthMode "secWinAD" -Cred "Jimmy"

ProcessReportFolder -BaseUrl $baseUrl -FolderId $folderId -Token $token -ObjectsArray $objectNames

Remove-SapBILogonToken -BaseUrl $baseURL -Token $token