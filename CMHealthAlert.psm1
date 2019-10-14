function Set-CMAPadCount ($val, $len = 4) {
    if (![string]::IsNullOrEmpty($val)) {
        $val.ToString().PadLeft($len, ' ')
    }
}

function Write-CMALog {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $Message,
        [parameter()][ValidateNotNullOrEmpty()] [string] $LogFile = $(Join-Path $env:TEMP "ds-utils-$(Get-Date -f 'yyyyMMdd').log"),
        [parameter()][ValidateSet('Info','Error','Warning')] [string] $Category = 'Info'
    )
    try {
        $strdata = "$(Get-Date -f 'yyyy-MM-dd hh:mm:ss') - $Category - $Message"
        $strdata | Out-File -FilePath $LogFile -Append
        switch ($Category) {
            'Warning' { Write-Warning $strdata }
            'Error' { Write-Warning $strdata }
            default { Write-Host $strdata -ForegroundColor Cyan }
        }
    }
    catch {
        Write-Error "[Write-CMALog] $($Error[0].Exception.Message)"
    }
}

function Get-CMAClientHealthSummary {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $SiteServer,
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $Database,
        [parameter()] [ValidateNotNullOrEmpty()] [string] $qfile = "cm_client_health_summary.sql"
    )
    Write-CMALog -Message "requesting client health data"
    $qpath = Join-Path $PSScriptRoot $qfile
    try {
        @(Invoke-DbaQuery -SqlInstance $SiteServer -Database $Database -File $qpath -EnableException)
    }
    catch {
        Write-CMALog -Message $($_.Exception.Message) -Category Error
    }
}

function Get-CMAAppDeploymentExceptionsSummary {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $SqlInstance,
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $Database,
        [parameter()] [ValidateNotNullOrEmpty()] $qfile = "cm_deployment_exceptions_summary.sql"
    )
    Write-CMALog -Message "requesting deployment exceptions summary data"
    $qpath = Join-Path $PSScriptRoot $qfile
    try {
        $qfile = Join-Path $PSScriptRoot 
        @(Invoke-DbaQuery -SqlInstance $SqlInstance -Database $Database -File $qpath)
    }
    catch {
        Write-CMALog -Message $($_.Exception.Message) -Category Error
    }
}

function Get-CMAAppDeploymentExceptions {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $SqlInstance,
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $Database
    )
    Write-CMALog -Message "requesting deployment exceptions detail data"
    $qfile1 = Join-Path $PSScriptRoot "cm_deployment_exceptions_summary.sql"
    try {
        Write-CMALog -Message "importing: $qfile1"
        $dset = @(Invoke-DbaQuery -SqlInstance $SqlInstance -Database $Database -File $qfile1 -EnableException)
        $groupnum = 1
        $groupall = $dset.Count
        Write-CMALog -Message "total exception groups: $($dset.Count)"
        $qfile2 = Join-Path $PSScriptRoot "cm_app_deployment_exceptions.sql"
        Write-Verbose "importing: $qfile2"
        $basequery = Get-Content $qfile2 -Raw
        #Write-CMALog -Message $basequery
    }
    catch {
        Write-CMALog -Message "first stage: $($Error[0].Exception.Message -join '`n')" -Category Error
        break
    }
    foreach ($row in $dset) {
        $aid = $row.AssignmentID
        $dn  = $row.DeploymentName
        $cid = $row.TargetCollectionID
        $cn  = $row.CollectionName
        $ec  = $row.Error.ToInt()
        $q2  = $basequery -replace 'XXCOLLECTIONID', $cid
        Write-CMALog -Message "query length: $($q2.Length)"
        Write-CMALog -Message "deployment: $dn / collection: $cn / errors: $ec"
        $clients = @(Invoke-DbaQuery -SqlInstance $SqlInstance -Database $Database -Query $q2)
        $clientall = $clients.Count
        $clientnum = 1
        if ($clients.Count -gt 0) {
            foreach ($client in $clients) {
                [pscustomobject]@{
                    ComputerName  = $client.Name.ToString()
                    ResourceID    = $client.ResourceID.ToString()
                    ClientVersion = $client.ClientVersion.ToString()
                    Manufacturer  = $client.Manufacturer.ToString()
                    Model         = $client.Model.ToString()
                    SerialNumber  = $client.SerialNumber.ToString()
                    Deployment    = $dn.ToString()
                    DeploymentID  = $aid.ToString()
                    Collection    = $cn.ToString()
                    CollectionID  = $cid.ToString()
                    OSName        = $client.OSName.ToString()
                    OSBuild       = $client.OSBuild.ToString()
                    ADSiteName    = $client.ADSite.ToString()
                    LastUser      = $client.LastUser.ToString()
                    StatusMsg     = $client.StatusMsg.ToString()
                    StatusMsgID   = $client.LastSMID.ToString()
                    LastError     = $client.LastError.ToString()
                    StatusTime    = $client.StateTime.ToString()
                    RunHost       = $env:COMPUTERNAME
                    RunUser       = $env:USERNAME
                    ClientNum     = "$clientnum of $clientall"
                    TotalErrors   = $ec
                    GroupNum      = "$groupnum of $groupall"
                }
                $clientnum++
            } # foreach client in group
        }
        else {
            [pscustomobject]@{
                ComputerName  = $null
                ResourceID    = $null
                ClientVersion = $null
                Manufacturer  = $null
                Model         = $null
                SerialNumber  = $null
                Deployment    = $dn
                DeploymentID  = $aid
                Collection    = $cn
                CollectionID  = $cid
                OSName        = $null
                OSBuild       = $null
                ADSiteName    = $null
                LastUser      = $null
                StatusMsg     = $null
                StatusMsgID   = $null
                LastError     = $null
                StatusTime    = $null
                RunHost       = $env:COMPUTERNAME
                RunUser       = $env:USERNAME
                ClientNum     = "0 of 0"
                TotalErrors   = $ec
                GroupNum      = "$groupnum of $groupall"
            }
        }
        $groupnum++
    } # foreach deployment exception group
}

function Get-CMAClientVersionsSummary {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $SqlInstance,
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $Database
    )
    Write-CMALog -Message "requesting client versions data"
    try {
        $q1 = "select count(*) as [Total] from dbo.v_r_system"
        $tc = @(Invoke-DbaQuery -SqlInstance $SqlInstance -Database $Database -Query $q1)
        Write-CMALog -Message "q1 rows = $($tc.Total)"
        $q2 = "select version from v_site"
        $sitever = @(Invoke-DbaQuery -SqlInstance $SqlInstance -Database $Database -Query $q2 | Select-Object -ExpandProperty version)
        Write-CMALog -Message "site version = $sitever"
        $q3 = "select distinct Client_Version0 as ClientVersion, COUNT(*) Installs from dbo.v_r_system group by Client_Version0 order by Client_Version0"
        $clients = @(Invoke-DbaQuery -SqlInstance $SqlInstance -Database $Database -Query $q3)
        Write-CMALog -Message "q2 rows = $($clients.Count)"
        $clients | ForEach-Object {
            if ([string]::IsNullOrEmpty($_.ClientVersion)) {
                $cv = "NULL"
            }
            else {
                $cv = $_.ClientVersion
            }
            Write-CMALog -Message "version: $cv"
            $qty = [int]$_.Installs
            Write-CMALog -Message "qty: $qty"

            $pct = $($qty / $tc.Total)
            $pct = [math]::Round($pct * 100,0)
            Write-CMALog -Message "pct: $pct"

            [pscustomobject]@{
                Version  = [string]$cv
                Installs = $qty
                Percent  = $pct
                Site     = [string]$sitever
            }
        }
    }
    catch {
        Write-CMALog -Message $($_.Exception.Message) -Category Error
    }
}

function Get-CMAClientVersionsDetailed {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $SqlInstance,
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $Database,
        [parameter()] [ValidateNotNullOrEmpty()] [string] $qfile = "cm_old_clients.sql"
    )
    Write-CMALog -Message "getting client versions summary"
    try {
        $qpath = Join-Path $PSScriptRoot $qfile
        Invoke-DbaQuery -SqlInstance $SqlInstance -Database $Database -File $qpath | Foreach-Object {
            [pscustomobject]@{
                ComputerName  = [string]$_.ComputerName
                ClientVersion = [string]$_.ClientVersion
                ADSiteName    = [string]$_.ADSiteName
                UserName      = [string]$_.UserName
                OSName        = [string]$_.OSName
                Model         = [string]$_.Model
                LastHWInv     = [string]$_.LastHwInv
            }
        }
    }
    catch {
        Write-CMALog -Message $($_.Exception.Message) -Category Error
    }
}

function Get-CMAUpdatesComplianceSummary {
    [CmdletBinding()]
    param(
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $SqlInstance,
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $Database,
        [parameter()] [ValidateNotNullOrEmpty()] [string] $qfile = "cm_updates_compliance.sql"
    )
    Write-CMALog -Message "getting updates compliance summary"
    try {
        $qpath = Join-Path $PSScriptRoot $qfile
        @(Invoke-DbaQuery -SqlInstance $SqlInstance -Database $Database -File $qpath | Select-Object Name,ArticleID,PatchStatus,Title,InfoUrl)
    }
    catch {
        Write-CMALog -Message $($_.Exception.Message) -Category Error
    }
}

function Get-CMAUpdateDemploymentExceptions {
    param (
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $SqlInstance,
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $Database,
        [parameter()] [ValidateNotNullOrEmpty()] [string] $qfile = "cm_update_deployment_exceptions.sql"
    )
    Write-CMALog -Message "getting update deployment exceptions"
    try {
        $qpath = Join-Path $PSScriptRoot $qfile
        Invoke-DbaQuery -SqlInstance $SqlInstance -Database $Database -File $qpath | Foreach-Object {
            $deptotal = [int]$_.Total
            $success  = [int]$_.Success
            if ($deptotal -gt 0 -and $success -lt $deptotal) {
                # total > 0 and success < total, so calculate percentage
                $comp = [math]::Round($($success / $deptotal) * 100,1)
            }
            elseif ($deptotal -eq 0) {
                # prevent divide-by-zero mess
                $comp = 0
            }
            else {
                # total > 0 and success = total
                $comp = 100
            }
            [pscustomobject]@{
                Assignment     = [string]$_.AssignmentName
                SoftwareName   = [string]$_.SoftwareName
                PackageName    = [string]$_.PackageName
                ProgramName    = [string]$_.ProgramName
                PackageID      = [string]$_.PackageID
                PackageType    = [string]$_.PackageType
                CollectionID   = [string]$_.CollectionID
                CollectionName = [string]$_.CollectionName
                Total          = [string]$_.Total
                Success        = [string]$_.Success
                Failed         = [string]$_.Failed
                InProgress     = [string]$_.InProgress
                Unknown        = [string]$_.Unknown
                Other          = [string]$_.Other
                Exceptions     = [string]$_.Exceptions
                Compliance     = $comp
            }
        } # foreach-object
    }
    catch {
        Write-CMALog -Message $($_.Exception.Message) -Category Error
    }
}



function New-CMADataFileName {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)][ValidateNotNullOrEmpty()] [string] $BaseName
    )
    Join-Path -Path $OutputFolder -ChildPath "$SiteCode`-$BaseName`_$(Get-Date -f 'yyyyMMdd').csv"
}

Function Get-RegistryValue {
    param (
        [parameter()] [String] $ComputerName,
        [parameter()] [ValidateSet('LocalMachine','ClassesRoot','CurrentConfig','Users')] [string] $AccessType = 'LocalMachine',
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $KeyName,
        [parameter()] [string] $KeyValue
    )
    Write-Verbose "Getting registry value from $($ComputerName), $($AccessType), $($keyname), $($keyvalue)"
    try {
        $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($AccessType, $ComputerName)
        $RegKey= $Reg.OpenSubKey($keyname)
	    if ($RegKey -ne $null) {
		    try { $return = $RegKey.GetValue($keyvalue) }
		    catch { $return = $null }
	    }
	    else { $return = $null }
        Write-Verbose "Value returned $return"
    }
    catch {
        $return = "ERROR: Unknown"
        $Error.Clear()
    }
    , $return
}

function Get-CMACMDatabaseProps {
    [CmdletBinding()]
    param(
        [parameter(Mandatory)][string] $SMSSiteServer
    )
    try {
        $SQLServerName  = Get-RegistryValue -ComputerName $SMSSiteServer -LogFile -KeyName 'SOFTWARE\\Microsoft\\SMS\\SQL Server\\Site System SQL Account' -KeyValue 'Server'
        $SQLServiceName = Get-RegistryValue -ComputerName $SMSSiteServer -LogFile -KeyName 'SOFTWARE\\Microsoft\\SMS\\SQL Server' -KeyValue 'Service Name'
        $SQLPort        = Get-RegistryValue -ComputerName $SMSSiteServer -LogFile -KeyName 'SOFTWARE\\Microsoft\\SMS\\SQL Server\\Site System SQL Account' -KeyValue 'Port'
        $SQLDBName      = Get-RegistryValue -ComputerName $SMSSiteServer -LogFile -KeyName 'SOFTWARE\\Microsoft\\SMS\\SQL Server\\Site System SQL Account' -KeyValue 'Database Name'
        [pscustomobject]@{
            ServerName  = $SQLServerName
            ServiceName = $SQLServiceName
            Database    = $SQLDBName
            Port        = $SQLPort
        }
    }
    catch {
        Write-Error $_.Exception.Message 
    }
}

<#
.SYNOPSIS
    Send ConfigMgr Client Health Compliance and Exception report data by email
.DESCRIPTION
    Send ConfigMgr Client Health Compliance and Exception report data by email
.PARAMETER SiteCode
    3-character ConfigMgr site code
.PARAMETER SiteServer
    SmsProvider host. Default is "localhost"
.PARAMETER SendAlert
    Invoke Send-MailMessage (default is no mail sending)
.PARAMETER Message
    Custom message to include in mail message (default is none)
.PARAMETER MailSubject
    Default is "ConfigMgr Exceptions Report YYYY-MM-DD"
.PARAMETER MailTo
    Default is customer address (internal)
.PARAMETER MailSender
    Email address for sending email messages when -SendAlert is used
.PARAMETER MailFormat
    HTML or Text. The default is HTML
.PARAMETER SenderPassword
    The password for mail sender account
.PARAMETER OutputFolder
    Data collection folder path.  The default is user Documents folder
.PARAMETER MaxAge
    The default is 14 (days)
.PARAMETER Detailed
    Include table with updates compliance exceptions (default is to show numbers only)
.PARAMETER Port
    SendMailMessage port number. Default is 587
.NOTES
.INPUTS
.OUTPUTS
.LINK
    https://github.com/Skatterbrainz/CMHealthAlert/tree/master/docs/Invoke-CMAClientReport.md
#>
function Invoke-CMAClientReport {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)] [ValidateLength(3,3)] [string] $SiteCode,
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string] $SiteServer,
        [parameter(Mandatory)] [ValidateNotNullOrEmpty()] [string[]] $MailTo,
        [parameter(Mandatory)] [string] $MailSender,
        [parameter()] [ValidateNotNullOrEmpty()] [securestring] $SenderPassword,
        [parameter()] [switch] $SendAlert,
        [parameter()] [string] $Message = "ConfigMgr Health Compliance Report for $(Get-Date -f 'MMMM yyyy')",
        [parameter()] [string] $MailSubject = "ConfigMgr Health Compliance Report - $(Get-Date -f 'MMMM yyyy')",
        [parameter()] [ValidateSet('HTML','TEXT')] [string] $MailFormat = "HTML",
        [parameter()] [string] $SmtpServer = "smtp.office365.com",
        [parameter()] [string] $OutputFolder = "c:\cmhealthalerts",
        [parameter()] [int] $MaxAge = 30,
        [parameter()] [switch] $Detailed,
        [parameter()] [int] $Port = 587,
        [parameter()] [string] $LogFile = "$env:TEMP\cat_compliancereport_run.log"
    )
    $t1 = Get-Date
    $dbprops  = Get-CMACMDatabaseProps -SMSSiteServer $SiteServer
    $DBServer = $dbprops.ServerName
    $Database = $dbProps.Database
    try {
        if (!(Test-Path $OutputFolder)) {
            mkdir $OutputFolder -Force
            Write-CMALog -Message "folder created: $OutputFolder"
        }
        Write-CMALog -Message "begin: verifying database connectivity"
        if (-not (Test-DbaConnection -SqlInstance $SiteServer -EnableException -ErrorAction SilentlyContinue)) {
            throw "$SiteServer could not be accessed"
        }
        $attachments = @()
        Write-CMALog -Message "collecting information from site database"
        $chs = Get-CMAClientHealthSummary -SqlInstance $DBServer -Database $Database
        $ads = Get-CMAAppDeploymentExceptionsSummary -SqlInstance $DBServer -Database $Database
        $add = Get-CMAAppDeploymentExceptions -SqlInstance $DBServer -Database $Database
        $cvv = Get-CMAClientVersionsSummary -SqlInstance $DBServer -Database $Database
        $cvd = Get-CMAClientVersionsDetailed -SqlInstance $DBServer -Database $Database
        $ucs = Get-CMAUpdatesComplianceSummary -SqlInstance $DBServer -Database $Database
        $uds = Get-CMAUpdateDemploymentExceptions -SqlInstance $DBServer -Database $Database
    
        if ($chs.Count -gt 0) {
            Write-CMALog -Message "connection successful"
        }
        else {
            throw "$SiteServer [$SiteCode] is inaccessible"
            break
        }

        if ($MailFormat -eq 'TEXT') {
            $msgBody = "ConfigMgr Monthly Compliance Report: $SiteCode"
            $msgBody += "script version: $scriptVersion"
        }
        else {
            $msgBody = "<html><head>
<title>ConfigMgr Monthly Compliance Report</title>
<meta charset=`"utf-8`" />
<meta name=`"viewport`" content=`"width=device-width, initial-scale=1.0`">
<meta name=`"app.author`" content=`"David Stein @skatterbrainz`" />
<meta name=`"app.date`" content=`"2019.10.05`" />
<meta name=`"app.url`" content=`"https://github.com/skatterbrainz/cmhealthalert`" />
<style type=`"text/css`">
H1,H2,H3,H4,H5,H6 {
    font-family:Arial,Tahoma, sans-serif;
}
BODY {font-family: verdana; font-size: 10pt;}
.pagebreak { page-break-before: always; }
TH  {
    background-color:#c0c0c0;
    color:#000;
    padding: 3px;
    border: 2px solid black;
    font-family: Verdana, Geneva, Tahoma, sans-serif;
    font-size: 10pt;
}
TD  {
    padding: 3px;
    border: 1px solid black;
    font-family: Verdana, Geneva, Tahoma, sans-serif;
    font-size: 10pt;
}
TR:nth-child(even) { background-color:#fff;}
TR:nth-child(odd)  { background-color:#e1e1e1;}
TABLE	{border-collapse: collapse; width:500px}
</style>
</head>
`n<body>
<h1>ConfigMgr Site Exceptions Report: $SiteCode</h1>
<h2>$(Get-Date -f 'MMMM dd, yyyy hh:mm tt')</h2>
`n<p style=font-size:10pt>script version: $scriptVersion</p>"
            if (![string]::IsNullOrEmpty($Message)) {
                $msgBody += "<p>Message: $Message</p>"
            }
        }
        Write-CMALog -Message "analyzing: client version instances"
        if ($cvv.Count -gt 0) {
            $csvFile = New-CMADataFileName -BaseName "ClientVersions"
            $cvv | Export-Csv -Path $csvFile -NoTypeInformation -Force
            $attachments += $csvFile
            Write-CMALog -Message "dataset saved to file: $csvFile"
            Write-CMALog -Message "appending message content"
            if ($MailFormat -eq 'HTML') {
                $msgBody += "`n<h2>ConfigMgr Client Version Exceptions: Summary</h2>
<p>Devices with outdated ConfigMgr client installations.</p>"
                $msgBody += $cvv | ConvertTo-Html -Fragment
            }
            $cvdFile = New-CMADataFileName -BaseName "ClientVersionDetailed"
            $cvd | Export-Csv -Path $cvdFile -NoTypeInformation -Force
            Write-CMALog -Message "dataset saved to file: $cvdFile"
            Write-CMALog -Message "appending message content"
            if ($MailFormat -eq 'HTML') {
                $msgBody += "`n
`n<p>Attachments:</p>
<ul>
    <li>Summary: $(Split-Path $csvFile -Leaf)</li>
    <li>Detailed: $(Split-Path $cvdFile -Leaf)</li>
</ul>"
            }
        }
        Write-CMALog -Message "analyzing: client health exceptions"
        if ($chs.Count -gt 0) {
            $csvFile = New-CMADataFileName -BaseName "ClientHealthSummary"
            $chs | Export-Csv -Path $csvFile -NoTypeInformation -Force
            $attachments += $csvFile
            Write-CMALog -Message "dataset saved to file: $csvFile"
            $oldhw = @($chs | Where-Object {![string]::IsNullOrEmpty($_.HwInvAge) -and $_.HwInvAge -gt $MaxAge})
            $oldsw = @($chs | Where-Object {![string]::IsNullOrEmpty($_.SwInvAge) -and $_.SwInvAge -gt $MaxAge})
            $nohw  = @($chs | Where-Object {[string]::IsNullOrEmpty($_.HwInvAge)})
            $inact = $chs | Where-Object {$_.ClientActiveStatus -ne 'Active'}
            Write-CMALog -Message "appending message content"
            if ($MailFormat -eq 'HTML') {
                $msgBody += "`n
`n<h2>Client Health Exceptions</h2>
`n<p>Devices with inactive clients or outdated inventory data.</p>
`n<table id=t1>
    <tr><td>Total Clients</td><td align=`"right`">$($chs.Count)</td><td></td></tr>
    <tr><td>Inactive Clients</td><td align=`"right`">$($inact.Count)</td><td>Not reporting into site</td></tr>
    <tr><td>Old HW Inventory</td><td align=`"right`">$($oldhw.Count)</td><td>Inventory is more than $MaxAge days old</td></tr>
    <tr><td>Old SW Inventory</td><td align=`"right`">$($oldsw.Count)</td><td>Inventory is more than $MaxAge days old</td></tr>
    <tr><td>No HW Inventory</td><td align=`"right`">$($nohw.Count)</td><td>Inventory never reported</td></tr>
</table>
`n<p>Recommended Actions:<p>
<ul>
    <li>Confirm device is online and responding</li>
    <li>Review the client installation log (c:\windows\ccmsetup\logs\ccmsetup.log)</li>
    <li>Review the client LocationServices.log (c:\windows\ccm\logs)</li>
    <li>Attempt Client Repair or Reinstall from console</li>
</ul>
`n<p>Attachments:</p>
<ul>
    <li>$(Split-Path $csvFile -Leaf)</li>
</ul>"
            }
            else {
                $msgBody += "Client Health Exceptions
----------------------------------
Total Clients...... $(Set-CMAPadCount $chs.Count)
Inactive Clients... $(Set-CMAPadCount $inact.Count) (not reporting into site)
Old HW inventory... $(Set-CMAPadCount $oldhw.Count) (hw inventory is more than $MaxAge days old)
Old SW inventory... $(Set-CMAPadCount $oldsw.Count) (sw inventory is more than $MaxAge days old)
No HW Inventory.... $(Set-CMAPadCount $nohw.Count) (client has not reported hw inventory)
`n* See attachments for more information: $(Split-Path $csvFile -Leaf)"
            }
        }

        <#
        Write-CMALog -Message "analyzing: software updates compliance data"
        if ($ucs.Count -gt 0) {
            $csvFile = New-CMADataFileName -BaseName "UpdatesCompliance"
            $ucs | Export-Csv -Path $csvFile -NoTypeInformation -Force
            $attachments += $csvFile
            Write-CMALog -Message "dataset saved to file: $csvFile"
            Write-CMALog -Message "appending message content"
            if ($MailFormat -eq 'HTML') {
                $msgBody += "<h2>Software Updates Compliance Summary</h2>"
                $msgBody += "<p>Missing updates: $(($ucs | Where-Object{$_.PatchStatus -eq 'MISSING'}).Count)</p>"
                $msgBody += "<p>Attachment: $(Split-Path $csvFile -Leaf)</p>"
                if ($Detailed) {
                    $msgBody += $ucs | ConvertTo-Html -Fragment
                }
            }
        }
        #>

        Write-CMALog -Message "analyzing: software updates compliance data"
        if ($uds.Count -gt 0) {
            $uds = $uds | 
                Where-Object {$_.Total -gt 0 -and $_.Compliance -lt 100} | 
                    Select-Object Assignment,CollectionName,Total,Success,Failed,InProgress,Unknown,Other,Compliance
            $csvFile = New-CMADataFileName -BaseName "UpdateDeployments"
            $uds | Export-Csv -Path $csvFile -NoTypeInformation -Force
            $attachments += $csvFile
            Write-CMALog -Message "dataset saved to file: $csvFile"
            Write-CMALog -Message "appending message content"
            if ($MailFormat -eq 'HTML') {
                $msgBody += "`n
<h2>Software Updates Deployments Summary</h2>
`n<table>
  <tr>
    <td>Deployed with less than 100 percent success</td>
    <td>$($uds.Count)</td>
  </tr>
</table>"
            if ($Detailed) {
                $msgBody += $uds | ConvertTo-Html -Fragment
            }
            $msgBody += "`n
<p>Recommended Actions:</p>
<ul>
    <li>Confirm client is active and inventory is recent</li>
    <li>Review client UpdateDeployment.log (c:\windows\ccm\logs)</li>
    <li>Review client UpdateHander.log (c:\windows\ccm\logs)</li>
</ul>
<p>Attachments:</p>
<ul>
    <li>$(Split-Path $csvFile -Leaf)</li>
</ul>"
            }
        }

        Write-CMALog -Message "analyzing: deployment exceptions summary"
        if ($ads.Count -gt 0) {
            $csvFile = New-CMADataFileName -BaseName "DepExceptionSummary"
            $ads | Export-Csv -Path $csvFile -NoTypeInformation -Force
            $attachments += $csvFile
            Write-CMALog -Message "dataset saved to file: $csvFile"
            $adsFail  = $ads | Where-Object {$_.Error -gt 0}
            $adsFailT = $($ads | Select-Object -ExpandProperty Error | Measure-Object -Sum).Sum
            $adsProg  = $ads | Where-Object {$_.InProgress -gt 0}
            $adsProgT = $($ads | Select-Object -ExpandProperty InProgress | Measure-Object -Sum).Sum
            $adsUnk   = $ads | Where-Object {$_.Unknown -gt 0}
            $adsUnkT  = $($ads | Select-Object -ExpandProperty Unknown | Measure-Object -Sum).Sum
            $tcount = $(Set-CMAPadCount
         $ads.Count)
            Write-CMALog -Message "appending message content"
            if ($MailFormat -eq 'HTML') {
                $msgBody += "`n
<h2>Application Deployment Exceptions: Summary</h2>
<p>Deployments with at least one exception.</p>
`n<table id=t1>
    <tr><td>Total deployments</td><td align=`"right`">$(Set-CMAPadCount $ads.Count)</td></tr>
    <tr><td>Failed/Error</td><td align=`"right`">$(Set-CMAPadCount $adsFail.Count) / $tcount total = $adsFailT</td></tr>
    <tr><td>In-Progress</td><td align=`"right`">$(Set-CMAPadCount $adsProg.Count) / $tcount total = $adsProgT</td></tr>
    <tr><td>Unknown</td><td align=`"right`">$(Set-CMAPadCount $adsUnk.Count) / $tcount total = $adsUnkT</td></tr>
</table>
`n<p>Recommended Actions:</p>
<ul>
    <li>Review client AppDiscovery.log (c:\windows\ccm\logs)</li>
    <li>Review client AppEnforce.log (c:\windows\ccm\logs)</li>
</ul>
`n<p>Attachments:</p>
<ul>
    <li>$(Split-Path $csvFile -Leaf)</li>
</ul>"
            }
            else {
                $msgBody += "`n`n
Application Deployment Exceptions: Summary
Deployments with at least one exception.
----------------------------------
Total deployments.. $(Set-CMAPadCount $ads.Count)
Failed/Error....... $(Set-CMAPadCount $adsFail.Count) / $tcount total = $adsFailT
In-Progress........ $(Set-CMAPadCount $adsProg.Count) / $tcount total = $adsProgT
Unknown............ $(Set-CMAPadCount $adsUnk.Count) / $tcount total = $adsUnkT
Attachment......... $(Split-Path $csvFile -Leaf)"
            }
        }

        Write-CMALog -Message "analyzing: deployment exceptions detailed machines"
        if ($add.Count -gt 0) {
            $csvFile = New-CMADataFileName -BaseName "DepExceptionDetail"
            $add | Export-Csv -Path $csvFile -NoTypeInformation -Force
            $attachments += $csvFile
            Write-CMALog -Message "dataset saved to file: $csvFile"
            $deps = $add | Group-Object -Property Deployment | Select-Object Count,Name
            $dpa  = $add | Group-Object -Property ADSiteName | Select-Object Count,Name
            Write-CMALog -Message "appending message content"
            if ($MailFormat -eq 'HTML') {
                $msgBody += "`n
<h2>Application Deployment Exceptions: Detailed</h2>
`n<h3>Exceptions by Deployment</h3>
`n<p>Deployment exceptions grouped by deployment configuration.</p>
`n$($deps | ConvertTo-Html -Fragment)
`n<h3>Exceptions by AD Site</h3>
<p>Deployment exceptions grouped by Active Directory Site</p>
`n$($dpa | ConvertTo-Html -Fragment)
`n<table id=t1>
    <tr><td>Total deployments</td><td align=`"right`">$(Set-CMAPadCount $add.Count)</td></tr>
</table>
`n<p>Recommended Actions:</p>
<ul>
    <li>Confirm device is online and client is healthy</li>
    <li>Attempt client machine policy refresh</li>
    <li>Review client AppDiscovery.log (c:\windows\ccm\logs)</li>
    <li>Review client AppEnforce.log (c:\windows\ccm\logs)</li>
</ul>
`n<p>Attachments:</p>
<ul>
    <li>$(Split-Path $csvFile -Leaf)</li>
</ul>"
            }
            else {
                $dpp  = ($deps | Foreach-Object { ((Set-CMAPadCount $_.Count), $_.Name) -join ' '}) -join "`n"
                $adsn = ($dpa | Foreach-Object { ((Set-CMAPadCount $_.Count), $_.Name) -join ' '}) -join "`n"
                $msgBody += "`n`nApplication Deployment Exceptions: Detailed
    ----------------------------------
    Total deployments.......... $(Set-CMAPadCount
 $add.Count)
    `nExceptions by Deployment...
    $dpp
    `nExceptions by AD Site...
    $adsn"
            }
        }
        $t2 = Get-Date
        Write-CMALog -Message "preparing message for sending"
        $rt = New-TimeSpan -Start $t1 -End $t2
        $runtime = "total runtime: $($rt.Minutes)m:$($rt.Seconds)s | host: $($env:COMPUTERNAME) | user: $($env:USERNAME)"
        Write-CMALog -Message "appending message content"
        if ($MailFormat -eq 'TEXT') {
            $msgBody += "`n`n$runtime"
            $msgFile = Join-Path $OutputFolder -ChildPath "msgtemp.txt"
        }
        else {
            $reflinks = Import-Csv -Path $(Join-Path $PSScriptRoot "weblinks.csv")
            $msgBody += "`n<h2>References</h2><ul>"
            $reflinks | Foreach-Object {
                $lnk = "<li><a href=`"$($_.url)`" target=_blank>$($_.description)</a></li>"
                $msgBody += $lnk
            }
            $msgBody += "</ul>"
            $msgBody += "<p>$runtime</p>"
            Write-CMALog -Message "appending message body data file path"
            $msgBody += "$(Join-Path $OutputFolder -ChildPath "msgtemp.html")"
        }
        $msgBody | Out-File $msgFile -Force
        if ($SendAlert) {
            Write-CMALog -Message "sending mail message to $MailTo"
            Write-CMALog -Message "attachments: $($attachments.Count)"
            $SecPwd = ConvertTo-SecureString -String $SenderPassword -AsPlainText -Force
            $cred = [pscredential]::new($MailSender, $SecPwd)
            $msgParams = @{
                From        = $MailSender
                To          = $MailTo
                Subject     = $MailSubject
                Body        = $msgBody
                SmtpServer  = $SmtpServer
                Port        = $Port
                BodyAsHtml  = $True
                Attachments = $attachments
                UseSSL      = $True
                Credential  = $cred
            }
            Send-MailMessage @msgParams -ErrorAction Stop
            Write-CMALog -Message "message has been sent"
        }
        else {
            Write-CMALog -Message "no message sent. content saved to $msgFile"
        }
        Write-Output "process completed"
    }
    catch {
        Write-CMALog -Message $($_.Exception.Message) -Category Error
    }
}