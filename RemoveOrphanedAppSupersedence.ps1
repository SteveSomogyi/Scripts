<#	Remove orphaned references to superseded applications

	Author  : viki 
	Version : 1.0
	Last Update : 2017/01/17
	**********************************************************************************************************
	This sample is not supported under any Microsoft standard support program or service. This sample
	is provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties
	including, without limitation, any implied warranties of merchantability or of fitness for a particular
	purpose. The entire risk arising out of the use or performance of this sample and documentation
	remains with you. In no event shall Microsoft, its authors, or anyone else involved in the creation,
	production, or delivery of this sample be liable for any damages whatsoever (including, without limitation,
	damages for loss of business profits, business interruption, loss of business information, or other
	pecuniary loss) arising out of the use of or inability to use this sample or documentation, even
	if Microsoft has been advised of the possibility of such damages.
	***********************************************************************************************************

    Note: To run the script through the ISE, at least PS 4.0 is required. 
    Recommended update: https://www.microsoft.com/en-us/download/details.aspx?id=50395

    Examples:

    1. Verify only
       .\RemoveOrphanedAppSupersedence.ps1

    2. Update all applications with invalid supersedence reference
       .\RemoveOrphanedAppSupersedence.ps1 -update 
	
    Execute from Admin-CMD:
	PowerShell.exe -ExecutionPolicy ByPass -nologo -noprofile -file <script path>\RemoveOrphanedAppSupersedence.ps1
    PowerShell.exe -ExecutionPolicy ByPass -nologo -noprofile -file <script path>\RemoveOrphanedAppSupersedence.ps1 -update
#>

[CmdletBinding()]
Param (
	[Parameter(Mandatory = $false)]
	[switch]$update
)

$ErrorActionPreference = 'Stop'

# Import SCCM Module and set PS Drive
function Import-ConfigMgrModule ($provider) {
    $ErrorActionPreference = 'Stop'

    Try {
	    $SiteCode = (gwmi -Namespace 'root\sms' -query "SELECT SiteCode FROM SMS_ProviderLocation").SiteCode

        Import-Module ($env:SMS_ADMIN_UI_PATH + '\..\ConfigurationManager.psd1')
        $PSD = Get-PSDrive -Name $SiteCode
        if($PSD -eq $null) {
            $PSD = New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $provider
        }

        CD "$($PSD):"
    }
    Catch {
        write-host "Exception: $($_.Exception.Message)" -ForegroundColor Red
        break
    }
    return "root\sms\Site_$($SiteCode)"
}

$SMSProv = Import-ConfigMgrModule ('localhost')
###

$arrOrphaned = New-Object System.Collections.ArrayList
$arrUpdated = New-Object System.Collections.ArrayList
$arrValid = New-Object System.Collections.ArrayList
$arrApps = New-Object System.Collections.ArrayList

try
{
    # Counting applications
    [int]$countApps = gwmi -Namespace $SMSProv -Query 'select count(*) from SMS_ApplicationLatest' | select -ExpandProperty Count
    Write-Host "`n$countApps applications found. Importing...`n"

    # Importing application CIs
    $allApps = gwmi -Namespace $SMSProv -Query "select CI_ID, LocalizedDisplayName from SMS_ApplicationLatest"
}
Catch {
    write-host "Exception: $($_.Exception.Message)" -ForegroundColor Red
    break
}


# Verifying supersedence
if ($allApps)
{
    [int]$appCounter = 1
    foreach ($a in $allApps)
    {
        Write-Progress -Activity 'Looping through all applications' -PercentComplete ($appCounter / $countApps * 100) -CurrentOperation "Verifying application $appCounter/$countApps"
        $app = Get-CMApplication -Id $a.CI_ID | ? {$_.SDMPackageXML.Contains("<Supersedes>")}

        if ($app) { [void]$arrApps.Add($app) } # Adding app with supersedence to array
        $appCounter++
    }
    Write-Progress -Activity 'Looping through all applications' -Completed
}

if (!$arrApps) 
{
    Write-Host '`nNo superseding application found!' -ForegroundColor DarkYellow
    break
}

# Start logging...
try {
    Start-Transcript -Path ("$env:windir\temp\" + $($((Split-Path $MyInvocation.MyCommand.Definition -leaf)).replace("ps1","log"))) -Append
    Write-Host ''
}
catch { write-host "Exception: $($_.Exception.Message)" -ForegroundColor Yellow }

$appCounter = 1

foreach ($app in $arrApps)
{
    Write-Progress -Activity 'Verifying supersedence' -PercentComplete ($appCounter / $arrApps.Count * 100) -CurrentOperation "Superseding application $appCounter/$($arrApps.Count)"

    $AppExists = $null
    [int]$startSupersedes = $app.SDMPackageXML.IndexOf("<Supersedes>") 
    [int]$endSupersedes = $app.SDMPackageXML.IndexOf("</Supersedes>")

    # Get XML of Supersedence Section
    [xml]$sXML = $app.SDMPackageXML.SubString($startSupersedes, ($endSupersedes - $startSupersedes + "</Supersedes>".Length))

    # Get application reference(s) of superseded application(s)
    [string]$superseded = $sXML.Supersedes.DeploymentTypeRule.DeploymentTypeIntentExpression.DeploymentTypeApplicationReference.LogicalName

    # Extracting array of superseded application(s), in case there's more than one
    $arrSuperseded = $superseded.split(' ')
    
    $arrOrphaned.Clear()
    $arrValid.Clear()

    # Determine if the superseded app reference is orphaned or valid
    foreach ($s in $arrSuperseded)
    {
        $AppExists = gwmi -Namespace $SMSProv -Query "select * from SMS_ApplicationLatest where ModelName like '%$s%'"
        if (!$AppExists) {[void]$arrOrphaned.Add($s)}
        else {[void]$arrValid.Add($s)}
    }
    
    # Remove orphaned supersedence
    if ($arrOrphaned.Count -gt 0)
    {
        Write-Host "`nOrphaned supersedence reference(s) found for application: " -NoNewline
        Write-Host $($app.LocalizedDisplayName) -ForegroundColor Yellow
        foreach ($o in $arrOrphaned)
        {
            Write-Host "Orphaned reference: " -NoNewline
            Write-Host $($o) -ForegroundColor Yellow
        }
        foreach ($v in $arrValid)
        {
            Write-Host "Valid reference: " -NoNewline
            Write-Host $($v) -ForegroundColor Green
        }

        if (($arrSuperseded.Count -eq 1) -or ($arrSuperseded.Count -eq $arrOrphaned.Count))
        {
            $app.SDMPackageXML = $app.SDMPackageXML.Remove($startSupersedes, ($endSupersedes - $startSupersedes + "</Supersedes>".Length))
        }
        else # Multiple references
        {
            foreach ($o in $arrOrphaned)
            {
                # Get first start/end position of first <DeploymentTypeRule> section
                [int]$startDTR = $app.SDMPackageXML.IndexOf("<DeploymentTypeRule", $startSupersedes) 
                [int]$endDTR = $app.SDMPackageXML.IndexOf("</DeploymentTypeRule>")

                # Get <DeploymentTypeRule> section
                [xml]$sXML = $app.SDMPackageXML.SubString($startDTR, ($endDTR - $startDTR + "</DeploymentTypeRule>".Length))
                [string]$strDTR = $sXML.DeploymentTypeRule.DeploymentTypeIntentExpression.DeploymentTypeApplicationReference.LogicalName

                while ([array]::IndexOf($arrOrphaned, $strDTR) -eq -1)
                {
                    # Get start/end position of next <DeploymentTypeRule> section
                    [int]$startDTR = $app.SDMPackageXML.IndexOf("<DeploymentTypeRule", $endDTR) 
                    [int]$endDTR = $app.SDMPackageXML.IndexOf("</DeploymentTypeRule>", $startDTR)

                    # Get <DeploymentTypeRule> section
                    [xml]$sXML = $app.SDMPackageXML.SubString($startDTR, ($endDTR - $startDTR + "</DeploymentTypeRule>".Length))
                    [string]$strDTR = $sXML.DeploymentTypeRule.DeploymentTypeIntentExpression.DeploymentTypeApplicationReference.LogicalName
                }

                # Update SDMPackageXML
                $app.SDMPackageXML = $app.SDMPackageXML.Remove($startDTR, ($endDTR - $startDTR + "</DeploymentTypeRule>".Length))

                # Get new start/end position of <Supersedes> section
                [int]$startSupersedes = $app.SDMPackageXML.IndexOf("<Supersedes>") 
                [int]$endSupersedes = $app.SDMPackageXML.IndexOf("</Supersedes>")
            }

        }

        if ($update)
        {
            # Updating application
            Write-Host "Updating application: $($app.LocalizedDisplayName)..." -ForegroundColor Green -NoNewline
            try
            {
                Set-CMApplication -InputObject $app
                [void]$arrUpdated.Add($app.LocalizedDisplayName)
                Write-Host "done!`n" -ForegroundColor Green
            }
            catch { write-host "Exception: $($_.Exception.Message)" -ForegroundColor Red }
        }
        else
        {
            Write-Host "Update parameter not provided. Not updating $($app.LocalizedDisplayName)..." -ForegroundColor DarkYellow
        }

        $counter = $true
    }
    else { Write-Host "Only valid references: $($app.LocalizedDisplayName)" -ForegroundColor Green }
    if ($appCounter -eq $arrApps.Count) { Write-Progress -Activity 'Verifying supersedence' -Completed }
    $appCounter++
}

if (!$counter) {Write-Host "`nNo orphaned superseded application found." -ForegroundColor Green}
if ($arrUpdated) 
{
    $arrUpdated | Out-File "$env:windir\temp\UpdatedApplications.txt"
    Write-Host "List of updated applications exported to '$env:windir\temp\UpdatedApplications.txt'"
}

Write-Host
Stop-Transcript

Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0 # Remove declared variables