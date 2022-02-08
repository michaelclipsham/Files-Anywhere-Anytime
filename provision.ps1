#========================================================================
#                         MAIN SCRIPT
#========================================================================
. "$PSScriptRoot\provision-facultySiteCollection.ps1"
. "$PSScriptRoot\spo-powershell-logging.ps1"
. "$PSScriptRoot\spo-powershell-security.ps1"
. "$PSScriptRoot\spo-powershell-site.ps1"
. "$PSScriptRoot\spo-powershell-siteTemplate.ps1"
. "$PSScriptRoot\catalog-sites.ps1"

#-----------------------------------------------------
# Load script configuration file
#-----------------------------------------------------
<###
 ### Expects:
	RUNCREATEFACULTYSITECOLLECTION=true|false
	SITETENANTURL=https://<tenantname>-admin.sharepoint.com
	SITEROOTURL=https://<tenantname.sharepoint.com
	PROVISIONINGUSER=<SharePoint administrator UPN>
	USEMULTIFACTORAUTH=true|false
	LOGFOLDER=<path to logs>
	FACULTYSITETEMPLATE=STS#3
	SITEDESIGNSCRIPTPATH=\assets\site-design-script.json
	SCHOOLSTOPROVISIONCSVPATH=\assets\doe-schools-provisioning.csv
	SITESTORAGEQUOTAMB=1024000
	MANAGEDPATH=sites
	CATALOGSITEURL=https://<tenantname>.sharepoint.com/<managedpath>/<siteurl>
	CATALOGLISTNAME=<list name>
	CATALOGLISTCLIENTID=<SPO app-only context client GUID>
	CATALOGLISTCLIENTSECRET=<SPO app-only context client secret>
	CATALOGLISTCLIENTREDIRECTURI=<SPO app-only context client redirect Uri>
	TENANTID=<Tenant GUID>
	TENANTDOMAIN=<tenantname>.sharepoint.com
	CATALOGLISTITEMTYPE=SP.Data.<CatalogListName>ListItem
 ### Optional:
 	SITEDESIGNID=<Site Design GUID>
####>
$scriptConfigPath = Join-Path $PSScriptRoot "\assets\script.config"
foreach ($keyValuePair in $(Get-Content $scriptConfigPath)){
    Set-Variable -Name $keyValuePair.split("=")[0] -Value $keyValuePair.split("=",2)[1]
}

# Faculty Site Settings:
### COME BACK TO THIS GUY
#$provisioningUser = $user
$facultySiteInfo = @(
	[PSCustomObject]@{Title = "Principal";         Url = "Principal" };
	[PSCustomObject]@{Title = "Executive";         Url = "Executive" };
	[PSCustomObject]@{Title = "Office";            Url = "Office" };
	[PSCustomObject]@{Title = "School Assistance"; Url = "Assistance" };
	[PSCustomObject]@{Title = "Staff Information"; Url = "StaffInfo" };
	[PSCustomObject]@{Title = "Local1";            Url = "Local1" };
	[PSCustomObject]@{Title = "Local2";            Url = "Local2" };
	[PSCustomObject]@{Title = "Teacher";           Url = "Teacher" }
)
$facultySiteUrls = @{}

#-----------------------------------------------------
# Resource File Paths
#-----------------------------------------------------
$JSONScriptPath = Join-Path $PSScriptRoot $SITEDESIGNSCRIPTPATH
$csvFilePath = Join-Path $PSScriptRoot $SCHOOLSTOPROVISIONCSVPATH

#-----------------------------------------------------
# Logging
#-----------------------------------------------------
$logFilePath = Join-Path $LOGFOLDER "log-doe-site-provisioning-$([datetimeoffset]::Now.ToString("yyyyMMddTHHmm")).txt"
if (!(Test-Path -Path $LOGFOLDER)) {
	New-Item -Path $LOGFOLDER -ItemType Directory
}

#-----------------------------------------------------
# Start provisioning
#-----------------------------------------------------
Start-Transcript -Path $logFilePath
$start = [DateTimeOffset]::Now
Write-Host "Start - $([datetimeoffset]::Now.ToString("yyyy-MM-dd HHmm"))"

Write-Host "Start - Provisioning Site Collection" -ForegroundColor Green
$runId = New-Guid
Write-Host "Run Id - ($runId)"

ConnectSPOOnlineTenant -SiteTenantUrl $SITETENANTURL -UseMultiFactorAuth ([System.Convert]::ToBoolean($USEMULTIFACTORAUTH))

# Create Site Template
if ($SITEDESIGNID) {
	# A site design has been designated, no need to re-create it
	$siteDesign = [PSCustomObject]@{
		Id=$SITEDESIGNID
	}
}
else {
	# Provision a new site design for school faculty drive
	$siteDesign = CreateSiteDesign -ScriptPath $JSONScriptPath
}

#region CSV validation
# Validate CSV File Headers
Write-Host "Validating CSV file $($csvFilePath)."
$correctHeaders = @("School Short Name", "School Code")
$csvFileHeaders = (Get-Content $csvFilePath -TotalCount 1).Split(",")
if ($csvFileHeaders.Length -ne $correctHeaders.Length) {
	Write-Error "The number of headers in the csv file '$($csvFilePath)' is invalid."
}
for ($i = 0; $i -lt $csvFileHeaders.Count; $i++) {
	if ($correctHeaders -notcontains $csvFileHeaders[$i].TrimStart()) {
		Write-Error "The csv file used is invalid.  Incorrect header '$($csvFileHeaders[$i].TrimStart())' found in csv file."
	}
}

# Validate content of CSV file rows
# TODO
#endregion

$schoolsToProvisionFile = Import-Csv -Path $csvFilePath

foreach ($row in $schoolsToProvisionFile) {	
	
	$schoolShortName = $row."School Short Name";
	$schoolCode = $row."School Code";
	$facultySiteUrls = @{}
	
	AddAuditLog -RunId $runId -shortName $schoolShortName -code $schoolCode -eventMessage "Provisioning Run Started"

	# Create Site Collections
	foreach ($facultySite in $facultySiteInfo) {
		$facultySiteUrl = "$($SITEROOTURL)/$MANAGEDPATH/$schoolCode-$($facultySite.Url)"
		$facultySiteTitle = "$($schoolShortName)-$($facultySite.Title)"
		$facultySiteAlias = "$schoolCode-$($facultySite.Url)"
		if ([System.Convert]::ToBoolean($RUNCREATEFACULTYSITECOLLECTION)){
			CreateFacultySiteCollection -SiteUrl $facultySiteUrl -SiteTitle $facultySiteTitle -SiteOwner $PROVISIONINGUSER -TeamSiteAlias $facultySiteAlias -FacultySiteTemplate $FACULTYSITETEMPLATE -FacultyStorageQuota $SITESTORAGEQUOTAMB
		}
		$facultySiteUrls.Add($facultySite.Title, $facultySiteUrl);
	}
	
	# Create a record in the catalog of created SharePoint Faculty Drives
	CatalogCreatedSitesForSchool `
		-SchoolShortName $schoolShortName `
		-SchoolUrls $facultySiteUrls `
		-ClientId $CATALOGLISTCLIENTID `
		-ClientSecret $CATALOGLISTCLIENTSECRET `
		-RedirectUri $CATALOGLISTCLIENTREDIRECTURI `
		-TenantId $TENANTID `
		-TenantDomain $TENANTDOMAIN `
		-CatalogSiteUrl $CATALOGSITEURL `
		-CatalogListName $CATALOGLISTNAME `
		-CatalogListItemType $CATALOGLISTITEMTYPE

	# Provision Site Collections - site design and permissions
	foreach ($facultySite in $facultySiteInfo) {
		$facultySiteUrl = "$($SITEROOTURL)/$MANAGEDPATH/$schoolCode-$($facultySite.Url)"
		$facultySiteTitle = "$($schoolShortName)-$($facultySite.Title)"
		$facultySiteAlias = "$schoolCode-$($facultySite.Url)"
		ProvisionFacultySiteCollection -SiteUrl $facultySiteUrl -SiteTitle $facultySiteTitle -SiteOwner $PROVISIONINGUSER -TeamSiteAlias $facultySiteAlias -SiteDesign $siteDesign -SchoolCode $schoolCode -SchoolShortName $schoolShortName
	}

	# Change owner, revoke SharePoint administrator access
	foreach ($facultySite in $facultySiteInfo) {
		$schoolPrincipalADGroupMail = "~SCH"+$schoolCode+"SP@det.nsw.edu.au"
		$facultySiteUrl = "$($SITEROOTURL)/$MANAGEDPATH/$schoolCode-$($facultySite.Url)"
		SetSiteCollectionOwner -SiteUrl $facultySiteUrl -OwnerClaim $schoolPrincipalADGroupMail
	}
	
	AddAuditLog -RunId $runId -shortName $schoolShortName -code $schoolCode -eventMessage "Provisioning Run Completed"
	
}

Write-Host "Completed - Provisioning Site Collection" -ForegroundColor Green
$end = [DateTimeOffset]::Now
$taken = $end - $start
Write-Host ""
Write-Host "Start: $($start.ToString("yyyy'-'MM'-'dd HH':'mm zzz"))"
Write-Host "End:   $($end.ToString("yyyy'-'MM'-'dd HH':'mm zzz"))"
Write-Host "Taken: $($taken.ToString())"
Stop-Transcript