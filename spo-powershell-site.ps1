#-----------------------------------------------------
# Creates a site collection if not exists
#-----------------------------------------------------
function CreateTeamSite() {
	param(
		[Parameter(Mandatory)]
		[String]
		$SiteTemplate,
		[Parameter(Mandatory)]
		[String]
		$SiteUrl,
		[Parameter(Mandatory)]
		[String]
		$SiteTitle,		
		[Parameter(Mandatory)]
		[String]
		$SiteOwner,
		[Parameter(Mandatory)]
		[String]
		$FacultyStorageQuota
	)

	$site = $null
	Write-Host "Start - Create Site Collection $($SiteUrl)" -ForegroundColor Green
	try {
		$site = Get-SPOSite -Identity $SiteUrl -ErrorAction SilentlyContinue
	}
	catch {
		Write-Host "No site exists for $($SiteUrl)"
	}

	if ($null -eq $site) {
		Write-Host "Creating new site..."
		New-SPOSite -Template $SiteTemplate -Title $SiteTitle -Owner $SiteOwner -Url $SiteUrl -StorageQuota $FacultyStorageQuota | Out-Null
	}
	else {
		Write-Warning "A Site $($SiteUrl) exists already. Skipping..."
	}
	Write-Host "Complete - Create Site Collection $($SiteUrl)" -ForegroundColor Green
}