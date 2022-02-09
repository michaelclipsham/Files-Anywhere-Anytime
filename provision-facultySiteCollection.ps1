#-----------------------------------------------------
# Provisions faculty site collections
#-----------------------------------------------------
function CreateFacultySiteCollection {	
	param(		
		[Parameter(Mandatory)]
		[String]
		$SiteUrl,		
		[Parameter(Mandatory)]
		[String]
		$SiteTitle,
		[Parameter(Mandatory)]
		[String]
		$SiteOwner,
		[String]
		$TeamSiteAlias,
		[Parameter(Mandatory)]
		[String]
		$FacultySiteTemplate,
		[Parameter(Mandatory)]
		[String]
		$FacultyStorageQuota
	)
	# Create Team Site
	try {
		Write-Host "Creating Faculty Site Collection for $($SiteTitle)"
		CreateTeamSite -SiteTemplate $FacultySiteTemplate -SiteUrl $SiteUrl -SiteTitle $SiteTitle -SiteOwner $SiteOwner -FacultyStorageQuota $FacultyStorageQuota
	}
	catch [System.Exception] {
		Write-Warning $_.Exception.ToString() 
	}
}

#-----------------------------------------------------
# Provisions faculty site collection site design and permissions
#-----------------------------------------------------
function ProvisionFacultySiteCollection {	
	param(		
		[Parameter(Mandatory)]
		[String]
		$SiteUrl,		
		[Parameter(Mandatory)]
		[String]
		$SiteTitle,
		[Parameter(Mandatory)]
		[String]
		$SiteOwner,
		[String]
		$TeamSiteAlias,
		[Parameter(Mandatory)]
		$SiteDesign,
		[Parameter(Mandatory)]
		[String]
		$SchoolCode,
		[Parameter(Mandatory)]
		[String]
		$SchoolShortName
	)
	try {
		#Apply Site Design
		Write-Host "Start - Applying Site Design to $($siteUrl)" -ForegroundColor Green
		Invoke-SPOSiteDesign -Identity $SiteDesign.Id -WebUrl $siteUrl
		Start-Sleep -Milliseconds 2000
		Write-Host "Complete - Applying Site Design to $($siteUrl)" -ForegroundColor Green

		# External Sharing - Off

		# Provision Site Permissions
		Write-Host "Start - Provisioning Site Permissions" -ForegroundColor Green
		$sgVisitors = "$($SiteTitle) Visitors"
		$sgMembers = "$($SiteTitle) Members"
		$sgOwners = "$($SiteTitle) Owners"

		# Revoke access for Members group and assign Contribute permission level
		Write-Host "Applying security to $($sgMembers)"
		$membersGroup = Get-SPOSiteGroup -Site $siteUrl -group $sgMembers
		ProvisionSiteSecurityGroups $siteUrl $membersGroup @('Contribute')

		# Add Security Groups to Site Groups
		ProvisionSiteSecurityGroupMembers -SiteUrl $siteUrl -group $sgMembers -GroupDisplayName "Cloud Migration Project Support"
		if ($facultySiteTitle -like "*Executive") {
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SE"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SP"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)DP"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SRP"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)RSE"
		}
		elseif ($facultySiteTitle -like "*Principal") {
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SP"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SRP"
		}
		elseif ($facultySiteTitle -like "*Local1") {
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SP"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)Local1"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SRP"
		}
		elseif ($facultySiteTitle -like "*Local2") {
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SP"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)Local2"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SRP"
		}
		elseif ($facultySiteTitle -like "*Office") {
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SP"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)Office"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)DP"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SRP"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)RSE"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SE"
		}
		elseif ($facultySiteTitle -like "*School Assistance") {
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SP"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)Office"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)DP"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgVisitors "~SCH$($schoolcode)All"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SE"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)Assistance"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SRP"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)RSE"
		}
		elseif ($facultySiteTitle -like "*Staff Information") {
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SP"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)Office"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)DP"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgVisitors "~SCH$($schoolcode)All"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SE"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SRP"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)RSE"
		}
		elseif ($facultySiteTitle -like "*Teacher") {
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SP"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)All"
			ProvisionSiteSecurityGroupMembers $siteUrl $sgMembers "~SCH$($schoolcode)SRP"
		}

		# Add Security Groups to Site Collection Administrator
		ProvisionSiteCollectionAdmins $siteUrl "Cloud Migration Project Support"
		ProvisionSiteCollectionAdmins $siteUrl $($schoolShortName + " - School.Principal")
		ProvisionSiteCollectionAdmins $siteUrl $($schoolShortName + " - School.RelPrincipal") ##"~SCH$($schoolcode)SRP"

		# Remove Project Support Group from Members group
		RevokeFacultySiteCollectionPermission -SiteUrl $siteUrl -IdentityClaim "Cloud Migration Project Support" -Group $sgMembers -IsGroup $true
		# Remove provisioning user from Owners group, which also strangely does site collection admin at the same time
		RevokeFacultySiteCollectionPermission -SiteUrl $siteUrl -IdentityClaim $SiteOwner -Group $sgOwners -IsGroup $false

		Write-Host "Complete - Provisioning Site Permissions" -ForegroundColor Green
	}
	catch [System.Exception] {
		Write-Warning $_.Exception.ToString() 
	}
}

#-----------------------------------------------------
# Configure faculty document library for UX naming conventions
#-----------------------------------------------------
function ConfigureDefaultDocumentLibrary {
	param(
		[Parameter(Mandatory)]
		[String]
		$SiteUrl,
		[Parameter(Mandatory)]
		[String]
		$DocumentLibraryName,
		[Parameter(Mandatory)]
		[String]
		$SchoolShortName
	)

	$siteScriptContent = @{}
	$siteScriptContent.Add("`$schema", "https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json")
	$verbs = @()
	$verb = @{}
	$verb.Add("verb", "setTitle")
	$verb.Add("title", $DocumentLibraryName)
	$verbs += $verb
	$siteScriptContent.Add("actions", $verbs)
	$siteScriptContent.Add("version", 1)
	
	Write-Host "Adding site script and design to rename default document library on $SchoolShortName to $DocumentLibraryName"
	$siteScript = $siteScriptContent | Add-SPOSiteScript -Title $( $SchoolShortName + " Document library rename site script" )
	$siteDesign = Add-SPOSiteDesign -Title $( $SchoolShortName + " Document library rename site design" ) -WebTemplate "1" -SiteScripts $siteScript.Id -Description "Renames default document library"
	Invoke-SPOSiteDesign -Identity $siteDesign -WebUrl $SiteUrl
	Write-Host "Invoked site design to rename default document library, pausing for design to be applied..."

	Start-Sleep -Seconds 5
	Write-Host "Removing site script and design to rename default document library"
	Remove-SPOSiteDesign -Identity $siteDesign.Id
	Remove-SPOSiteScript -Identity $siteScript.Id
	Write-Host "Completed renaming default document library"
}