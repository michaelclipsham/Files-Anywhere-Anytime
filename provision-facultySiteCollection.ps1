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
		[Param(Mandatory)]
		[String]
		$SchoolCode,
		[Param(Mandatory)]
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
		#$sgOwners = "$($SiteTitle) Owners"

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
		Start-Sleep -Milliseconds 2000 # Just pausing for effect... nah kidding, there might need to be a short delay to allow the SPOUser object be created
		ProvisionSiteCollectionAdmins $siteUrl "Cloud Migration Project Support"
		ProvisionSiteCollectionAdmins $siteUrl $($schoolShortName + " - School.Principal")
		ProvisionSiteCollectionAdmins $siteUrl $($schoolShortName + " - School.RelPrincipal") ##"~SCH$($schoolcode)SRP"

		Write-Host "Complete - Provisioning Site Permissions" -ForegroundColor Green
	}
	catch [System.Exception] {
		Write-Warning $_.Exception.ToString() 
	}
}