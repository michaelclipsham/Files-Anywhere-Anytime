#-----------------------------------------------------
# Provisions site collection - security
#-----------------------------------------------------
function ProvisionSiteSecurityGroups($siteUrl, $group, $permissions) {
	
	Write-Host "Start - Updating permission levels on $($group.Title)"
	# Remove all permission levels for the group
	try {
		$roles = GetRoles -Group $group.Title -SiteUrl $siteUrl
		if ($roles -contains "Full Control") {
			Set-SPOSiteGroup -Site $siteUrl -Identity $group.Title -PermissionLevelsToRemove "Full Control" -ErrorAction SilentlyContinue
			Write-Host "Removed  'Full Control' permission levels from group '$($group.Title)'"
		}
		if ($roles -contains "Contribute") {
			Set-SPOSiteGroup -Site $siteUrl -Identity $group.Title -PermissionLevelsToRemove "Contribute" -ErrorAction SilentlyContinue
			Write-Host "Removed  'Contribute' permission levels from group '$($group.Title)'"
		}
		if ($roles -contains "Design") {
			Set-SPOSiteGroup -Site $siteUrl -Identity $group.Title -PermissionLevelsToRemove "Design" -ErrorAction SilentlyContinue
			Write-Host "Removed  'Design' permission levels from group '$($group.Title)'"
		}
		if ($roles -contains "Edit") {
			Set-SPOSiteGroup -Site $siteUrl -Identity $group.Title -PermissionLevelsToRemove "Edit" -ErrorAction SilentlyContinue
			Write-Host "Removed  'Edit' permission levels from group '$($group.Title)'"
		}
		if ($roles -contains "Read") {
			Set-SPOSiteGroup -Site $siteUrl -Identity $group.Title -PermissionLevelsToRemove "Read" -ErrorAction SilentlyContinue
			Write-Host "Removed  'Read' permission levels from group '$($group.Title)'"
		}
	} catch { 
		Write-Warning "Provision web security $($group.Title) failed. Error: $($Error[0])"
	}

	# Add the specified permission levels to the group
	try {
		Set-SPOSiteGroup -Site $siteUrl -Identity $group.Title -PermissionLevelsToAdd $permissions
		Write-Host "Added '$([system.String]::Join(", ", $permissions))' permission levels from group '$($group.Title)'"
	} catch { 
		Write-Warning "Provision web security $($group.Title) failed. Error: $($Error[0])"
	}

	Write-Host "Complete - Updating permission levels on $($group.Title)"
}

function GetRoles() {
	param(
		[Parameter(Mandatory)]
		[String]
		$SiteUrl,
		[Parameter(Mandatory)]
		$Group
	)

	try {
		$group = Get-SPOSiteGroup -Group $Group -Site $SiteUrl
		return $group.Roles
	}
	catch {
		Write-Warning "Error: $($Error[0])"
	}
}
function ProvisionSiteSecurityGroupMembers() {
	param(
		[Parameter(Mandatory)]
		[String]
		$SiteUrl,
		[Parameter(Mandatory)]
		$group,
		[Parameter(Mandatory)]
		[String]
		$GroupDisplayName
	)

	try {
		# add the members to the group
		Add-SPOUser -Site $SiteUrl -LoginName $GroupDisplayName -Group $group
		Write-Host "Added '$($GroupDisplayName)' to '$($group)' security group"
	}
	catch { 
		Write-Warning "Adding '$($GroupDisplayName)' to '$($group)' security group failed. Error: $($Error[0])"
	}
}

function ProvisionSiteCollectionAdmins() {
	param(
		[Parameter(Mandatory)]
		[String]
		$SiteUrl,
		[Parameter(Mandatory)]
		[String]
		$GroupDisplayName
	)

	try {
		# Add Site Collection Admins
		$adGroup = GetADSecurityGroup -DisplayName $GroupDisplayName -SiteUrl $SiteUrl
		$loginName = "c:0t`.c`|tenant`|$($adGroup.LoginName)"
		Set-SPOUser -Site $SiteUrl -LoginName $loginName -IsSiteCollectionAdmin $true
		Write-Host "Added '$($GroupDisplayName)' as a Site Collection Administrator"
	}
	catch { 
		Write-Warning "Adding '$($GroupDisplayName)' as a Site Collection Administrator failed. Error: $($Error[0])"
	}
}

function GetADSecurityGroup() {
	param(
		[Parameter(Mandatory)]
		[String]
		$DisplayName,
		[Parameter(Mandatory)]
		[String]
		$SiteUrl
	)

	try {
		return Get-SPOUser -Site $SiteUrl -Limit All | Where-Object { $_.IsGroup -and $_.DisplayName -eq $DisplayName } | Select-Object -First 1 
	}
	catch { 
		Write-Warning "Error: $($Error[0])"
	}
}

function ConnectSPOOnlineTenant() {
	param(
		[Parameter(Mandatory)]
		[bool]
		$UseMultifactorAuth,
		[Parameter(Mandatory)]
		[String]
		$SiteTenantUrl
	)
	try {
		Disconnect-SPOService -ErrorAction SilentlyContinue
	}
	catch { 
	}
    
	if ($useMultifactorAuth) {
		Connect-SPOService -Url $SiteTenantUrl
	}
	else {
		Connect-SPOService -Url $SiteTenantUrl -Credential (CreateCredentials) -ErrorAction SilentlyContinue		
 }	
}

function CreateCredentials {
	return New-Object System.Management.Automation.PSCredential ($user, $secpasswd)
}

function SetSiteCollectionOwner() {
	param(
		[Parameter(Mandatory)]
		[String]
		$SiteUrl,
		[Parameter(Mandatory)]
		[String]
		$OwnerClaim
	)

	try {
		Set-SPOSite -Identity $SiteUrl -Owner $OwnerClaim
		Write-Host "Set $($OwnerClaim) as Owner of $($SiteUrl)"
	}
	catch {
		Write-Warning "Adding '$($OwnerClaim)' as Site Collection Owner failed. Error: $($Error[0])"
	}
}

function RevokeFacultySiteCollectionPermission () {
	param(
		[Parameter(Mandatory)]
		[String]
		$SiteUrl,
		[Parameter(Mandatory)]
		[String]
		$IdentityClaim,
		[Parameter(Mandatory)]
		[String]
		$Group,
		[Parameter(Mandatory)]
		[bool]
		$IsGroup
	)
	
	try {
		if ($IsGroup) {
			$adGroup = GetADSecurityGroup -DisplayName $IdentityClaim -SiteUrl $SiteUrl
			$loginName = "c:0t`.c`|tenant`|$($adGroup.LoginName)"
			Remove-SPOUser -Site $SiteUrl -LoginName $loginName -Group $Group
		}
		else {
			Remove-SPOUser -Site $SiteUrl -LoginName $IdentityClaim -Group $Group
		}
		Write-Host "Removed $($IdentityClaim) from $($Group) in $($SiteUrl)"
	}
	catch {
		Write-Warning "Revoking $($IdentityClaim) from $($Group) in $(SiteUrl) failed. Error: $($Error[0])"
	}
}