
#-----------------------------------------------------
# Creating Site Script template
#-----------------------------------------------------

function CreateSiteDesign {
	param (
		[Parameter(Mandatory)]
		$ScriptPath
	)
	# Add Site Script and Site Design to Tenant
	# https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-json-schema
	
	$siteScriptTitle = "DoE Faculty Site Script"
	$siteDesignTitle = "DoE Faculty Site Template"

	Write-Host "Start - Create Site Template in Tenant" -ForegroundColor Green

	$siteScript = Get-SPOSiteScript | Where-Object { $_.Title -eq $siteScriptTitle } 
	if ($null -eq $siteScript) {
		Write-Host "  Adding Site Script '$($siteScriptTitle)' to Tenant."
		$siteScript = (Get-Content $ScriptPath -Raw | Add-SPOSiteScript -Title $siteScriptTitle)		
	}
	else {
		Write-Warning "A Site Script with the title $($siteScriptTitle) already exists. Skipping..."
	}

	$siteDesign = Get-SPOSiteDesign | Where-Object { $_.Title -eq $siteDesignTitle }
	if ($null -eq $siteDesign) {
		Write-Host "  Adding Site Design '$($siteDesign)' to Tenant."
		$webTemplate = "1" #64 = Team Site, 68 = Communication Site, 1 = Team Site without Group
		$siteDesign = Add-SPOSiteDesign -Title $siteDesignTitle -WebTemplate $webTemplate -SiteScripts $siteScript.Id -Description "DoE Faculty Site Template"

	}
	else {
		Write-Warning "A Site Design with the title $($siteDesignTitle) already exists. Skipping..."
	}

	Write-Host "Complete - Create Site Template in Tenant" -ForegroundColor Green

	return $siteDesign	
}
