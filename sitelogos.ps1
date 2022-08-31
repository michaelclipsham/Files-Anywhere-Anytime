#-----------------------------------------------------
# Upload and set the website logo
# Note: this function depends on PnP.PowerShell
#-----------------------------------------------------
function UploadSetSiteLogo {
	param(
		[Parameter(Mandatory)]
		[String]
		$SiteUrl,
		[Parameter(Mandatory)]
		[String]
		$SiteRole,
        [Parameter(Mandatory)]
        [String]
        $LogoFolderPath
	)

    #Initialise a few things, connect to the SharePoint website
    $LogoFileName = "64x64_" + $SiteRole + ".jpg"
    $LogoFilePath = $LogoFolderPath + "\" + $LogoFileName
    $SiteAssetsLibraryName = "Site Assets"
    $SiteAssetsLibraryURL = "SiteAssets"
    $LogoURL = $SiteUrl + "/" + $SiteAssetsLibraryURL + "/" + $LogoFileName
    Connect-PnPOnline -Url $SiteUrl -Interactive ##hopefully that's not a prompt each and every website
    
    #Ensure Site Assets library
    $SiteAssets = Get-PnPList -Identity $SiteAssetsLibraryName -ErrorAction SilentlyContinue
    If ($null -eq $SiteAssets) {
        New-PnPList -Title $SiteAssetsLibraryName -Template DocumentLibrary -Url $SiteAssetsLibraryURL
    }

    #Upload logo to Site Assets library
    Add-PnPFile -Path $LogoFilePath -Folder $SiteAssetsLibraryURL

    #Set site logo
    Set-PnPWeb -SiteLogoUrl $LogoURL
}