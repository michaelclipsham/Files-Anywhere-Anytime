# Refer to https://gist.github.com/wictorwilen/db67725a66a3e40789e3#file-apponly-acs-powershell-sample-ps1-L13
function CatalogCreatedSitesForSchool() {
	param(
		[Parameter(Mandatory)]
		[String]
		$SchoolShortName,
		[Parameter(Mandatory)]
		$SchoolUrls,
        [Parameter(Mandatory)]
        [String]
        $ClientId,
        [Parameter(Mandatory)]
        [String]
        $ClientSecret,
        [Parameter(Mandatory)]
        [String]
        $RedirectUri,
        [Parameter(Mandatory)]
        [String]
        $TenantId,
        [Parameter(Mandatory)]
        [String]
        $TenantDomain,
        [Parameter(Mandatory)]
        [String]
        $CatalogSiteUrl,
        [Parameter(Mandatory)]
        [String]
        $CatalogListName,
        [Parameter(Mandatory)]
        [String]
        $CatalogListItemType
	)
    $SPOIdentifier = "00000003-0000-0ff1-ce00-000000000000"
    [System.Reflection.Assembly]::LoadWithPartialName("System.Web") | Out-Null
    $authNBody = "grant_type=client_credentials"
    $authNBody += "&client_id=" +[System.Web.HttpUtility]::UrlEncode($ClientId + "@" + $TenantId)
    $authNBody += "&client_secret=" +[System.Web.HttpUtility]::UrlEncode($ClientSecret)
    $authNBody += "&redirect_uri=" +[System.Web.HttpUtility]::UrlEncode($RedirectUri)
    $authNBody += "&resource=" +[System.Web.HttpUtility]::UrlEncode($SPOIdentifier + "/" + $TenantDomain + "@" + $TenantId)

    $authN = Invoke-WebRequest -Uri "https://accounts.accesscontrol.windows.net/$TenantId/tokens/OAuth/2" `
        -Method Post `
        -Body $authNBody `
        -ContentType "application/x-www-form-urlencoded"
    $authNJson = $authN.Content | ConvertFrom-Json

    $headers = @{
        "Authorization" = "Bearer " + $authNJson.access_token;
        "Accept" = "application/json;odata=verbose";
        "Content-Type" = "application/json;odata=verbose"
    }

    $newItemBody  = '{';
    $newItemBody += '    "__metadata": {';
    $newItemBody += '        "type": "' + $CatalogListItemType + '"';
    $newItemBody += '    },';
    $newItemBody += '    "Title": "' + $SchoolShortName + '",';
    $newItemBody += '    "Principal": "' + $SchoolUrls["Principal"] + '",';
    $newItemBody += '    "Executive": "' + $SchoolUrls["Executive"] + '",';
    $newItemBody += '    "Office": "' +    $SchoolUrls["Office"]    + '",';
    $newItemBody += '    "SchoolAssistance": "' + $SchoolUrls["School Assistance"] + '",';
    $newItemBody += '    "StaffInformation": "' + $SchoolUrls["Staff Information"] + '",';
    $newItemBody += '    "Local1": "' +    $SchoolUrls["Local1"]    + '",';
    $newItemBody += '    "Local2": "' +    $SchoolUrls["Local2"]    + '",';
    $newItemBody += '    "Teacher": "' +   $SchoolUrls["Teacher"]   + '"';
    $newItemBody += '}';
    Invoke-RestMethod -Uri "$($CatalogSiteUrl)/_api/lists/GetByTitle(`'$CatalogListName`')/Items" -Method Post -Headers $headers -Body $newItemBody | Out-Null
    Write-Host "Created catalog record for $($SchoolShortName)"
}