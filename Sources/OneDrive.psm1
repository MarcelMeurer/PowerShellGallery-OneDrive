function Get-ODAuthentication
{
	<#
	.DESCRIPTION
	Connect to OneDrive for authentication with a given client id (get your free client id on https://apps.dev.microsoft.com) For a step-by-step guide: https://github.com/MarcelMeurer/PowerShellGallery-OneDrive
	.PARAMETER ClientId
	ClientId of your "app" from https://apps.dev.microsoft.com
	.PARAMETER AppKey
	The client secret for your OneDrive "app". If AppKey is set the authentication mode is "code." Code authentication returns a refresh token to refresh your authentication token unattended.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.PARAMETER Scope
	Comma-separated string defining the authentication scope (https://dev.onedrive.com/auth/msa_oauth.htm). Default: "onedrive.readwrite,offline_access". Not needed for OneDrive 4 Business access.
	.PARAMETER RefreshToken
	Refreshes the authentication token unattended with this refresh token. 
	.PARAMETER AutoAccept
	In token mode the accept button in the web form is pressed automatically.
	.PARAMETER RedirectURI
	Code authentication requires a correct URI. Use the same as in the app registration e.g. http://localhost/logon. Default is https://login.live.com/oauth20_desktop.srf. Don't use this parameter for token-based authentication. 
	.PARAMETER DontShowLoginScreen
	Suppresses the logon screen. Be careful: If you suppress the logon screen you cannot logon if your credentials are not passed through. 
	.PARAMETER LogOut
	Performs a logout. 

	.EXAMPLE
    $Authentication=Get-ODAuthentication -ClientId "0000000012345678"
	$AuthToken=$Authentication.access_token
	Connect to OneDrive for authentication and save the token to $AuthToken
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$True)]
		[string]$ClientId = "unknown",
		[string]$Scope = "onedrive.readwrite,offline_access",
		[string]$RedirectURI ="https://login.live.com/oauth20_desktop.srf",
		[string]$AppKey="",
		[string]$RefreshToken="",
		[string]$ResourceId="",
		[switch]$DontShowLoginScreen=$false,
		[switch]$AutoAccept,
		[switch]$LogOut,
		[switch]$enableglobal
	)
	$optResourceId=""
	$optOauthVersion="/v2.0"
	if ($ResourceId -ne "")
	{
		write-debug("Running in OneDrive 4 Business mode")
		$optResourceId="&resource=$ResourceId"
		$optOauthVersion=""
	}
	$Authentication=""
	if ($AppKey -eq "")
	{ 
		$Type="token"
	} else 
	{ 
		$Type="code"
	}
	if ($RefreshToken -ne "")
	{
		write-debug("A refresh token is given. Try to refresh it in code mode.")
		$body="client_id=$ClientId&redirect_URI=$RedirectURI&client_secret=$([uri]::EscapeDataString($AppKey))&refresh_token="+$RefreshToken+"&grant_type=refresh_token"
		write-host $body
		$webRequest=Invoke-WebRequest -Method POST -Uri "https://login.microsoftonline.com/common/oauth2$optOauthVersion/token" -ContentType "application/x-www-form-URLencoded" -Body $Body -UseBasicParsing
		$Authentication = $webRequest.Content |   ConvertFrom-Json
	} else
	{
		write-debug("Authentication mode: " +$Type)
		$ErrorActionPreference="SilentlyContinue"
		$null=[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
		$Verify= $?
		[Reflection.Assembly]::LoadWithPartialName("System.Drawing") | out-null
		[Reflection.Assembly]::LoadWithPartialName("System.Web") | out-null
		if ($Logout)
		{
			$URIGetAccessToken="https://login.live.com/logout.srf"
		}
		else
		{
			if ($ResourceId -ne "")
			{
				# OD4B
				$URIGetAccessToken="https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&client_id=$ClientId&redirect_URI=$RedirectURI"
			}
			else
			{
				# OD private
				$URIGetAccessToken="https://login.live.com/oauth20_authorize.srf?client_id="+$ClientId+"&scope="+$Scope+"&response_type="+$Type+"&redirect_URI="+$RedirectURI
			}
		}
		if($Verify){
		$form = New-Object Windows.Forms.Form
		$form.text = "Authenticate to OneDrive"
		$form.size = New-Object Drawing.size @(700,600)
		$form.Width = 675
		$form.Height = 750
		$web=New-object System.Windows.Forms.WebBrowser
		$web.IsWebBrowserContextMenuEnabled = $true
		$web.Width = 600
		$web.Height = 700
		$web.Location = "25, 25"
		$web.navigate($URIGetAccessToken)
		$DocComplete  = {
			if ($web.Url.AbsoluteUri -match "access_token=|error|code=|logout") {$form.Close() }
			if ($web.DocumentText -like '*ucaccept*') {
				if ($AutoAccept) {$web.Document.GetElementById("idBtn_Accept").InvokeMember("click")}
			}
		}
		$web.Add_DocumentCompleted($DocComplete)
		$form.Controls.Add($web)
		if ($DontShowLoginScreen -or -not($Verify))
		{
			write-debug("Logon screen suppressed by flag -DontShowLoginScreen")
			$form.Opacity = 0.0;
		}
		$form.showdialog() | out-null
		}else{
		@("A refresh token is given. Try to refresh it in code mode.",$URIGetAccessToken)|out-host
		$regex= 'access_token=[^&]+'
		do {
		$key=read-host -prompt 'URL'
		if (-not($key -match $regex)){write-host 'ERROR' -ForegroundColor DarkCyan|out-host}
		}until($key -match $regex)
		$Global:ODAccessToken=($key|select-string -pattern '(?<=token=)[^&]+' ).matches.value
		$Global:odtokentime=get-date
		$web=[PSobject]@{Url=$key}
		}
		# Build object from last URI (which should contains the token)
		$ReturnURI=($web.Url).ToString().Replace("#","&")
		if ($LogOut) {return "Logout"}
		if ($Type -eq "code")
		{
			write-debug("Getting code to redeem token")
			$Authentication = New-Object PSObject
			ForEach ($element in $ReturnURI.Split("?")[1].Split("&")) 
			{
				$Authentication | add-member Noteproperty $element.split("=")[0] $element.split("=")[1]
			}
			if ($Authentication.code)
			{
				$body="client_id=$ClientId&redirect_URI=$RedirectURI&client_secret=$([uri]::EscapeDataString($AppKey))&code="+$Authentication.code+"&grant_type=authorization_code"+$optResourceId
				$webRequest=Invoke-WebRequest -Method POST -Uri "https://login.microsoftonline.com/common/oauth2$optOauthVersion/token" -ContentType "application/x-www-form-urlencoded" -Body $Body -UseBasicParsing
				$Authentication = $webRequest.Content |   ConvertFrom-Json
			} else
			{
				write-error("Cannot get authentication code. Error: "+$ReturnURI)
			}
		} else
		{
			$Authentication = New-Object PSObject
			ForEach ($element in $ReturnURI.Split("?")[1].Split("&")) 
			{
				$Authentication | add-member Noteproperty $element.split("=")[0] $element.split("=")[1]
			}
			if ($Authentication.PSobject.Properties.name -match "expires_in")
			{
				$Authentication | add-member Noteproperty "expires" ([System.DateTime]::Now.AddSeconds($Authentication.expires_in))
			}
		}
	}
	if (!($Authentication.PSobject.Properties.name -match "expires_in"))
	{
		write-warning("There is maybe an errror, because there is no access_token!")
	}
	$Authentication | add-member Noteproperty "ClientId" ($ClientId)
	if ($ResourceId){
	$Authentication | add-member Noteproperty "ResourceId" ($ResourceId)
	}
	if($enableglobal){
	$Global:Authentication=$Authentication
	}
	return $Authentication 
}
function Get-ODRootUri 
{
	PARAM(
		[String]$ResourceId=""
	)
	if ($ResourceId -ne "")
	{
		return $ResourceId+"_api/v2.0"
	}
	else
	{
		return "https://api.onedrive.com/v1.0"
	}
}

function Get-ODWebContent 
{
	<#
	.DESCRIPTION
	Internal function to interact with the OneDrive API
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.PARAMETER rURI
	Relative path to the API.
	.PARAMETER Method
	Webrequest method like PUT, GET, ...
	.PARAMETER Body
	Payload of a webrequest.
	.PARAMETER BinaryMode
	Do not convert response to JSON.
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$false)]
		[string]$AccessToken,
		[String]$ResourceId="",
		[string]$rURI = "",
		[ValidateSet("PUT","GET","POST","PATCH","DELETE")] 
        [String]$Method="GET",
		[String]$Body,
		[switch]$BinaryMode
	)
	if ($Body -eq "") 
	{
		$xBody=$null
	} else
	{
		$xBody=$Body
	}
	
	$ODRootURI=Get-ODRootUri -ResourceId $ResourceId
	$string=@()
	$string=@{ enableglobal=$true}
	if($Authentication -and -not($ResourceId)){$string=@{ClientId = $Authentication.ClientId}} 
	if($Authentication.ResourceId -and -not($ResourceId)){$string=@{ResourceId=$Authentication.ResourceId}}elseif($ResourceId){$string=@{ResourceId=$ResourceId}}
do{
$doCount++
if($doCount -ne 1){
$null=Get-ODAuthentication @string
if(-not($?)){break}}
if (-not($AccessToken)){$AccessToken=$Authentication.access_token}
	try {
		$webRequest=Invoke-WebRequest -Method $Method -Uri ($ODRootURI+$rURI) -Header @{ Authorization = "BEARER "+$AccessToken} -ContentType "application/json" -Body $xBody -UseBasicParsing -ErrorAction SilentlyContinue
	} 
	catch
	{
		write-error("Cannot access the api. Webrequest return code is: "+$_.Exception.Response.StatusCode+"`n"+$_.Exception.Response.StatusDescription)
	}
		$errorcode=$?
}until($errorcode -or $doCount -ne 1)
if (-not($errorcode)){break}
	switch ($webRequest.StatusCode) 
    { 
        200 
		{
			if (!$BinaryMode) {$responseObject = ConvertFrom-Json $webRequest.Content}
			return $responseObject
		} 
        201 
		{
			write-debug("Success: "+$webRequest.StatusCode+" - "+$webRequest.StatusDescription)
			if (!$BinaryMode) {$responseObject = ConvertFrom-Json $webRequest.Content}
			return $responseObject
		} 
        204 
		{
			write-debug("Success: "+$webRequest.StatusCode+" - "+$webRequest.StatusDescription+" (item deleted)")
			$responseObject = "0"
			return $responseObject
		} 
        default {write-warning("Cannot access the api. Webrequest return code is: "+$webRequest.StatusCode+"`n"+$webRequest.StatusDescription)}
    }
}

function Get-ODDrives
{
	<#
	.DESCRIPTION
	Get user's drives.
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.EXAMPLE
    Get-ODDrives -AccessToken $AuthToken
	List all OneDrives available for your account (there is normally only one).
	.NOTES
	The application for OneDrive 4 Business needs "Read items in all site collections" on application level (API: Office 365 SharePoint Online)
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$false)]
		[string]$AccessToken,
		[String]$ResourceId=""
	)
	$ResponseObject=Get-ODWebContent -AccessToken $AccessToken -ResourceId $ResourceId -Method GET -rURI "/drives" 
	return $ResponseObject.Value
}

function Get-ODSharedItems
{
	<#
	.DESCRIPTION
	Get items shared with the user
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.EXAMPLE
    Get-ODDrives -AccessToken $AuthToken
	List all OneDrives available for your account (there is normally only one).
	.NOTES
	The application for OneDrive 4 Business needs "Read items in all site collections" on application level (API: Office 365 SharePoint Online)
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$false)]
		[string]$AccessToken,
		[String]$ResourceId=""
	)
	$ResponseObject=Get-ODWebContent -AccessToken $AccessToken -ResourceId $ResourceId -Method GET -rURI "/drive/oneDrive.sharedWithMe"
	return $ResponseObject.Value
}

function Format-ODPathorIdString
{
	<#
	.DESCRIPTION
	Formats a given path like '/myFolder/mySubfolder/myFile' into an expected URI format
	.PARAMETER Path
	Specifies the path of an element. If it is not given, the path is "/"
	.PARAMETER ElementId
	Specifies the id of an element. If Path and ElementId are given, the ElementId is used with a warning
	.PARAMETER DriveId
	Specifies the OneDrive drive id. If not set, the default drive is used
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[string]$Path="",
		[string]$DriveId="",
		[string]$ElementId=""
	)
	if (!$ElementId -eq "")
	{
		# Use ElementId parameters
		if (!$Path -eq "") {write-debug("Warning: Path and ElementId parameters are set. Only ElementId is used!")}
		$drive="/drive"
		if ($DriveId -ne "") 
		{	
			# Named drive
			$drive="/drives/"+$DriveId
		}
		return $drive+"/items/"+$ElementId
	}
	else
	{
		# Use Path parameter
		# replace some special characters
		$Path = ((((($Path -replace '%', '%25') -replace ' ', ' ') -replace '=', '%3d') -replace '\+', '%2b') -replace '&', '%26') -replace '#', '%23'
		# remove substring starts with "?"
		if ($Path.Contains("?")) {$Path=$Path.Substring(1,$Path.indexof("?")-1)}
		# replace "\" with "/"
		$Path=$Path.Replace("\","/")
		# filter possible string at the end "/children" (case insensitive)
		$Path=$Path+"/"
		$Path=$Path -replace "/children/",""
		# encoding of URL parts
		$tmpString=""
		foreach ($Sub in $Path.Split("/")) {$tmpString+=$Sub+"/"}
		$Path=$tmpString
		# remove last "/" if exist 
		$Path=$Path.TrimEnd("/")
		# insert drive part of URL
		if ($DriveId -eq "") 
		{	
			# Default drive
			$Path="/drive/root:"+$Path+":"
		}
		else
		{
			# Named drive
			$Path="/drives/"+$DriveId+"/root:"+$Path+":"
		}
		return ($Path).replace("root::","root")
	}
}

function Get-ODItemProperty
{
	<#
	.DESCRIPTION
	Get the properties of an item (file or folder).
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.PARAMETER Path
	Specifies the path to the element/item. If not given, the properties of your default root drive are listed.
	.PARAMETER ElementId
	Specifies the id of the element/item. If Path and ElementId are given, the ElementId is used with a warning.
	.PARAMETER SelectProperties
	Specifies a comma-separated list of the properties to be returned for file and folder objects (case sensitive). If not set, name, size, lastModifiedDateTime and id are used. (See https://dev.onedrive.com/odata/optional-query-parameters.htm).
	If you use -SelectProperties "", all properties are listed. Warning: A complex "content.downloadUrl" is listed/generated for download files without authentication for several hours.
	.PARAMETER DriveId
	Specifies the OneDrive drive id. If not set, the default drive is used.
	.EXAMPLE
    Get-ODItemProperty -AccessToken $AuthToken -Path "/Data/documents/2016/AzureML with PowerShell.docx"
	Get the default set of metadata for a file or folder (name, size, lastModifiedDateTime, id)

	Get-ODItemProperty -AccessToken $AuthToken -ElementId 8BADCFF017EAA324!12169 -SelectProperties ""
	Get all metadata of a file or folder by element id ("" select all properties)	
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$false)]
		[string]$AccessToken,
		[string]$ResourceId="",
		[string]$Path="/",
		[string]$ElementId="",
		[string]$SelectProperties="name,size,lastModifiedDateTime,id",
		[string]$DriveId=""
	)
	return Get-ODChildItems -AccessToken $AccessToken -ResourceId $ResourceId -Path $Path -ElementId $ElementId -SelectProperties $SelectProperties -DriveId $DriveId -ItemPropertyMode
}

function Get-ODChildItems
{
	<#
	.DESCRIPTION
	Get child items of a path. Return count is not limited.
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.PARAMETER Path
	Specifies the path of elements to be listed. If not given, the path is "/".
	.PARAMETER ElementId
	Specifies the id of an element. If Path and ElementId are given, the ElementId is used with a warning.
	.PARAMETER SelectProperties
	Specifies a comma-separated list of the properties to be returned for file and folder objects (case sensitive). If not set, name, size, lastModifiedDateTime and id are used. (See https://dev.onedrive.com/odata/optional-query-parameters.htm).
	If you use -SelectProperties "", all properties are listed. Warning: A complex "content.downloadUrl" is listed/generated for download files without authentication for several hours.
	.PARAMETER DriveId
	Specifies the OneDrive drive id. If not set, the default drive is used.
	.EXAMPLE
    Get-ODChildItems -AccessToken $AuthToken -Path "/" | ft
	Lists files and folders in your OneDrives root folder and displays name, size, lastModifiedDateTime, id and folder property as a table

    Get-ODChildItems -AccessToken $AuthToken -Path "/" -SelectProperties ""
	Lists files and folders in your OneDrives root folder and displays all properties
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$false)]
		[string]$AccessToken,
		[String]$ResourceId="",
		[string]$Path="/",
		[string]$ElementId="",
		[string]$SelectProperties="name,size,lastModifiedDateTime,id",
		[string]$DriveId="",
		[Parameter(DontShow)]
		[switch]$ItemPropertyMode,
		[Parameter(DontShow)]
		[string]$SearchText,
		[parameter(DontShow)]
        [switch]$Loop=$false
	)

	$ODRootURI=Get-ODRootUri -ResourceId $ResourceId
	if ($Path.Contains('$skiptoken=') -or $Loop)
	{	
		# Recursive mode of odata.nextLink detection
		write-debug("Recursive call")
		$rURI=$Path	
	}
	else
	{
		$rURI=Format-ODPathorIdString -path $Path -ElementId $ElementId -DriveId $DriveId
		$rURI=$rURI.Replace("::","")
		$SelectProperties=$SelectProperties.Replace(" ","")
		if ($SelectProperties -eq "")
		{
			$opt=""
		} else
		{
			$SelectProperties=$SelectProperties.Replace(" ","")+",folder"
			$opt="?select="+$SelectProperties
		}
		if ($ItemPropertyMode)
		{
			# item property mode
			$rURI=$rURI+$opt
		}
		else
		{
			if (!$SearchText -eq "") 
			{
				# Search mode
				$opt="/view.search?q="+$SearchText+"&select="+$SelectProperties
				$rURI=$rURI+$opt
			}
			else
			{
				# child item mode
				$rURI=$rURI+"/children"+$opt
			}
		}
	}
	write-debug("Accessing API with GET to "+$rURI)
	$ResponseObject=Get-ODWebContent -AccessToken $AccessToken -ResourceId $ResourceId -Method GET -rURI $rURI
	if ($ResponseObject.PSobject.Properties.name -match "@odata.nextLink") 
	{
		write-debug("Getting more elements form service (@odata.nextLink is present)")
		write-debug("LAST: "+$ResponseObject.value.count)
		Get-ODChildItems -AccessToken $AccessToken -ResourceId $ResourceId -SelectProperties $SelectProperties -Path $ResponseObject."@odata.nextLink".Replace($ODRootURI,"") -Loop
	}
	if ($ItemPropertyMode)
	{
		# item property mode
		return $ResponseObject
	}
	else
	{
		# child item mode
		return $ResponseObject.value
	}
}

function Search-ODItems
{
	<#
	.DESCRIPTION
	Search for items starting from Path or ElementId.
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.PARAMETER SearchText
	Specifies search string.
	.PARAMETER Path
	Specifies the path of the folder to start the search. If not given, the path is "/".
	.PARAMETER ElementId
	Specifies the element id of the folder to start the search. If Path and ElementId are given, the ElementId is used with a warning.
	.PARAMETER SelectProperties
	Specifies a comma-separated list of the properties to be returned for file and folder objects (case sensitive). If not set, name, size, lastModifiedDateTime and id are used. (See https://dev.onedrive.com/odata/optional-query-parameters.htm).
	If you use -SelectProperties "", all properties are listed. Warning: A complex "content.downloadUrl" is listed/generated for download files without authentication for several hours.
	.PARAMETER DriveId
	Specifies the OneDrive drive id. If not set, the default drive is used.
	.EXAMPLE
    Search-ODItems -AccessToken $AuthToken -Path "/My pictures" -SearchText "FolderA" 
	Searches for items in a sub folder recursively. Take a look at OneDrives API documentation to see how search (preview) works (file and folder names, in files, …)
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$false)]
		[string]$AccessToken,
		[String]$ResourceId="",
		[Parameter(Mandatory=$true,Position=0)]
		[string]$SearchText,
		[string]$Path="/",
		[string]$ElementId="",
		[string]$SelectProperties="name,size,lastModifiedDateTime,id",
		[string]$DriveId=""
	)
	return Get-ODChildItems -AccessToken $AccessToken -ResourceId $ResourceId -Path $Path -ElementId $ElementId -SelectProperties $SelectProperties -DriveId $DriveId -SearchText $SearchText	
}

function New-ODFolder
{
	<#
	.DESCRIPTION
	Create a new folder.
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.PARAMETER FolderName
	Name of the new folder.
	.PARAMETER Path
	Specifies the parent path for the new folder. If not given, the path is "/".
	.PARAMETER ElementId
	Specifies the element id for the new folder. If Path and ElementId are given, the ElementId is used with a warning.
	.PARAMETER DriveId
	Specifies the OneDrive drive id. If not set, the default drive is used.
	.EXAMPLE
    New-ODFolder -AccessToken $AuthToken -Path "/data/documents" -FolderName "2016"
	Creates a new folder "2016" under "/data/documents"
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$false)]
		[string]$AccessToken,
		[String]$ResourceId="",
		[Parameter(Mandatory=$True)]
		[string]$FolderName,
		[string]$Path="/",
		[string]$ElementId="",
		[string]$DriveId=""
	)
	$rURI=Format-ODPathorIdString -path $Path -ElementId $ElementId -DriveId $DriveId
	$rURI=$rURI+"/children"
	return Get-ODWebContent -AccessToken $AccessToken -ResourceId $ResourceId -Method POST -rURI $rURI -Body ('{"name": "'+$FolderName+'","folder": { },"@name.conflictBehavior": "fail"}')
}

function Remove-ODItem
{
	<#
	.DESCRIPTION
	Delete an item (folder or file).
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.PARAMETER Path
	Specifies the path of the item to be deleted.
	.PARAMETER ElementId
	Specifies the element id of the item to be deleted.
	.PARAMETER DriveId
	Specifies the OneDrive drive id. If not set, the default drive is used.
	.EXAMPLE
    Remove-ODItem -AccessToken $AuthToken -Path "/Data/documents/2016/Azure-big-picture.old.docx"
	Deletes an item
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$false)]
		[string]$AccessToken,
		[String]$ResourceId="",
		[string]$Path="",
		[string]$ElementId="",
		[string]$DriveId=""
	)
	if (($ElementId+$Path) -eq "") 
	{
		write-error("Path nor ElementId is set")
	}
	else
	{
		$rURI=Format-ODPathorIdString -path $Path -ElementId $ElementId -DriveId $DriveId
		return Get-ODWebContent -AccessToken $AccessToken -ResourceId $ResourceId -Method DELETE -rURI $rURI 
	}
}

function Get-ODItem
{
	<#
	.DESCRIPTION
	Download an item/file. Warning: A local file will be overwritten.
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.PARAMETER Path
	Specifies the path of the file to download.
	.PARAMETER ElementId
	Specifies the element id of the file to download. If Path and ElementId are given, the ElementId is used with a warning.
	.PARAMETER DriveId
	Specifies the OneDrive drive id. If not set, the default drive is used.
	.PARAMETER LocalPath
	Save file to path (if not given, the current local path is used).
	.PARAMETER LocalFileName
	Local filename. If not given, the file name of OneDrive is used.
	.EXAMPLE
    Get-ODItem -AccessToken $AuthToken -Path "/Data/documents/2016/Powershell array custom objects.docx"
	Downloads a file from OneDrive
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$false)]
		[string]$AccessToken,
		[String]$ResourceId="",
		[string]$Path="",
		[string]$ElementId="",
		[string]$DriveId="",
		[string]$LocalPath="",
		[string]$LocalFileName
	)
	if (($ElementId+$Path) -eq "") 
	{
		write-error("Path nor ElementId is set")
	}
	else
	{

		$Download=Get-ODItemProperty -AccessToken $AccessToken -ResourceId $ResourceId -Path $Path -ElementId $ElementId -DriveId $DriveId -SelectProperties "name,@content.downloadUrl,lastModifiedDateTime"
		if ($LocalPath -eq "") {$LocalPath=Get-Location}
		$LocalPath=resolve-Path ($LocalPath.TrimEnd("\")+"\")
		if ($LocalFileName -eq "")
		{
			$SaveTo=$LocalPath+$Download.name
		}
		else
		{
			$SaveTo=$LocalPath+$LocalFileName		
		}
		try
		{
			[System.Net.WebClient]::WebClient
			$client = New-Object System.Net.WebClient
			$client.DownloadFile($Download."@content.downloadUrl",$SaveTo)
			$file = Get-Item $saveTo
            $file.LastWriteTime = $Download.lastModifiedDateTime
			write-verbose("Download complete")
			return 0
		}
		catch
		{
			write-error("Download error: "+$_.Exception.Response.StatusCode+"`n"+$_.Exception.Response.StatusDescription)
			return -1
		}
	}	
}
function Add-ODItem
{
	<#
	.DESCRIPTION
	Upload an item/file. Warning: An existing file will be overwritten.
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.PARAMETER Path
	Specifies the path for the upload folder. If not given, the path is "/".
	.PARAMETER ElementId
	Specifies the element id for the upload folder. If Path and ElementId are given, the ElementId is used with a warning.
	.PARAMETER DriveId
	Specifies the OneDrive drive id. If not set, the default drive is used.
	.PARAMETER LocalFile
	Path and file of the local file to be uploaded (C:\data\data.csv).
	.EXAMPLE
    Add-ODItem -AccessToken $AuthToken -Path "/Data/documents/2016" -LocalFile "AzureML with PowerShell.docx" 
    Upload a file to OneDrive "/data/documents/2016"
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$false)]
		[string]$AccessToken,
		[String]$ResourceId="",
		[string]$Path="/",
		[string]$ElementId="",
		[string]$DriveId="",
		[Parameter(Mandatory=$True)]
		[string]$LocalFile=""
	)
	$rURI=Format-ODPathorIdString -path $Path -ElementId $ElementId -DriveId $DriveId
	$string=@()
	$string=@{ enableglobal=$true}
	if($Authentication -and -not($ResourceId)){$string=@{ClientId = $Authentication.ClientId}} 
	if($Authentication.ResourceId -and -not($ResourceId)){$string=@{ResourceId=$Authentication.ResourceId}}elseif($ResourceId){$string=@{ResourceId=$ResourceId}}
do{
$doCount++
if($doCount -ne 1){
$null=Get-ODAuthentication @string
if(-not($?)){break}}
if (-not($AccessToken)){$AccessToken=$Authentication.access_token}
	try
	{
		$spacer=""
		if ($ElementId -ne "") {$spacer=":"}
		$ODRootURI=Get-ODRootUri -ResourceId $ResourceId
		$ruri1=(($ODRootURI+$rURI).TrimEnd(":")+$spacer+"/"+[System.IO.Path]::GetFileName($LocalFile)+":/content").Replace("/root/","/root:/")
		return $webRequest=Invoke-WebRequest -Method PUT -InFile $LocalFile -Uri $rURI1 -Header @{ Authorization = "BEARER "+$AccessToken} -ContentType "multipart/form-data"  -UseBasicParsing -ErrorAction SilentlyContinue|%{$_.content}|convertfrom-json
	}
	catch
	{
		write-error("Upload error: "+$_.Exception.Response.StatusCode+"`n"+$_.Exception.Response.StatusDescription)
		#return -1
	}	
	$errorcode=$?
}until($errorcode -or $doCount -ne 1)
}
function Add-ODItemLarge {
	<#
		.DESCRIPTION
		Upload a large file with an upload session. Warning: Existing files will be overwritten.
		For reference, see: https://docs.microsoft.com/en-us/onedrive/developer/rest-api/api/driveitem_createuploadsession?view=odsp-graph-online
		.PARAMETER AccessToken
		A valid access token for bearer authorization.
		.PARAMETER ResourceId
		Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
		.PARAMETER Path
		Specifies the path for the upload folder. If not given, the path is "/".
		.PARAMETER ElementId
		Specifies the element id for the upload folder. If Path and ElementId are given, the ElementId is used with a warning.
		.PARAMETER DriveId
		Specifies the OneDrive drive id. If not set, the default drive is used.
		.PARAMETER LocalFile
		Path and file of the local file to be uploaded (C:\data\data.csv).
		.EXAMPLE
		Add-ODItem -AccessToken $AuthToken -Path "/Data/documents/2016" -LocalFile "AzureML with PowerShell.docx" 
		Upload a file to OneDrive "/data/documents/2016"
		.NOTES
		Author: Benke Tamás - (funkeninduktor@gmail.com)
	#>
	
	PARAM(
		[Parameter(Mandatory=$false)]
		[string]$AccessToken,
		[String]$ResourceId="",
		[string]$Path="/",
		[string]$ElementId="",
		[string]$DriveId="",
		[Parameter(Mandatory=$True)]
		[string]$LocalFile=""
	)

	$rURI=Format-ODPathorIdString -path $Path -ElementId $ElementId -DriveId $DriveId
	$string=@()
	$string=@{ enableglobal=$true}
	if($Authentication -and -not($ResourceId)){$string=@{ClientId = $Authentication.ClientId}} 
	if($Authentication.ResourceId -and -not($ResourceId)){$string=@{ResourceId=$Authentication.ResourceId}}elseif($ResourceId){$string=@{ResourceId=$ResourceId}}
do{
$doCount++
if($doCount -ne 1){
$null=Get-ODAuthentication @string
if(-not($?)){break}}
if (-not($AccessToken)){$AccessToken=$Authentication.access_token}
	Try	{
		# Begin to construct the real (full) URI
		$spacer=""
		if ($ElementId -ne "") {$spacer=":"}
		$ODRootURI=Get-ODRootUri -ResourceId $ResourceId
		
		# Construct the real (full) URI
		$rURI1=(($ODRootURI+$rURI).TrimEnd(":")+$spacer+"/"+[System.IO.Path]::GetFileName($LocalFile)+":/createUploadSession").Replace("/root/","/root:/")
		
		# Initialize upload session
		$webRequest=Invoke-WebRequest -Method PUT -Uri $rURI1 -Header @{ Authorization = "BEARER "+$AccessToken} -ContentType "application/json" -UseBasicParsing -ErrorAction SilentlyContinue

		# Parse the response JSON (into a holder variable)
		$convertResponse = ($webRequest.Content | ConvertFrom-Json)
		# Get the uploadUrl from the response (holder variable)
		$uURL = $convertResponse.uploadUrl
		# echo "HERE COMES THE CORRECT uploadUrl: $uURL"
		
		# Get the full size of the file to upload (bytes)
		$totalLength = (Get-Item $LocalFile).length
		# echo "Total file size (bytes): $totalLength"
		
		# Set the upload chunk size (Recommended: 5MB)
		$uploadLength = 5 * 1024 * 1024; # == 5242880 byte == 5MB.
		# echo "Size of upload fragments (bytes): $uploadLength" # == 5242880
		
		# Set the starting byte index of the upload (i. e.: the index of the first byte of the file to upload)
		$startingIndex = 0
		
		# Start an endless cycle to run until the last chunk of the file is uploaded (after that, BREAK out of the cycle)
		while($True){
			# If startingIndex (= the index of the starting byte) is greater than, or equal to totalLength (= the total length of the file), stop execution, so BREAK out of the cycle
			if( $startingIndex -ge $totalLength ){
				break
			}
			
			# Otherwise: set the suitable indices (variables)
			
			# (startingIndex remains as it was!)
			
			# Set the size of the chunk to upload
			# The remaining length of the file (to be uploaded)
			$remainingLength = $($totalLength-$startingIndex)
			# If remainingLength is smaller than the normal upload length (defined above as uploadLength), then the new uploadLength will be the remainingLength (self-evidently, only for the last upload chunk)
			if( $remainingLength -lt $uploadLength ){
				$uploadLength = $remainingLength
			}
			# Set the new starting index (just for the next iteration!)
			$newStartingIndex = $($startingIndex+$uploadLength)
			# Get the ending index (by means of newStartingIndex)
			$endingIndex = $($newStartingIndex-1)
			
			# Get the bytes to upload into a byte array (using properly re-initialized variables)
			$buf = new-object byte[] $uploadLength
			$fs = new-object IO.FileStream($LocalFile, [IO.FileMode]::Open)
			$reader = new-object IO.BinaryReader($fs)
			$reader.BaseStream.Seek($startingIndex,"Begin") | out-null
			$reader.Read($buf, 0, $uploadLength)| out-null
			$reader.Close()
			# echo "Chunk size is: $($buf.count)"
			
			# Upoad the actual file chunk (byte array) to the actual upload session.
			# Some aspects of the chunk upload:
				# We don't have to authenticate for the chunk uploads, since the uploadUrl contains the upload session's authentication data as well.
				# We above calculated the length, and starting and ending byte indices of the actual chunk, and the total size of the (entire) file. These should be set into the upload's PUT request headers.
				# If the upload session is alive, every file chunk (including the last one) should be uploaded with the same command syntax.
				# If the last chunk was uploaded, the file is automatically created (and the upload session is closed).
				# The (default) length of an upload session is about 15 minutes!
			
			# Set the headers for the actual file chunk's PUT request (by means of the above preset variables)
			$actHeaders=@{"Content-Length"="$uploadLength"; "Content-Range"="bytes $startingIndex-$endingIndex/$totalLength"};
			
			# Execute the PUT request (upload file chunk)
			write-debug("Uploading chunk of bytes. Progress: "+$endingIndex/$totalLength*100+" %")
			$uploadResponse=Invoke-WebRequest -Method PUT -Uri $uURL -Headers $actHeaders -Body $buf -UseBasicParsing -ErrorAction SilentlyContinue
			
			# startingIndex should be incremented (with the size of the actually uploaded file chunk) for the next iteration.
			# (Since the new value for startingIndex was preset above, as newStartingIndex, here we just have to overwrite startingIndex with it!)
			$startingIndex = $newStartingIndex
		}
		# The upload is done!
		
		# At the end of the upload, write out the last response, which should be a confirmation message: "HTTP/1.1 201 Created"
		write-debug("Upload complete")
		return ($uploadResponse.Content | ConvertFrom-Json)
	}
	Catch {
		write-error("Upload error: "+$_.Exception.Response.StatusCode+"`n"+$_.Exception.Response.StatusDescription)
		#return -1
	}
	$errorcode=$?
}until($errorcode -or $doCount -ne 1)
}
function Move-ODItem
{
	<#
	.DESCRIPTION
	Moves a file to a new location or renames it.
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.PARAMETER Path
	Specifies the path of the file to be moved.
	.PARAMETER ElementId
	Specifies the element id of the file to be moved. If Path and ElementId are given, the ElementId is used with a warning.
	.PARAMETER DriveId
	Specifies the OneDrive drive id. If not set, the default drive is used.
	.PARAMETER TargetPath
	Save file to the target path in the same OneDrive drive (ElementId for the target path is not supported yet).
	.PARAMETER NewName
	The new name of the file. If missing, the file will only be moved.
	.EXAMPLE
	Move-ODItem  -AccessToken $at -path "/Notes.txt" -TargetPath "/x" -NewName "_Notes.txt"
	Moves and renames a file in one step

	Move-ODItem  -AccessToken $at -path "/Notes.txt" -NewName "_Notes.txt" # Rename a file
	
	Move-ODItem  -AccessToken $at -path "/Notes.txt" -TargetPath "/x"      # Move a file
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$false)]
		[string]$AccessToken,
		[String]$ResourceId="",
		[string]$Path="",
		[string]$ElementId="",
		[string]$DriveId="",
		[string]$TargetPath="",
		[string]$NewName=""
	)
	if (($ElementId+$Path) -eq "") 
	{
		write-error("Path nor ElementId is set")
	}
	else
	{
		if (($TargetPath+$NewName) -eq "")
		{
			write-error("TargetPath nor NewName is set")
		}
		else
		{	
			$body='{'
			if (!$NewName -eq "") 
			{
				$body=$body+'"name": "'+$NewName+'"'
				If (!$TargetPath -eq "")
				{
					$body=$body+','
				}
			}
			if (!$TargetPath -eq "") 
			{
				$rTURI=Format-ODPathorIdString -path $TargetPath -DriveId $DriveId
				$body=$body+'"parentReference" : {"path": "'+$rTURI+'"}'
			}
			$body=$body+'}'
			$rURI=Format-ODPathorIdString -path $Path -ElementId $ElementId -DriveId $DriveId
			return Get-ODWebContent -AccessToken $AccessToken -ResourceId $ResourceId -Method PATCH -rURI $rURI -Body $body
		}
	}
}
function get-odsharelinkdownload
	<#
	.DESCRIPTION
	Download a shared file
	.PARAMETER URL
	onedrive Share links
	.PARAMETER path
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.EXAMPLE
    Get-ODDrives -URL https://1drv.ms/f/s!AtftJLuuzIqngqg598UpNi1x5YJ8bQ
	Download a file
	Get-ODDrives -URL xxx -path \d\d\
	.NOTES
	Avoid downloading large files
	The application for OneDrive 4 Business needs "Read items in all site collections" on application level (API: Office 365 SharePoint Online)
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
        {PAram(
        [Parameter(Mandatory=$true,Position=0)]
        [string]$uri,
        [Parameter(Mandatory=$true,Position=1)]
        [string]$path) 
        if(-not(Test-Path $path)){break;write-host 'error path'}
        if($path -match '[^/]$'){
        $path=  (resolve-path -path $path).path+"/"}
$ProgressPreference=    "SilentlyContinue"
function Runspace0{
param($ScriptBlock)
$throttleLimit = 8
$SessionState = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
$Pool = [runspacefactory]::CreateRunspacePool(1, $throttleLimit, $SessionState, $Host)
$Pool.Open()
if ($ScriptBlock -is [string]){[array]$ScriptBlock=$ScriptBlock}
$threads = @()
$handles = for ($x = 0; $x -lt $ScriptBlock.length; $x++) {
$fg=$ScriptBlock[$x]
$scriptblock1=@"
param(`$id)
`$ErrorActionPreference='SilentlyContinue';
`$WarningPreference='SilentlyContinue';
`$ProgressPreference='SilentlyContinue';
`$code=($fg)
[PScustomobject]@{id=`$id;code=`$code}
"@
    $powershell = [powershell]::Create().AddScript($scriptblock1).AddArgument($x)
    $powershell.RunspacePool = $Pool
    $powershell.BeginInvoke()
    $threads += $powershell}
if ($handles -is [string]){[array]$handles=$handles}
$ss=@()
do {
$done = $true
# ($handles -ne $null).length
for ($x=0;$X -lt $handles.length;$x++){
$bi=$handles[$x].IsCompleted -like 'true'
if ($bi){
$ss=$ss += $threads[$x].EndInvoke($handles[$x])
$threads[$x].Dispose()
$handles[$x]=$null
$threads[$x]=$null}}
if ($handles.IsCompleted -ne $null){$done = $false}
if (-not $done) { Start-Sleep -Milliseconds 900 }
} until ($done)
($ss |sort-object -Property id).code}
function Folder-downloads {
PAram([string]$URL,[string]$path='./')
[array]$URL=$URL
$path=(resolve-path -path $path).path
$data=@()
[array]$path1=$path
$yuan='https://storage.live.com/items/'
$folderID=@()
do {
if ($folderID[0] -ne $null -and $folderID[0] -ne $replace){
[string]$replace=$folderID[0]
$url=@()
[array]$folder=$path1|select-object -Skip $folder.length
$path1=@()
for ($i=0;$i -lt $folder.length;$i++) {
for ($x=0;$X -lt $folderName.length;$x++) {
$path1+=$folder[$i]+$folderName[$x]+'/'
$itempath=$folder[$i]+$folderName[$x]+"/"
$null=New-Item -path "$itempath" -itemtype Directory -force
#$folder[$i] -Name $folderName[$x]}}
$folderID|%{$url+="$yuan$_$key"}
Remove-variable folderName,folderID  -force}
for ($x=0;$x -lt $URL.length;$x++){
$xml=irm $URL[$x]
$dd= [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::GetEncoding(28591).GetBytes($xml))
$key='?'+ ($URL[$x] |select-string -pattern 'authkey.*$').matches.value
$dd=$([XML]$($dd  -replace '^[^<]+')).Folder.Items
[array]$ResourceID=$dd.Document.ResourceID
[array]$RelationshipName=$dd.Document.RelationshipName
for ($i=0;$I -lt $ResourceID.length;$i++){
$data+=[PScustomobject]@{
ResourceID=$ResourceID[$i]
RelationshipName=$RelationshipName[$i]
path= $path1[$x]+$RelationshipName[$i]
url=($yuan+$ResourceID[$i]+$key)}}
[array]$folderID=$dd.folder.ResourceID
[array]$folderName=$dd.folder.RelationshipName
#if ($folderID -eq $null){break}}
} until ($folderID -eq $null)
$script=@()
For ( $x=0; $x -lt $data.length;$x++){
$path=$data.path[$x]
$URL=$data.url[$x]
$script+="iwr '$URL' -o '$path'"}
if ($script -is [array] -and $script.length -gt 5){
runspace0 $script
}elseif($script -is [string]){iex $script}else{
for ($x=0;$x -lt $script.length;$x++){
iex $script[$x]}}
write-host "文件数量："$ResourceID.count -ForegroundColor Blue|out-host
$data}

if ($uri -notmatch 'onedrive|skydrive|storage'){
        $link=iwr  $uri -MaximumRedirection 0 -SkipHttpErrorCheck -ErrorAction Ignore
        $link=$link.headers.location 
}else{$link=$uri}
if ($link -match 'ithint\=folder'){
        $link= $link -replace '^(.*?)(?:onedrive|skydrive)(\..*)?(?:redir|download)\?resid\=(.*?\d)(\&a.*)$','$1storage$2items/$3?$4'
        [PScustomobject]@{链接地址=$link}|format-table -wrap
        Folder-downloads -URL $link -path $path
}else{
	#$link -replace '^(.*?)(?:onedrive|skydrive)(\..*)?(?:redir|download)(.*)$','$1skydrive$2download$3'|out-host
	#$link -replace '^(.*?)(?:onedrive|skydrive)(\..*)?(?:redir|download)\?resid\=(.*?\d)(\&a.*)$','$1storage$2items/$3?$4'|out-host
	$link= $link -replace '^(.*?)(?:onedrive|skydrive)(\..*)?(?:redir|download)(.*)$','$1skydrive$2download$3'
	#$link=iwr  $uri -MaximumRedirection 0 -SkipHttpErrorCheck -ErrorAction Ignore
	[PScustomobject]@{链接地址=$link}|format-table -wrap
	$data=iwr "$link"
	$replace=$data.headers.'Content-Disposition' -replace '^.*?(?=[^ ''"]+$)'
	$name=[System.Web.HttpUtility]::UrlDecode($replace)
	set-content -value $data -path "$path$name"
	[PScustomobject]@{
	name=$name
	path=$path
	URL=$link}}}