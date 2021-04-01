<#
	Copyright © Ruslan Urban
	https://github.com/Ruslan-Urban/RUPS
	
	Free for use and distribution. Keep the credits.
#>

param(
    [string] $OnPremSiteUrl = "https://sp.domain.com/MM",
    [string] $SPOSiteUrl = "https://tenant.sharepoint.com/sites/blog",
    $OnPremCredentialName = "ENV-SP2013",
    [int] $PostId = $null
)

$Error.Clear();

$thisFile = Get-Item $PSCommandPath;
$thisFolderPath = $thisFile.Directory.FullName;

#Import modules and types
Import-Module SharePointPnPPowerShellOnline -DisableNameChecking;
Import-Module CredentialManager;

# Copy DLLs from module SharePointPnPPowerShell2013 or update the path to module location
Add-Type -path "$thisFolderPath\CSOM\Microsoft.SharePoint.Client.dll";
Add-Type -path "$thisFolderPath\CSOM\Microsoft.SharePoint.Client.Runtime.dll";
Add-Type -Path "$thisFolderPath\htmlagilitypack.1.11.28\lib\Net45\HtmlAgilityPack.dll";


$spoConnection = Connect-PnPOnline -Url $spoSiteUrl -UseWebLogin -ReturnConnection;

#$onPremCredential = Get-PnPStoredCredential -Name $OnPremCredentialName -Type PSCredential;
if(!$onPremCredential) {
    $onPremCredential = Get-Credential -Message $OnPremSiteUrl;
}

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12;
$context = New-Object Microsoft.SharePoint.Client.ClientContext($OnPremSiteUrl);
$context.Credentials = $onPremCredential;
$web = $context.Web
$context.Load($web)  #just a test to ensure connection
$context.ExecuteQuery()


#pull the list of news posts to sync. only needed posts after Oct 2017
$posts = $web.Lists.GetByTitle("Posts")
$context.Load($posts);
$context.ExecuteQuery();

if (!$postId){
    #$caml = New-Object Microsoft.SharePoint.Client.CamlQuery;
    #$caml.ViewXml = "<View><Query><OrderBy><FieldRef Name=`"PublishedDate`"/></OrderBy></Query></View>";
    $caml = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery();
    $postItems = $posts.GetItems($caml)

} else {
    $postItems = $posts.GetItemById($postId);
}

$context.Load($postitems)
$context.ExecuteQuery()

$onPremSiteUri = New-Object System.Uri($OnPremSiteUrl);

foreach ($post in $postItems){

    try {

        $title = $post.FieldValues.Title;  # -replace $specialChars, "-"; #clean out the title
        $text = $post.FieldValues.Body + @"
<br />
<p>
    <span style="font-color: #bcbcbc; font-size: 0.8em;">Published by $($post.FieldValues.Editor.LookupValue) on $($post.FieldValues.PublishedDate.ToString('MMMM d, yyyy \a\t h:mm tt'))</span>
</p>
"@;

        $thumbnailUrl = "";
        if($post.FieldValues.PublishingRollupImage -match '\ssrc=\"(?<src>[^\"]+)\"\s' -and $Matches.src){
            
            $thumbnailUrl = $Matches.src;
            $thumbnailUrl = $thumbnailUrl -replace ('^'+$web.Url), $spoConnection.Url;
            $thumbnailUrl = $thumbnailUrl -replace ('^'+$OnPremSiteUri.AbsolutePath), $spoConnection.Url;
        }

        #max length for a title is 400 characters
        if ($title.Length -gt 400){
            $title = $Title.substring(0,400)
        } 

        $normalizedTitle = ($title -replace '\s','') -replace '\W','-';
        
        $existingPage = Get-PnPClientSidePage -Identity $normalizedTitle -Connection $spoConnection -ErrorAction Ignore;
        if($existingPage) {
            Write-Host "Updating post: " -NoNewline; Write-Host $title -ForegroundColor Yellow -NoNewline;
            Get-PnPClientSideComponent -Page $existingPage -Connection $spoConnection | % {
                Remove-PnPClientSideComponent -Page $existingPage -InstanceId $_.InstanceId -Force -Connection $spoConnection | Out-Null;
            }

            $pageText = Add-PnPClientSideText -Page $existingPage -Section 1 -Text $text -Column 1  -Connection $spoConnection;
            $existingPage = Set-PnPClientSidePage -Identity $newPage -Title $title -ThumbnailUrl $thumbnailUrl -PromoteAs NewsArticle -Publish -Connection $spoConnection;
            Write-Host " OK " -ForegroundColor Green;
            
        } else {
            Write-Host "Creating post: " -NoNewline; Write-Host $title -ForegroundColor Yellow -NoNewline;
            $newPage = Add-PnPClientSidePage -Name $normalizedTitle -LayoutType Article -PromoteAs NewsArticle -Connection $spoConnection -ErrorAction Stop;
            $pageSection = Add-PnPClientSidePageSection -Page $newPage -SectionTemplate OneColumn  -Connection $spoConnection;
            $pageText = Add-PnPClientSideText -Page $newPage -Section 1 -Text $text -Column 1  -Connection $spoConnection;
            $newPage = Set-PnPClientSidePage -Identity $newPage -Title $title -ThumbnailUrl $thumbnailUrl -PromoteAs NewsArticle -Publish -Connection $spoConnection;
            Write-Host " OK " -ForegroundColor Green;
        }
    } catch {
        Write-Host " ERROR " -ForegroundColor Red;
        $Error | % { 
            Write-Host $_.Exception.Message -ForegroundColor Red;
            Write-Host $_.Exception.StackTrace -ForegroundColor Gray;
        }
        $Error.Clear();
    }

}

Write-Host "";
Write-Host " DONE " -ForegroundColor Green;