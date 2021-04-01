<#
	Copyright © Ruslan Urban
	https://github.com/Ruslan-Urban/RUPS
	
	Free for use and distribution. Keep the credits.
#>

$siteUrl = 'https://<tenant>.sharepoint.com/';

cls;
$Error.Clear();

function WriteErrors([string] $Message, [switch]$Clear){    

    if($Message){
        Write-Host $Message -ForegroundColor Red;
    }

    $global:Error | % { 
        Write-Host $_.ToString() -ForegroundColor Red;
        Write-Host $_.ScriptStackTrace -ForegroundColor White;
        Write-Host $_.Exception.ToString() -ForegroundColor Gray;
    };

    if($Clear) {
        $global:Error.Clear();
    }
}

$thisFile = Get-Item $PSCommandPath;
$thisFolderPath = $thisFile.Directory.FullName;
$rootFolderPath = $thisFile.Directory.Parent.FullName;

$executionErrors = @();
$timestamp = Get-Date -Format "yyyy-MMdd-HHmm";

Push-Location $thisFolderPath -ErrorVariable +executionErrors;

if(!$LogsFolder) {
    $LogsFolder = "$($thisFolderPath)\Logs";
}
md $LogsFolder -ErrorAction Ignore | Out-Null;
Start-Transcript "$LogsFolder\$($Timestamp)-$($thisFile.BaseName)-ps.log" -ErrorVariable +executionErrors;


try {

    
    $executionErrors = @();

    #Import-Module PnP.PowerShell -DisableNameChecking -ErrorVariable +executionErrors;
    Import-Module SharePointPnPPowerShellOnline -DisableNameChecking -ErrorVariable +executionErrors;

    if(!$TemplatesFolder){
        $TemplatesFolder = "$($thisFolderPath)\SiteTemplates";
    }

    Set-PnPTraceLog -On -LogFile "$LogsFolder\$($Timestamp)-$($thisFile.BaseName)-pnp.log" -Level Debug -ErrorVariable +executionErrors;


    $siteConnection = Connect-PnPOnline $siteUrl -Interactive -ReturnConnection -Verbose;
    Set-PnPTraceLog -On -LogFile "$baseFolder\$($templateFile.BaseName).pnp.log" -Level Debug;

    $site = Get-PnPSite -Includes ID,URL;
    $web = Get-PnPWeb -Includes ID;

    $pagesLibrary = Get-PnPList 'SitePages' -Includes RootFolder, RootFolder.Files;
    $pageFiles = $pagesLibrary.RootFolder.Files | sort Name;

    $pagesFolder = Get-PnPFolder -Url 'SitePages' -Includes Files;
    $pageFiles = $pagesFolder.Files | sort Name;

    $webpartsFolder = "$($thisFolderPath)\SiteTemplates\Webparts";
    md $webpartsFolder -ErrorAction Ignore;

    $results = [ordered]@{
        'List' = @();
        'Quicklinks' = @();
        'Highlightedcontent' = @();
        'PageText' = @();
        'News' = @();
        'Events' = @();
    }

    foreach ($pageFile in $pageFiles)
    {

        try {
            $pageName  = [System.IO.Path]::GetFileNameWithoutExtension($pageFile.Name);        
            Write-Host " » Discovering web parts: " -NoNewline; Write-Host $pageFile.Name -ForegroundColor Yellow -NoNewline;
            $page = Get-PnPClientSidePage -Identity $pageName -Connection $siteConnection -ErrorAction Ignore ;
            if(!$page){
                Write-Host " SKIPPED: not a modern page" -ForegroundColor Gray;
                continue;
            }
            Write-Host "";
            foreach ($control in $page.Controls)
            {
                try {                    
                    Write-Host "   » $($control.Type.Name): " -NoNewline; Write-Host $control.Title -ForegroundColor Yellow -NoNewline;
                    $webpart = Get-PnPPageComponent -InstanceId $control.InstanceId -Page $page;
                    $wpType = ($webpart.Type.Name, ($webpart.Title -replace '\W',''))[[bool]$webpart.Title];

                    $resultItem = [ordered] @{
                        Title = $($page.PageTitle);
                        Page = $($pageFile.Name);
                        Url = $web.Url.Replace($web.ServerRelativeUrl, '') + $pageFile.ServerRelativeUrl;
                        Section=$($webpart.Section.Order);
                        Column = $($webpart.Column.Order);
                        Order = $($webpart.Order);
                        Class = $($webpart.Type.Name);
                        Type = $($wpType);
                    };

                    try{ $resultItem.Title = $webpart.ServerProcessedContent.GetProperty('searchablePlainTexts').GetProperty('title').GetString(); } catch { $Error.Clear(); }

                    if($webpart.PropertiesJson) {

                        $wpFileName = "$webpartsFolder\$($pageName)-$($webpart.Section.Order)-$($webpart.Column.Order)-$($webpart.Order)-$($webpart.Type.Name)-$($wpType).json";
                        $webpart.PropertiesJson > $wpFileName;

                        $wpProperties = $webpart.PropertiesJson | ConvertFrom-Json;

                        switch ($wpType.ToLowerInvariant())
                        {
                            'list' {
                                $list = Get-PnPList -Identity $wpProperties.selectedListId -Includes Views,RootFolder;
                                $view = $list.Views | ? {$_.Id -eq $wpProperties.selectedViewId} | select -First 1;

                                $resultItem.DocumentLibrary = $wpProperties.isDocumentLibrary;    
                                $resultItem.List = $list.Title;
                                $resultItem.ListUrl = $web.Url.Replace($web.ServerRelativeUrl, '') + $list.RootFolder.ServerRelativeUrl;
                                $resultItem.View = $view.Title;
                            }
                            'news' {
                                $resultItem.Layout = $wpProperties.layoutId;
                                $resultItem.NewsDataSourceProp = $wpProperties.newsDataSourceProp;
                                $resultItem.Filers = "$($wpProperties.filters | ConvertTo-Json -Compress)";
                                $resultItem.Sites  = "$($wpProperties.newsSiteList | ConvertTo-Json -Compress)";
                            }
                            'events' {
                                $list = Get-PnPList -Identity $wpProperties.selectedListId;
                                $resultItem.List = $list.Title;
                                $resultItem.ListUrl = $web.Url.Replace($web.ServerRelativeUrl, '') + $list.RootFolder.ServerRelativeUrl;
                                $resultItem.SelectedCategory = $wpProperties.selectedCategory;
                                $resultItem.DateRangeOption = $wpProperties.dateRangeOption;
                                $resultItem.StartDate = $wpProperties.startDate;
                                $resultItem.EndDate = $wpProperties.endDate;
                                $resultItem.Layout = $wpProperties.layoutId;
                            }
                            'quicklinks' {
                                $resultItem.Layout = $wpProperties.layoutId;
                                $resultItem.Items = "$($wpProperties.items | ConvertTo-Json -Compress)";
                            }
                            'highlightedcontent' {
                                $resultItem.IsTitleEnabled = $wpProperties.isTitleEnabled;
                                $resultItem.TemplateId = $wpProperties.templateId;
                                $resultItem.MaxItemsPerPage = $wpProperties.maxItemsPerPage;
                                $resultItem.LayoutId = $wpProperties.layoutId;
                                $resultItem.DataProviderId = $wpProperties.dataProviderId;
                                $resultItem.QueryMode = $wpProperties.queryMode;
                                $resultItem.Query = "$($wpProperties.query | ConvertTo-Json -Compress)";
                                $resultItem.Sites = "$($wpProperties.sites | ConvertTo-Json -Compress)";
                            }
                            Default {
                                $resultItem.Data = ($webpart | fl) | Out-String;
                            }
                        }

                    } elseif($webpart.Type.Name -like 'PageText') {
                        $resultItem.Text = $webpart.Text;
                        $wpFileName = "$webpartsFolder\$($pageName)-$($webpart.Section.Order)-$($webpart.Column.Order)-$($webpart.Order)-$($webpart.Type.Name)-$($wpType).html";
                        $webpart.Text > $wpFileName;
                    } else {
                        $resultItem.Data = ($webpart | fl) | Out-String;
                    }

                    $results[$wptype] = @() + $results[$wptype] + [pscustomobject]$resultItem;

                    Write-Host " DONE" -ForegroundColor Green;
                } catch {  WriteErrors " ERROR " -Clear;  }
            }
        } catch {  WriteErrors " ERROR " -Clear;  }
    }

    foreach ($resultSet in $results.GetEnumerator())
    {
        try {
            $resultFilePath = "$webpartsFolder\$($web.ServerRelativeUrl.TrimStart('/') -replace '/','--')-WebParts-$($resultSet.Key).csv";
            Write-Host " » Saving $($resultSet.Key): " -NoNewline; Write-Host $resultFilePath -ForegroundColor Yellow -NoNewline;        
            $resultSet.Value | Export-Csv -Path $resultFilePath -NoTypeInformation -ErrorAction Stop;
            Write-Host " DONE" -ForegroundColor Green;
        } catch {  WriteErrors " ERROR " -Clear;  }
    }

} catch { 
    WriteErrors " ERROR " -Clear; 
} finally {

    if($siteConnection.Context) {
        try { 
            #Disconnect-PnPOnline -ErrorAction Ignore; 
        } catch {  }
    }

    try { Set-PnPTraceLog -Off -ErrorAction Ignore; } catch { }

    Pop-Location;

    if ($executionErrors.Length) {
        Write-Host "Execution errors occurred" -ForegroundColor Red;
        $executionErrors | % { Write-Host $_ -ForegroundColor Red }
    }

    if ($Error) {
        WriteErrors "Errors occurred" -Clear;
    }

    Stop-Transcript;

}

