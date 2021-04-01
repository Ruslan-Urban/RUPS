<#
	Copyright © Ruslan Urban
	https://github.com/Ruslan-Urban/RUPS
	
	Free for use and distribution. Keep the credits.
#>

param(

    [string] $SiteUrl = "https://<tenant>.sharepoint.com/",
    [string[]] $ListNames = @("Pages", "Site Pages")
)


cls;
$Error.Clear();

$thisFile = Get-Item $PSCommandPath;
$thisFolderPath = $thisFile.Directory.FullName;
$rootFolderPath = $thisFile.Directory.Parent.FullName;

$executionErrors = @();
$timestamp = Get-Date -Format "yyyy-MMdd-HHmm";

Push-Location $thisFolderPath -ErrorVariable +executionErrors;

$LogsFolder = "$($thisFolderPath)\Logs";
md $LogsFolder -ErrorAction Ignore | Out-Null;
Start-Transcript "$LogsFolder\$($Timestamp)-$($thisFile.BaseName)-ps.log" -ErrorVariable +executionErrors;

try {

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12;
    Import-Module SharePointPnPPowerShellOnline -DisableNameChecking -ErrorVariable +executionErrors;

    Set-PnPTraceLog -On -LogFile "$LogsFolder\$($Timestamp)-$($thisFile.BaseName)-pnp.log" -Level Debug -ErrorVariable +executionErrors;

    Connect-PnPOnline -Url $SiteUrl -UseWebLogin;

    Write-Host "Ensure the modern page feature is enabled..." -NoNewline;
    Enable-PnPFeature -Identity "B6917CB1-93A0-4B97-A84D-7CF49975D4EC" -Scope Web -Force -ErrorAction Stop;
    Write-Host " DONE " -ForegroundColor Green;

    foreach ($listName in $ListNames)
    {

        
        $list = Get-PnPList -Identity $listName -ErrorAction Ignore;

        if(!$list){
            Write-Warning "List $($listName) not found" ;
            continue;
        }

        Write-Host "Processing list " -NoNewline; Write-Host "$($list.Title)" -ForegroundColor Yellow;
        $rootFolder = $list.RootFolder;
        $list.Context.Load($rootFolder);
        Invoke-PnPQuery;

        $pages = Get-PnPListItem -List $list;

        foreach($page in $pages)
        {
            try {
                $pageName = $page.FieldValues["FileLeafRef"];
                if($pageName -like 'Previous_*'){
                    continue;
                }

                Write-Host " »  Converting page to modern " -NoNewline; Write-Host "$($pageName)" -ForegroundColor Cyan -NoNewline;
        
                if ($page.FieldValues["ClientSideApplicationId"] -eq "b6917cb1-93a0-4b97-a84d-7cf49975d4ec" ) 
                { 
                    Write-Host " MODERN " -ForegroundColor DarkYellow;
                } 
                else 
                {
                    
            
                    # -TakeSourcePageName:
                    # The old pages will be renamed to Previous_<pagename>.aspx. If you want to 
                    # keep the old page names as is then set the TakeSourcePageName to $false. 
                    # You then will see the new modern page be named Migrated_<pagename>.aspx

                    # -Overwrite:
                    # Overwrites the target page (needed if you run the modernization multiple times)
            
                    # -LogVerbose:
                    # Add this switch to enable verbose logging if you want more details logged

                    # KeepPageCreationModificationInformation:
                    # Give the newly created page the same page author/editor/created/modified information 
                    # as the original page. Remove this switch if you don't like that

                    # -CopyPageMetadata:
                    # Copies metadata of the original page to the created modern page. Remove this
                    # switch if you don't want to copy the page metadata

                    $convertedPageRelativeUrl = ConvertTo-PnPClientSidePage -Identity $page.FieldValues["ID"] `
                        -Folder $rootFolder `
                        -Overwrite `
                        -TakeSourcePageName `
                        -LogType File `
                        -LogFolder $LogsFolder `
                        -LogSkipFlush `
                        -LogVerbose `
                        -KeepPageCreationModificationInformation `
                        -CopyPageMetadata `
                        -ErrorAction Stop;

                    if($convertedPageFullUrl) {
                        Write-Host " DONE " -ForegroundColor Green;
                        $convertedPageFullUrl = "$($web.Url.Replace($web.ServerRelativeUrl, ''))/$convertedPageRelativeUrl";
                        start $convertedPageFullUrl;
                    } else {
                        Write-Host " ERROR " -ForegroundColor Red;
                    }
                }
    
                

            } catch {
                Write-Host " ERROR " -ForegroundColor Red;

                $Error | % { Write-Host $_.ToString() -ForegroundColor Red; }
                $Error.Clear();
    
            }

        }
    }

    # Write the logs to the folder
    Write-Host "Writing the conversion log file..." -ForegroundColor Green
    Save-PnPClientSidePageConversionLog;

    Write-Host "Wiki and web part page modernization complete! :)" -ForegroundColor Green

} finally {

    if($siteConnection.Context) {
        try { Disconnect-PnPOnline -Connection $siteConnection -ErrorAction Ignore; } catch {  }
    }

    if($adminConnection.Context) {
        try { Disconnect-PnPOnline -Connection $adminConnection -ErrorAction Ignore; } catch {  }
    }

    try { Set-PnPTraceLog -Off -ErrorAction Ignore; } catch { }
    try { Disconnect-SPOService -ErrorAction Ignore; } catch { }

    Pop-Location;

    if ($executionErrors.Length) {
        Write-Host "Execution errors occurred" -ForegroundColor Red;
        $executionErrors | % { Write-Host $_ -ForegroundColor Red }
    }

    if ($Error) {
        Write-Host "Errors occurred" -ForegroundColor Red;
        $Error | % { 
            Write-Host $_.ToString() -ForegroundColor Red;
            Write-Host $_.Exception -ForegroundColor Gray;
        };
        $Error.Clear();
    }

    Stop-Transcript;

}
