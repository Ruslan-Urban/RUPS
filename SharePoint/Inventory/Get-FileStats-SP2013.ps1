<#
	Copyright © Ruslan Urban
	https://github.com/Ruslan-Urban/RUPS
	
	Free for use and distribution. Keep the credits.
#>

cls;
$Error.Clear();

if(!(Get-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Ignore)) {
    Add-PSSnapin Microsoft.SharePoint.PowerShell;
}

$timestamp = Get-Date -Format "yyyy-MMdd-HHmm";

$thisFile = Get-Item $PSCommandPath;
$thisFolderPath = $thisFile.Directory.FullName;
$parentFolderPath = $thisFile.Directory.FullName;
Push-Location $thisFolderPath;

$logsFolderPath = "$thisFolderPath\Logs";
$outputFolderPath = "$thisFolderPath\Reports";
$reportFilePath = "$outputFolderPath\$($thisFile.BaseName)-$($env:COMPUTERNAME)-$($timestamp).csv";

md -Path $logsFolderPath -ErrorAction Ignore | Out-Null;
md -Path $outputFolderPath -ErrorAction Ignore | Out-Null;

Start-Transcript "$logsFolderPath\$($timestamp)-$($thisFile.BaseName)-ps.log" -IncludeInvocationHeader;


function GetFilesReport($WebApp, $Site, $Web, $List, $Folder){

    $report = [ordered]@{
        WebAppUrl = $WebApp.Url;
        SiteUrl = $Site.Url;
        WebUrl = $Web.Url;
        Web = $Web.Title;
        List = $List.Title;
        Folders = 1;
        Files = 0;
        FilesVersions = 0;
        DraftVersions = 0;
        CheckoutCount = 0;
        LockedCount = 0;
        FilesSize = 0.0;
        FilesVersionsSize = 0.0;
        LevelCheckout = 0;
        LevelDraft = 0;
        LevelPublished = 0;
        ApprovalDenied = 0;
        ApprovalDraft = 0;
        ApprovalPending = 0;
        ApprovalApproved = 0;
    }

    foreach($file in $folder.Files){
        $report.Files++;
        $report.FilesSize += $file.TotalLength;
        if($file.MinorVersion -ne 0 ){
            $report.DraftVersions++;
        }
        if($file.LockType -ne 'None' ){
            $report.LockedCount++;
        }
        if($file.CheckOutStatus -ne 'None' ){
            $report.CheckoutCount++;
        }
        if($file.ListItemAllFields.ModerationInformation.Status){
            $approvalStatus = "Approval$($file.ListItemAllFields.ModerationInformation.Status)";
            if(!($report.Contains($approvalStatus))){
                $report.$approvalStatus = 1;
            } else {
                $report.$approvalStatus = $report.$approvalStatus + 1;
            }
        }
        if($file.Versions){
            foreach ($fileVersion in $file.Versions)
            {
                $report.FilesVersions ++;
                $report.FilesVersionsSize += [int]$fileVersion.Size;
            }
        }
        $level = "Level$($file.Level)";
        if(!($report.Contains($level))){
            $report.$level = 1;
        } else {
            $report.$level = $report.$level + 1;
        }
    }

    foreach($subfolder in $folder.Folders){
        try {
            $subfolderReport = GetFilesReport -WebApp $WebApp -Site $Site -Web $web -List $List -Folder $subfolder;
            $report.Folders += $subfolderReport.Folders; 
            $report.Files += $subfolderReport.Files; 
            $report.FilesSize += $subfolderReport.FilesSize; 
            $report.FilesVersions += $subfolderReport.FilesVersions; 
            $report.FilesVersionsSize += $subfolderReport.FilesVersionsSize; 
        } catch {}
    }
    
    return [pscustomobject]$report;
}

try {

    $listReports = @();
    $reportedCount = 0;

    #$webApps = Get-SPWebApplication;
    $webApps = Get-SPWebApplication -Identity 'https://svvapp204.agutl.com/';

    foreach ($webApp in $webApps | sort Url)
    {
        try {

            $sites =  $webApp | Get-SPSite -Limit All;

            foreach ($site in $sites | sort Url)
            {
                try {

                    foreach ($web in $site.AllWebs | sort Url)
                    {
                        try {
                            foreach($list in $web.Lists)
                            {
                                try {
                                    $listReport = GetFilesReport -WebApp $webApp -Site $site -Web $web -List $list -Folder $list.RootFolder;
                                    $listReports += $listReport;
                                    Write-Host "$(($listReport.Files + $listReport.FilesVersions).ToString('###,##0').PadLeft(6, ' '))`t$(($listReport.FilesSize + $listReport.FilesVersionsSize).ToString('###,###,###,##0').PadLeft(15, ' '))`t$($listReport.WebUrl)`t|`t$($listReport.List)"
                                } catch {}
                            }
                            if($listReports.Count - $reportedCount -gt 1000) {
                                Remove-Item -Path $reportFilePath -Force -Confirm:$false -ErrorAction Ignore;
                                $listReports | Export-CSV $reportFilePath -NoTypeInformation -Force;
                                $reportedCount = $listReports.Count;
                            }
                        } catch {}
                    }
                } catch {}
            }

        } catch {}
    }


} finally {
    Pop-Location;
}

Remove-Item -Path $reportFilePath -Force -Confirm:$false -ErrorAction Ignore;
$listReports | Export-CSV $reportFilePath -NoTypeInformation -Force;

if($Error.Count) {
    $Error | % { Write-Host $_.Exception -ForegroundColor Red };
    $Error.Clear();
}

Stop-Transcript;

Write-Host "DONE" -ForegroundColor Green;
