<#
	Copyright © Ruslan Urban
	https://github.com/Ruslan-Urban/RUPS
	
	Free for use and distribution. Keep the credits.
#>

$wpInventory = Import-Csv -Path 'WebParts.csv'; # Rencore Migration tool scan output saved as CSV
$credentialName = 'ENV-SP2013'; # Generic Windows store credential name

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

cls;
$Error.Clear();

$thisFile = Get-Item $PSCommandPath;
$thisFolderPath = $thisFile.Directory.FullName;
$rootFolderPath = $thisFile.Directory.Parent.FullName;

$executionErrors = @();
$timestamp = Get-Date -Format "yyyy-MMdd-HHmm";

Push-Location $thisFolderPath -ErrorVariable +executionErrors;

$reportsFolder = "$($thisFolderPath)\Reports";
md $reportsFolder -ErrorAction Ignore | Out-Null;

$webPartsFolder = "$($thisFolderPath)\WebParts";
md $webPartsFolder -ErrorAction Ignore | Out-Null;

$reportFilePath = "$reportsFolder\$($Timestamp)-$($thisFile.BaseName).csv";
'"Site","Page","Title","ContentLink"' > $reportFilePath;

$LogsFolder = "$($thisFolderPath)\Logs";
md $LogsFolder -ErrorAction Ignore | Out-Null;

Start-Transcript "$LogsFolder\$($Timestamp)-$($thisFile.BaseName)-ps.log" -ErrorVariable +executionErrors; 

$result = @();

try {


    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12;
    Import-Module SharePointPnPPowerShell2013 -DisableNameChecking -ErrorVariable +executionErrors;

    Set-PnPTraceLog -On -LogFile "$LogsFolder\$($Timestamp)-$($thisFile.BaseName)-pnp.log" -Level Debug -ErrorVariable +executionErrors;
    $wpPages = $wpInventory | group SpcFileUrl;

    $credential = Get-PnPStoredCredential -Name $credentialName -Type OnPrem;
    if(!$credential){
        $credential = Get-Credential -Message $credentialName;
        Add-PnPStoredCredential -Name $credentialName -Username $credential.UserName -Password $credential.Password;
    }

    $p1 = [ordered]@{ Total = $wpPages.Count; Current=0; Completed = 0.0; }
 
    foreach ($wpPage in $wpPages)
    {

        try
        {

            $p1.Current++; $p1.Completed = [Math]::Round($p1.Current * 100.0 / $p1.Total, 1);
            Write-Progress -Id 1 -Activity "Discovering Content Editor webparts with content links" -Status "$($p1.Completed.ToString('0'))%, [$($p1.Current) of $($p1.Total)]" -PercentComplete $p1.Completed -CurrentOperation "[$($wpPage.Name)";            

            Write-Host "Page: " -NoNewline; Write-Host " $($wpPage.Name)" -ForegroundColor Yellow;

            $pageUrl = "$($wpPage.Name)".ToLowerInvariant();
            $siteUrl = "$($wpPage.Group[0].SpcSiteUrl)".ToLowerInvariant();
            $webUrl = $pageUrl;
            $relUrlParts = $pageUrl.Split('/');
            $webConnection = $null;
            
            while(!$webConnection.Context -and $relUrlParts.Length -gt 0 -and $webUrl.Length -ge $siteUrl.Length){
                try {
                    $relUrlParts = $relUrlParts | select -SkipLast 1;
                    $webUrl = $($relUrlParts -join '/');
                    Write-Host " » Connecting to: " -NoNewline; Write-Host " $($webUrl)" -ForegroundColor Cyan -NoNewline;
                    $webConnection = Connect-PnPOnline $webUrl -Credentials $credentialName -ReturnConnection -ErrorAction Stop;
                    
                } catch { 
                    Write-Host " ERROR" -ForegroundColor Red;
                    $Error.Clear(); 
                }
            }
            if($webConnection.Context){
                Write-Host " DONE" -ForegroundColor Green;
            } else {
                Write-Host "ERROR: Cannot connect" -ForegroundColor Red;
                continue;
            }

            $site = Get-PnPSite;
            $web = Get-PnPWeb;

            $pageServerRelativeUrl = $web.ServerRelativeUrl + $pageUrl.Replace("$($webUrl)",'');
            
            $webParts = Get-PnPWebPart -ServerRelativePageUrl $pageServerRelativeUrl -ErrorAction Stop;
           

            foreach ($webPart in $webParts)
            {
                try {
                    Write-Host "   » Webpart: " -NoNewline; 
                    Write-Host " $($webpart.WebPart.Title)" -ForegroundColor Yellow -NoNewline;

                    $webpartXml = Get-PnPWebPartXml -ServerRelativePageUrl $pageServerRelativeUrl -Identity $webPart.Id -ErrorAction Stop;
                    $webPartFilePath = "$webPartsFolder\$($pageServerRelativeUrl -replace '[\/\.\s]','-')-$($webpart.WebPart.Title -replace '\W','').xml";
                    $webpartXml > $webPartFilePath;

                    if($webPart.WebPart.Properties.FieldValues.ContentLink)   {
                        Write-Host " $($webPart.WebPart.Properties.FieldValues.ContentLink)" -ForegroundColor Magenta -NoNewline;
                        Write-Host " FOUND" -ForegroundColor Green;

                        $wpInfo =  [pscustomobject][ordered]@{
                            Site = $webUrl;
                            Page = $wpPage.Name;
                            Title = $webpart.WebPart.Title;
                            ContentLink = $webPart.WebPart.Properties.FieldValues.ContentLink;
                        };

                        $wpCSV = ($wpInfo | ConvertTo-Csv -NoTypeInformation).Split("`r`n") | select -last 1;
                        $wpCSV >> $reportFilePath;
                        $result += $wpInfo;
                    } else {
                        Write-Host " SKIPPED" -ForegroundColor Gray;
                    }
                } catch { WriteErrors " ERROR " -Clear; }
            }
            
        } catch { WriteErrors " ERROR " -Clear; }

        Write-Host '';

        if($webConnection.Context) {
            try { Disconnect-PnPOnline -Connection $webConnection -ErrorAction Ignore; } catch {  }
        }

    }

    
} catch {

    WriteErrors " ERROR " -Clear; 

} finally {

    Write-Progress -Id 1 -Completed -Activity "$($thisFile.BaseName -replace '\W', ' ')" -PercentComplete 100;

    $result | Export-Csv -Path "$reportsFolder\$($Timestamp)-$($thisFile.BaseName).copy.csv" -NoTypeInformation -Delimiter "`t";

    if($webConnection.Context) {
        #try { Disconnect-PnPOnline -Connection $webConnection -ErrorAction Ignore; } catch {  }
    }

    if($adminConnection.Context) {
        #try { Disconnect-PnPOnline -Connection $adminConnection -ErrorAction Ignore; } catch {  }
    }

    try { Set-PnPTraceLog -Off -ErrorAction Ignore; } catch { }
    #try { Disconnect-SPOService -ErrorAction Ignore; } catch { }

    Pop-Location;

    if ($executionErrors.Length) {
        Write-Host "Execution errors occurred" -ForegroundColor Red;
        $executionErrors | % { Write-Host $_ -ForegroundColor Red }
    }

    

    if ($Error) {
        Write-Host "Errors occurred" -ForegroundColor Red;
        WriteErrors -Clear;
    }

    Stop-Transcript;

}