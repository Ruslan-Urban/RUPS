<#
	Copyright © Ruslan Urban
	https://github.com/Ruslan-Urban/RUPS
	
	Free for use and distribution. Keep the credits.
#>

param(
    [string] $WebApplication,
    [string] $SiteUrl
)

cls;
$Error.Clear();

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Continue;

$thisFile = Get-Item $PSCommandPath;
$thisFolderPath = $thisFile.Directory.FullName;
$reportsFolderPath = "$thisFolderPath\Reports";

md -Path $reportsFolderPath -Force -ErrorAction Ignore;

Push-Location $reportsFolderPath;

$report = @();

function formatMemberDisplayName($member){
    
    $memberDisplayName = "[$($member.Name)]";

    if($member.DisplayName){
        if($member.Email){
            $memberDisplayName = "$($member.DisplayName) <$($member.Email)>";                        
        } else {
            $memberDisplayName = "$($member.DisplayName) ($($member.UserLogin))";
        }
    }
    return $memberDisplayName;
}

try {

    if($WebApplication){
        $webApps = @() + (Get-SPWebApplication -Identity $WebApplication);
    } else {
        $webApps = @() + (Get-SPWebApplication);
    }

    if($SiteUrl){
        $sites = $webApps | Get-SPSite -Identity $SiteUrl; 
    } else {
        $sites = $webApps | Get-SPSite -Limit All;
    }

    $p1 = [ordered]@{ Total = $sites.Count; Current=0; Completed = 0.0; }

    foreach ($site in $sites | sort URL)
    {
        $p1.Current++; $p1.Completed = [Math]::Round($p1.Current * 100.0 / $p1.Total, 1);
        Write-Progress -Id 1 -Activity "$($p1.Completed.ToString('0'))%, [$($p1.Current) of $($p1.Total)] site collections" -Status $site.Url -PercentComplete $p1.Completed -CurrentOperation $site.RootWeb.Title;

        $webs = $site.AllWebs | sort Url;

        $p2 = [ordered]@{ Total = $webs.Count; Current=0; Completed = 0.0; }

        foreach ($web in $webs | sort URL)
        {
            $p2.Current++; $p2.Completed = [Math]::Round($p2.Current * 100.0 / $p2.Total, 1);
            Write-Progress -Id 2 -Activity "$($p2.Completed.ToString('0'))%, [$($p2.Current) of $($p2.Total)] sites" -Status $web.Url -PercentComplete $p2.Completed -CurrentOperation $web.Title -ParentId 1;

            $lists = $web.Lists | sort Title

            $p3 = [ordered]@{ Total = $lists.Count; Current=0; Completed = 0.0; }

            foreach ($list in $lists | sort Title)
            {

                $p3.Current++; $p3.Completed = [Math]::Round($p3.Current * 100.0 / $p3.Total, 1);
                Write-Progress -Id 3 -Activity "$($p3.Completed.ToString('0.0'))%, [$($p3.Current) of $($p3.Total)] lists" -Status $list.RootFolder.ServerRelativeUrl -PercentComplete $p3.Completed -CurrentOperation $list.Title -ParentId 2;

                if($list.HasUniqueRoleAssignments){

                    foreach ($ra in $list.RoleAssignments)
                    {
                        $reportItem = [pscustomobject][ordered]@{
                            Type = "List";
                            Site = $site.Url;
                            Web = $web.Url;
                            Url = "$($web.Url)/$($list.RootFolder.Url)";
                            Title = $list.Title;
                            Member =  formatMemberDisplayName($ra.Member);
                            Users = $users;
                            Roles = ($ra.Member.Roles | % {$_.Name}) -join ", ";
                            Message = "";
                        };

                        if($ra.Member.Users){
                            $reportItem.Users = ($ra.Member.Users | % { formatMemberDisplayName($_); } ) -join ";`n";
                        }

                        $report += $reportItem;
                    }
                }
            }

            if($web.HasUniqueRoleAssignments) {
                foreach ($ra in $web.RoleAssignments)
                {
                    $reportItem = [pscustomobject][ordered]@{
                        Type = ("Site", "SiteCollection")[$web.Url -eq $site.Url];
                        Site = $site.Url;
                        Web = $web.Url;
                        Url = $web.Url;
                        Title = $web.Title;
                        Member =  formatMemberDisplayName($ra.Member);
                        Users = "";
                        Roles = ($ra.Member.Roles | % {$_.Name}) -join "; ";
                        Message = "";
                    };

                    if($ra.Member.Users){
                        $reportItem.Users = ($ra.Member.Users | % { formatMemberDisplayName($_); } ) -join ";`n";
                    }

                    $report += $reportItem;
                }

                
            }
        }
    }

} finally {

    Write-Progress -Id 3 -Completed -Activity "Lists" -PercentComplete 100;
    Write-Progress -Id 2 -Completed -Activity "Sites" -PercentComplete 100;
    Write-Progress -Id 1 -Completed -Activity "Site collections" -PercentComplete 100;

    $report | Export-Csv -Path "$reportsFolderPath\$($thisFile.BaseName)-$($env:COMPUTERNAME)-$(Get-Date -Format 'yyyy-MMdd-HHmm').csv" -NoTypeInformation;
    Pop-Location;

}