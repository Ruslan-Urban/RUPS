<#
	Copyright © Ruslan Urban
	https://github.com/Ruslan-Urban/RUPS
	
	Free for use and distribution. Keep the credits.
#>

asnp *sharepoint*;

$sitesFileName = "Sites.csv";
$websFileName = "Webs.csv";
$listsFileName = "Lists.csv";
$ltFileName = "ListTemplates.csv";

$thisFile = Get-Item $PSCommandPath;
cd $thisFile.Directory.FullName;
$sites = Get-SPWebApplication | Get-SPSite -Limit ALL | sort ServerRelativeUrl ;

"Title`tURL`tSub-sites" > $sitesFileName;
"SiteURL`tTitle`tURL`tSubsites`tLists`tCreated`tWebTemplate`tWebTemplateId" > $websFileName;
"SiteURL`tWebURL`tListUrl`tListTitle`tBaseType`tBaseTemplate`tHidden`tItems`tLastItemModifiedDate`tLastItemDeletedDate`tHasUniqueRoleAssignments`tRoleAssignments`tUseFormsForDisplay`tCanReceiveEmail`tCheckedOutFiles`tContentTypes`tFields`tViews`tEventReceivers" > $listsFileName;

$p1 = [ordered]@{Total = $sites.Count; Current=0; Completed = 0.0;}

foreach ($site in $sites)
{
    $p1.Current++; $p1.Completed = [Math]::Round($p1.Current * 100.0 / $p1.Total, 1);
    Write-Progress -Id 1 -Activity "$($p1.Completed.ToString('0'))%, [$($p1.Current) of $($p1.Total)] site collections" -Status $site.Url -PercentComplete $p1.Completed -CurrentOperation $site.RootWeb.Title;

    "$($site.RootWeb.Title) ($($site.ServerRelativeUrl.Split('/') | select -last 1))`t$($site.URL)`t$($site.AllWebs.Count - 1)" >> $sitesFileName;    

    $p2 = [ordered]@{Total = $site.AllWebs.Count; Current=0; Completed = 0.0;}

    foreach ($web in $site.AllWebs)
    {
        $p2.Current++; $p2.Completed = [Math]::Round($p2.Current * 100.0 / $p2.Total, 1);
        Write-Progress -Id 2 -Activity "$($p2.Completed.ToString('0'))%, [$($p2.Current) of $($p2.Total)] sites" -Status $web.Url -PercentComplete $p2.Completed -CurrentOperation $web.Title -ParentId 1;

        "$($site.URL)`t$($web.Title)`t$($web.URL)`t$($web.Webs.Count)`t$($web.Lists.Count)`t$($web.Created.ToString('yyyy-MM-dd'))`t$($web.WebTemplate)`t$($web.WebTemplateId)" >> $websFileName;

        $p3 = [ordered]@{Total = $web.Lists.Count; Current=0; Completed = 0.0;}

        foreach ($list in $web.Lists)
        {
         try {   
            $p3.Current++; $p3.Completed = [Math]::Round($p3.Current * 100.0 / $p3.Total, 1);
            Write-Progress -Id 3 -Activity "$($p3.Completed.ToString('0.0'))%, [$($p3.Current) of $($p3.Total)] lists" -Status $list.RootFolder.ServerRelativeUrl -PercentComplete $p3.Completed -CurrentOperation $list.Title -ParentId 2;

            "$($site.URL)`t$($web.URL)`t$($site.URL)$($list.RootFolder.ServerRelativeUrl)`t""$($list.Title)""`t$($list.BaseType)`t$($list.BaseTemplate)" `
            +"`t$($list.Hidden)`t$($list.Items.Count)`t$($list.LastItemModifiedDate.ToString('yyyy-MM-dd'))`t$($list.LastItemDeletedDate.ToString('yyyy-MM-dd'))" `
            +"`t$($list.HasUniqueRoleAssignments)`t$($list.RoleAssignments.Count)`t$($list.UseFormsForDisplay)`t$($list.CanReceiveEmail)`t$($list.CheckedOutFiles.Count)" `
            +"`t""$($list.ContentTypes.Name -join '; ')""`t""$(($list.Fields | ? {!$_.FromBaseType} | % { $_. Title}) -join '; ')""`t""$($list.Views.Title -join '; ')""" `
            +"`t$(($list.EventReceivers | % {$_.Name} | sort | select -Unique) -join '; ')" >> $listsFileName;
            } catch {
                $Error | % {Write-Host $_.Exception.Message -ForegroundColor Red};
                $Error.Clear();
            }
        }
    }
}

Get-Content $sitesFileName;
Get-Content $websFileName;
