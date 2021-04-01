<#
	Copyright © Ruslan Urban
	https://github.com/Ruslan-Urban/RUPS
	
	Free for use and distribution. Keep the credits.
#>

$site = Get-SPSite 'https://sp.domain.com'
#$site.AllWebs | select Title, Url | Out-GridView

$wfas = @();

foreach ($web in $site.AllWebs)
{
   
    $web.WorkflowAssociations | ? {$_.Name -NotMatch "\(Previous version\:" } | % { 
        $wfas += [pscustomobject][ordered]@{
            Type="Site";
            Nintex = $_.InstantiationUrl -Like  "*/NintexWorkflow/*";
            WebUrl = $web.Url ;
            ListUrl = "";
            URL = $_.InstantiationUrl;
            Web = $web.Title;
            List = "";
            Workflow = $_.Name;
            WFA = $_;
        }
    };
    foreach ($list in $web.Lists)
    {
        $list.WorkflowAssociations | ? {$_.Name -NotMatch "\(Previous version\:" } | % { 
            $wfas += [pscustomobject][ordered]@{
                Type="Site";
                Nintex = $_.InstantiationUrl -Like  "*/NintexWorkflow/*";
                WebUrl = $web.Url ; 
                ListUrl = $web.Site.Url + $list.RootFolder.ServerRelativeUrl;
                URL = $_.InstantiationUrl;
                Web = $web.Title;
                List = $list.Title;
                Workflow = $_.Name;
                WFA = $_;
            }
        };
    
    }
}

$wfas | sort Nintex, WebUrl, ListUrl | select Nintex, Type, Web, List, Workflow | ft -AutoSize


