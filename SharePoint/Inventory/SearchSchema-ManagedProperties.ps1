<#
	Copyright © Ruslan Urban
	https://github.com/Ruslan-Urban/RUPS
	
	Free for use and distribution. Keep the credits.
#>

cls;
$Error.Clear();

$thisFile = Get-Item $PSCommandPath;
$thisFolderPath = $thisFile.Directory.FullName;
$reportsFolderPath = "$thisFolderPath\Reports";

Add-PSSnapin *sharepoint*;

$ssa = Get-SPEnterpriseSearchServiceApplication;
$mps = Get-SPEnterpriseSearchMetadataManagedProperty -SearchApplication $ssa;
$customMPs = $mps | ? {!$_.SystemDefined}

md -Path $reportsFolderPath -Force -ErrorAction Ignore;

$reportMPProperties = "Name","ManagedType","Searchable","Queryable","Retrievable","Refinable","Sortable";
#"ID","PID","SystemDefined","Name","SplitStringCharacters","SplitStringOnSpace","Description",
#"ManagedType","Searchable","FullTextQueriable","Queryable","Retrievable","Refinable","Sortable",
#"FullTextIndex","Context","RefinerConfiguration","SortableType","UseAAMMapping","MappingDisallowed","DeleteDisallowed",
#"IsReadOnly","QueryIndependentRankCustomizationDisallowed","EnabledForScoping","NameNormalized","RespectPriority",
#"RemoveDuplicates","HasMultipleValues","OverrideValueOfHasMultipleValues","IsInDocProps","IncludeInMd5",
#"IncludeInAlertSignature","SafeForAnonymous","NoWordBreaker","EntityExtractorBitMap",
#"IndexOptions","IndexOptionsNoPrefix","IndexOptionsNoPositions","UserFlags","Weight","LengthNormalization",
#"EnabledForQueryIndependentRank","DefaultForQueryIndependentRank","IsInFixedColumnOptimizedResults","RetrievableForResultsOnly",
#"PutInPropertyBlob","QueryPropertyBlob","UsePronunciationString","MaxCharactersInPropertyStoreIndex","MaxCharactersInPropertyStoreNonIndex",
#"EqualityMatchOnly","MaxCharactersInPropertyStoreForRetrieval","DecimalPlaces","UpdateGroup","ExtraProperties","NicknameExpansion",
#"TokenNormalization","MaxNumberOfValues","ResultFallback","URLNormalization","AliasesOverridden","MappingsOverridden","IsBacked"

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

$p1 = [ordered]@{ Total = $customMPs.Count; Current=0; Completed = 0.0; }

foreach ($customMP in $customMPs | sort Name)
{
    $p1.Current++; $p1.Completed = [Math]::Round($p1.Current * 100.0 / $p1.Total, 1);
    Write-Progress -Id 1 -Activity "Managed Search Schema" -Status "$($p1.Completed.ToString('0'))%, [$($p1.Current) of $($p1.Total)]" -PercentComplete $p1.Completed -CurrentOperation "$($customMP.Name)";

    Write-Host " » Managed search property: " -NoNewline; Write-Host $customMP.Name -ForegroundColor Yellow -NoNewline;

    try {

        $resultItem = [ordered]@{};
        foreach ($propertyName in $reportMPProperties)
        {
            $resultItem.$propertyName = $customMP.$propertyName;
        }
        $mmp = Get-SPEnterpriseSearchMetadataMapping -SearchApplication $ssa -ManagedProperty $customMP;

        $resultItem.MappedProperties = $mmp.CrawledPropertyName -join ';';

        $result += [pscustomobject]$resultItem;
        Write-Host " DONE" -ForegroundColor Green;

    } catch { WriteErrors " ERROR " -Clear; }

}

$result | Export-Csv -Path "$reportsFolderPath\SearchSchema-ManagedProperties.csv" -NoTypeInformation -Force;