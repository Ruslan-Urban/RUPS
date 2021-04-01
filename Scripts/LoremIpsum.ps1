<#
	Copyright © Ruslan Urban
	https://github.com/Ruslan-Urban/RUPS
	
	Free for use and distribution. Keep the credits.
#>

function New-LoremIpsum
{
    param(
		[int]$MinWords = 2, 
		[int]$MaxWords = 5, 
		[int]$MinSentences = 2, 
		[int]$MaxSentences = 5, 
		[int]$MinParagraphs = 2,
		[int]$MaxParagraphs = 5,
        [string] $SentenceEnd = ".",
        [string] $ParagrapDelimiter = "`r`n",
        [switch] $Capitalize
	) 
    
    if ($MinWords -le 0 -or $MaxWords -le 0 -or $MinSentences -le 0 -or $MaxSentences -le 0 -or $MinParagraphs -le 0 -or $MaxParagraphs -le 0)
    {
        throw "Min/Max parameters must be greater than 0.";
    }
    
    if ($MinWords -gt $MaxWords)
    {
        throw "MinWords cannot be greater than MaxWords.";
    }
    
    if ($MinSentences -gt $MaxSentences)
    {
        throw "MinSentences cannot be greater than MaxSentences.";
    }

    if ($MinParagraphs -gt $MaxParagraphs)
    {
        throw "MinParagraphs cannot be greater than MaxParagraphs.";
    }

	$lipsum = @(
		"a","ac","accumsan","ad","adipiscing","aenean","aliquam","aliquet","amet","ante","aptent","arcu",
		"at","auctor","augue","bibendum","blandit","class","commodo","condimentum","congue","consectetur",
		"consequat","conubia","convallis","cras","cubilia","curabitur","curae","cursus","dapibus","diam",
		"dictum","dictumst","dignissim","dis","dolor","donec","dui","duis","efficitur","egestas","eget",
		"eleifend","elementum","elit","enim","erat","eros","est","et","etiam","eu","euismod","ex","facilisi",
		"facilisis","fames","faucibus","felis","fermentum","feugiat","finibus","fringilla","fusce","gravida",
		"habitant","habitasse","hac","hendrerit","himenaeos","iaculis","id","imperdiet","in","inceptos",
		"integer","interdum","ipsum","justo","lacinia","lacus","laoreet","lectus","leo","libero","ligula",
		"litora","lobortis","lorem","luctus","maecenas","magna","magnis","malesuada","massa","mattis","mauris",
		"maximus","metus","mi","molestie","mollis","montes","morbi","mus","nam","nascetur","natoque","nec","neque",
		"netus","nibh","nisi","nisl","non","nostra","nulla","nullam","nunc","odio","orci","ornare","parturient",
		"pellentesque","penatibus","per","pharetra","phasellus","placerat","platea","porta","porttitor","posuere",
		"potenti","praesent","pretium","primis","proin","pulvinar","purus","quam","quis","quisque","rhoncus","ridiculus",
		"risus","rutrum","sagittis","sapien","scelerisque","sed","sem","semper","senectus","sit","sociosqu","sodales",
		"sollicitudin","suscipit","suspendisse","taciti","tellus","tempor","tempus","tincidunt","torquent","tortor",
		"tristique","turpis","ullamcorper","ultrices","ultricies","urna","ut","varius","vehicula","vel","velit",
		"venenatis","vestibulum","vitae","vivamus","viverra","volutpat","vulputate"
	);
    
    $numParagraphs = Get-Random -Minimum $MinParagraphs -Maximum ($MaxParagraphs + 1);
    $numSentences = Get-Random -Minimum $MinSentences -Maximum ($MaxSentences + 1);
    $numWords = Get-Random -Minimum $MinWords -Maximum ($MaxWords + 1);

    $result = @();
	$paragraphs = @();
	
    for($p = 0; $p -lt $numParagraphs; $p++) {
		$sentences = @();
        for($s = 0; $s -lt $numSentences; $s++) {
			$words = @();
            for($w = 0; $w -lt $numWords; $w++) {
				$word = $lipsum[(Get-Random -Minimum 0 -Maximum $lipsum.Length)];
				if($w -eq 0 -or ($Capitalize -and $word.Length -gt 2)) { $word = $word.Substring(0,1).ToUpperInvariant() + $word.Substring(1, $word.Length-1)}
				$words += $word;
            }
            $sentence = ($words -join " ") + $SentenceEnd;
			$sentences += $sentence;
        }
        $paragraph = $sentences -join " ";
		$paragraphs += $paragraph;
    }

    $result = $paragraphs -join $ParagraphDelimiter;

    return $result;
}
