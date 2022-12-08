cls

# a word of warning...
Write-Host "================================================="
Write-Host "This is both experimental and a work in progress."
Write-Host "Remove all "" from your powershell code first"
Write-Host "to give it the best chance of succeeding"
Write-Host "================================================="
Write-Host ""

# file to convert
$fileName = ".\ExampleScriptToConvert.ps1"

# debug options
$showMinified = $false;
$showMshta = $false;

# ------------------------------------------------------------

function ConvertNumberToMinimizedSymbol([int] $number) {
  $letters = ''
  while ($number -gt 26) {
    $letters = [char](($number % 26) + 64) + $letters
    $number = [math]::Floor($number / 26)
  }
  $letters = [char]($number + 64) + $letters
  return $letters
}

function MakeTwinCATSafe([string] $string){

    # Make $ string safe in TwinCAT, replace $ with $$
    $string = $string -replace ('\$', '$$$');

    # Make ' string safe in TwinCAT, replace ' with $'
    $string = $string -replace ('''', '$$''');

    return $string
}


# ----------------------------------------------------------
# First step is to minimize the script as much as possible
# this will include removing new lines, spaces and replacing
# variables with short names
# ----------------------------------------------------------

# start the counter at 0
$variableCounter = 0;

# remove blank lines
$script = Get-Content $fileName;

# regex to find comments
$commentsRegex = "#.*"

# remove any comments
$script = $script -replace ($commentsRegex, "");

# remove empty lines
$script = ($script -replace "(?m)^\s*`r`n",'').trim();

# replace new lines with semicolon
$script = [string]::join(";",($script.Split("`n")));

# regex to find single or multicharacter variable names
$symbolNameRegex = "\$[A-Za-z][A-Za-z0-9_]*";

# find all symbol names (make unique and in decending size order)
$symbolNames = Select-String -Pattern $symbolNameRegex  -InputObject $script -AllMatches | foreach-object {$_.matches} | foreach-object {$_.groups[0].value} | Select-Object -Unique | Sort-Object -Property Length -Descending

# replace all symbols with short variable names
$symbolNames | ForEach-Object {

    $variableCounter = $variableCounter + 1;
    $replaceWith = ConvertNumberToMinimizedSymbol($variableCounter);
    $script = $script.replace($_, '$PowershellConvertToNTStartProcess_SuperSafeReplaceVariableNameWith_'+$replaceWith);
  
}

# remove safe replace variable
$script = $script -replace ('PowershellConvertToNTStartProcess_SuperSafeReplaceVariableNameWith_', "");

# regex to find equals which is surrounded by space
$equalsRegex = "\s*=\s*";

# remove any space around = symbol
$script = $script -replace ($equalsRegex, "=");

# remove extra end of lines
$script = $script -replace ';;', ';'
$script = $script -replace '; ;', ';'

if ($showMinified) { 
    Write-Host("Minimized Powershell");
    Write-Host("--------------------");
    Write-Host($script);
    Write-Host("");
}

# ----------------------------------------------------------
# Next we convert the script to be compatible with mshta
# this involves adding the command line call, plus escaping
# any quote marks
# ----------------------------------------------------------

$mshtaHeader = 'C:\Windows\System32\mshta vbscript:Execute("CreateObject(""WScript.Shell"").Run ""powershell -c ';
$mshtaFooter = '"", 0: window.close")';

# replace " with """"
$script = $script -replace ('"', '""""');

# combine with the header and footer
$script = $mshtaHeader + $script + $mshtaFooter;

if ($showMshta) { 
Write-Host("mshta script");
Write-Host("------------");
Write-Host($script);
Write-Host("");
}

# ----------------------------------------------------------
# Finally we escape the $ with $$ (this does not come out of
# the 255 limit
# ----------------------------------------------------------

# First check that the script is under 254 * 2 in lenght.
if ($script.Length -gt 254*2){
 Write-Host("Unable to conver script to NT_StartProcess as the total lines is " + ( $script.Length - (254 * 2) + " too long"));
 exit;
}

# now form the two parameters of NT_StartProcess
$pathStr = $script.SubString(0,[math]::min(254,$script.length));

if ($script.Length -gt 254){
    $comndLine = $script.SubString(255,$script.Length - 255);
} else {
    $comndLine = "";
}

$pathStr = MakeTwinCATSafe($pathStr);
$comndLine = MakeTwinCATSafe($comndLine);

$command = "ntStart(
	NetId := '',
	PathStr := '"+$pathStr+"',
	DirName := 'C:\Windows\System32',
	ComndLine := '"+$comndLine+"',
	Start := start,
	TmOut := DEFAULT_ADS_TIMEOUT
);";

# final output
Write-Host "TwinCAT code";
Write-Host "------------";
Write-Host $command -ForegroundColor Yellow;
Write-Host "";

# auto copy to clipboard
Set-Clipboard -Value $command;
Write-Host("TwinCAT code has been copied to your clipboard");