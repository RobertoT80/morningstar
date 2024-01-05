<#
.SYNOPSIS
    Morningstar fund performance automation
.DESCRIPTION
    Retrieves values from funds in MorningStar via http.
	The ID of the funds have to be provided in a text file, each in one line.
	The output will be written to the current directory as 'fund_list_<date>.csv'.
.EXAMPLE
	.\get-fundValues.ps1 -listFile 'C:\test\fund_list.txt'
    Retrieves values for the IDs of the funds listed in C:\test\fund_list.txt
.NOTES
    Author:  Roberto Toso
    Date:    05/01/2024
    Version: 1.0
	Requires: Selenium powershell module, Firefox browser
#>

[CmdletBinding()]
param(
	[Parameter(Mandatory=$true)]
	$listFile
)

$script:FUND_IDS = @()
$csv_output_name = [string]::format("fund_list_{0}.csv", (get-date -format yyyyMMdd) )
$csv_output_fullpath = [io.path]::combine($PWD, $csv_output_name)
[array] $csv_columns = @('web_id', 'name', 'isin', 'value_eur', 'var_1day_fund_%', 'var_4week_fund_%', 'var_1day_cat_%', 'var_4week_cat_%')

$ErrorActionPreference = 'stop'
set-strictmode -Version 2

function init()
{
	# Import selenium - both this and firefox must be installed
	try
	{
		import-module selenium
	}
	catch
	{
		write-error "Selenium module not found. Install with: 'install-module selenium'"
		exit
		
	}
	
	# Init header in csv output file
	set-content -value ($csv_columns -join '|') -path $csv_output_fullpath
	
	# Handle relative path if passed at console
	if([System.IO.Path]::IsPathRooted($listFile) -eq $false)
	{
		$script:listFile = [io.path]::combine($PWD, $listFile)
	}
	
	if((test-path $script:listFile -erroraction silentlycontinue) -eq $false)
	{
		write-host "File not found: $($script:listFile)" -foregroundcolor red
		exit 1
	}
	
	# Load list from input file
	try
	{
		$rdr = [io.file]::ReadAllLines($script:listFile)
		if($rdr.length -eq 0)
		{
			write-warning "File is empty: $($script:listFile)"
			exit 2
		}
		foreach($line in $rdr)
		{
			if($line -ne '') { $script:FUND_IDS += $line }
		}
		
	}
	catch
	{
		write-host "Cannot read file: $($script:listFile)" -foregroundcolor red
		exit 3
	}
	write-host ([string]::format("{0}{1}IDs found: {2}{1}{0}", (("=" * 14), "`n", $script:FUND_IDS.count)))
	
}

function askConfirmation([string] $msg = '')
{
   write-host $msg -foregroundcolor 'yellow'
   [string] $answer = ''

   while($answer.tolower() -notmatch 'yes|no')
   {
				  Write-Host 'Continue? (' -NoNewline
				  Write-Host 'yes' -NoNewline -ForegroundColor Green
				  Write-Host ', ' -NoNewline
				  Write-Host 'no' -NoNewline -ForegroundColor Red
				  Write-Host ')'
				  $answer = Read-Host
   }

   if($answer -eq 'no') { exit }
}

function get-seleniumDriver()
{
	kill -name firefox -erroraction silentlycontinue -force
    $seleniumOptions = New-Object OpenQA.Selenium.Firefox.FirefoxOptions
    $seleniumOptions.AddArgument('-headless')
    $seleniumDriver = New-Object OpenQA.Selenium.Firefox.FirefoxDriver -ArgumentList @($seleniumOptions)
    return $seleniumDriver
}

function output-csv($web_id)
{
	write-host "Retrieving data for fund: '$web_id'" -foregroundcolor green -backgroundcolor black
	$fund = (new-fund $web_id)
	
	$csv_columns_values = @()
	foreach($column in $csv_columns)
	{
		$csv_columns_values += ($fund.$column)
	
	}
	add-content ($csv_columns_values -join '|') -path $csv_output_fullpath
	write-host "=> Fund ID '$web_id' exported to $csv_output_fullpath`n" -foregroundcolor green

}


function new-fund($web_id)
{
	$seleniumDriver = get-seleniumDriver
	$url = "https://quantalys.it/Fonds/$($web_id)"
	
	try
	{
		$seleniumDriver.url = "https://quantalys.it/Fonds/$($web_id)"
	}
	catch
	{
		write-error "URL unreachable: $url"
		exit
	}
	
	if($seleniumDriver.Title -eq 'The resource cannot be found.')
	{
		write-error "Page not found: $url"
		exit
	}
	
    $fund = new-object System.Collections.Specialized.OrderedDictionary
	$fund.add('web_id', $web_id)
	try
	{
		$name = $seleniumDriver.title.split('|')[0]
		$fund.add('name', $name)
		$isin = $seleniumDriver.title.split('|')[-1].trim().split()[0]
		$fund.add('isin', $isin)
		$value_eur = $seleniumDriver.FindElementsByClassName('vl-box-converted-value').text.split('=')[1].split('EUR')[0].trim()
		$fund.add('value_eur', $value_eur)

		$var_1day_fund = $seleniumDriver.FindElementsByCssSelector('table')[0].text.split()[8]
		$fund.add('var_1day_fund_%', $var_1day_fund)
		$var_4week_fund = $seleniumDriver.FindElementsByCssSelector('table')[0].text.split()[16]
		$fund.add('var_4week_fund_%', $var_4week_fund)

		$var_1day_cat = $seleniumDriver.FindElementsByCssSelector('table')[0].text.split()[10]
		$fund.add('var_1day_cat_%', $var_1day_cat)
		$var_4week_cat = $seleniumDriver.FindElementsByCssSelector('table')[0].text.split()[18]
		$fund.add('var_4week_cat_%', $var_4week_cat)
	}
	catch
	{
		write-error "Page in unexpected format: $url"
		exit
	}

	write-verbose "Data retrieved for fund ID: $web_id"
	
	$seleniumDriver.quit()
	
	return $fund

}

function main()
{
	askConfirmation "This script will terminate any Firefox sessions you might have open."
	
	init
	
	for($i = 0; $i -lt $script:FUND_IDS.count; $i++)
	{
		write-progress ([string]::format("Retrieving data... ({0}/{1})", $i+1, $script:FUND_IDS.count))
		output-csv $script:FUND_IDS[$i];
	}
	
}

main