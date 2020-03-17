Param($offset)
<#
.SYNOPSIS
    Monitor print queus on a provided list of printservers
.DESCRIPTION
    This scripts monitors print queues on a list of printservers. The windows printqueue status and the IP port are tested and the result is rendered in a HTML file per printserver.
.PARAMETER offset
    This script will loop continuesly. To be able to schedule more than one run, an offset is introduced, resulting in a faster html output refresh cycle
.INPUTS
  Offset, text file with printservers
.OUTPUTS
  HTML file per printsserver
.NOTES
  Version:        1.0
  Author:         Bart Jacobs - @Cloudsparkle
  Creation Date:  16/03/2020
  Purpose/Change: Monitor printservers
  Original script by: Jason Poyner, jason.poyner@deptive.co.nz, techblog.deptive.co.nz
.EXAMPLE
  None
#>

$currentDir = Split-Path $MyInvocation.MyCommand.Path
# Feel free to specify another folder to store the html output files
$outputdir = $currentDir

# Get the printservers to monitor
$serverlistfilename = "printservers.txt"
$serverlistfile = Join-Path $currentDir $serverlistfilename

# Check if the file actually exists
$fileexists = test-path $serverlistfile
if ($fileexists -eq $false)
  {
    Write-Error ("The file " + $serverlistfile + "does not exist. Exiting...")
    exit 1
  }

#==============================================================================================
# The headers of the report
$headerNames  = "Printerstatus", "Ping"
$headerWidths = "4",              "4"

# ==============================================================================================
# ==                                       FUNCTIONS                                          ==
# ==============================================================================================

Function LogMe()
{
	Param(
		[parameter(Mandatory = $true, ValueFromPipeline = $true)] $logEntry,
		[switch]$display,
		[switch]$error,
		[switch]$warning,
		[switch]$progress
	)

	if ($error) {
		$logEntry = "[ERROR] $logEntry" ; Write-Host "$logEntry" -Foregroundcolor Red}
	elseif ($warning) {
		Write-Warning "$logEntry" ; $logEntry = "[WARNING] $logEntry"}
	elseif ($progress) {
		Write-Host "$logEntry" -Foregroundcolor Green}
	elseif ($display) {
		Write-Host "$logEntry" }

	#$logEntry = ((Get-Date -uformat "%D %T") + " - " + $logEntry)
	$logEntry | Out-File $logFile -Append
}

Function writeHtmlHeader
{
  param($title, $fileName)
  $date = ( Get-Date -format R)
  $head = @"
  <html>
  <head>
  <meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
  <title>$title</title>
  <STYLE TYPE="text/css">
  <!--
  td {
    font-family: Tahoma;
    font-size: 11px;
    border-top: 1px solid #999999;
    border-right: 1px solid #999999;
    border-bottom: 1px solid #999999;
    border-left: 1px solid #999999;
    padding-top: 0px;
    padding-right: 0px;
    padding-bottom: 0px;
    padding-left: 0px;
    overflow: hidden;
  }
  body {
    margin-left: 5px;
    margin-top: 5px;
    margin-right: 0px;
    margin-bottom: 10px;
    table {
      table-layout:fixed;
      border: thin solid #000000;
    }
    -->
    </style>
    </head>
    <body>
    <table width='1200'>
    <tr bgcolor='#CCCCCC'>
    <td colspan='7' height='48' align='center' valign="middle">
    <font face='tahoma' color='#003399' size='4'>
    <!--<img src="http://servername/administration/icons/xenapp.png" height='42'/>-->
    <! <strong>$title - $date</strong></font>
    </td>
    </tr>
    </table>
"@
$head | Out-File $fileName
}

Function writeTableHeader
{
  param($fileName)
  $tableHeader = @"
  <table width='1200'><tbody>
  <tr bgcolor=#CCCCCC>
  <td width='6%' align='center'><strong>PrinterName</strong></td>
"@

  $i = 0
  while ($i -lt $headerNames.count)
  {
	   $headerName = $headerNames[$i]
	   $headerWidth = $headerWidths[$i]
	   $tableHeader += "<td width='" + $headerWidth + "%' align='center'><strong>$headername</strong></td>"
	   $i++
   }
   $tableHeader += "</tr>"
   $tableHeader | Out-File $fileName -append
}

Function writeData
{
	param($data, $fileName)

	$data.Keys | sort | foreach
    {
      $tableEntry += "<tr>"
		  $PrinterName = $_
      $tableEntry += ("<td bgcolor='#CCCCCC' align=center><font color='#003399'>$PrinterName</font></td>")
      #$data.$_.Keys | foreach {
      $headerNames | foreach
        {
          #"$PrinterName : $_" | LogMe -display
			try
      {
				if ($data.$PrinterName.$_[0] -eq "SUCCESS") { $bgcolor = "#387C44"; $fontColor = "#FFFFFF" }
				elseif ($data.$PrinterName.$_[0] -eq "WARNING") { $bgcolor = "#FF7700"; $fontColor = "#FFFFFF" }
				elseif ($data.$PrinterName.$_[0] -eq "ERROR") { $bgcolor = "#FF0000"; $fontColor = "#FFFFFF" }
				else { $bgcolor = "#CCCCCC"; $fontColor = "#003399" }
				$testResult = $data.$PrinterName.$_[1]
			}
			catch
      {
				$bgcolor = "#CCCCCC"; $fontColor = "#003399"
				$testResult = ""
			}
			$tableEntry += ("<td bgcolor='" + $bgcolor + "' align=center><font color='" + $fontColor + "'>$testResult</font></td>")
    }
		$tableEntry += "</tr>"
  }
  $tableEntry | Out-File $fileName -append
}

Function Ping([string]$hostname, [int]$timeout = 500, [int]$retries = 3)
{
$result = $true
$ping = new-object System.Net.NetworkInformation.Ping #creates a ping object
$i = 0
do {
    $i++
		#write-host "Count: $i - Retries:$retries"

		try
    {
      #write-host "ping"
			$result = $ping.send($hostname, $timeout).Status.ToString()
    }
    catch
    {
			#Write-Host "error"
			continue
		}
		if ($result -eq "success") { return $true }

    } until ($i -eq $retries)
    return $false
}

# ==============================================================================================
# ==                                       MAIN SCRIPT                                        ==
# ==============================================================================================

# Check offset parameter, if not set, set to zero
If ($offset -eq $Null)
  {
    $offset = 0
  }

# Script loop
while ($true)
{
  "Sleeping " + $offset + "s" | LogMe -display -progress
  start-sleep -s $offset

  # Get Start Time
  $startDTM = (Get-Date)

  "Getting list of printservers..." | LogMe -display
  $printservers = Get-Content $serverlistfile

  # Check if the file provided is not empty
  $fileempty = if (get-content $serverlistfile -TotalCount 1) {$fales} else {$true}

  if ($fileempty -eq $true)
  {
    Write-Error ("The file " + $serverlistfile + " is empty. Exiting...")
    Exit 1
  }

  ForEach ($printserver in $printservers)
  {
    # Check if the server provided is online, simple ping test
    $serveronline = Ping $printserver 100
    if ($serveronline -eq $false)
    {
      ($printserver + " does not exist or is offline. Skipping...    ") | LogMe -display -warning
    }
    else
    {
      $resultfilename = $printserver + "_Results.htm"
      $errorfilename = $printserver + "_Errors.htm"
      $logfilename = 	$printserver + "_Logfile.log"
      $logfile    = Join-Path $outputDir $logfilename
      $resultsHTM = Join-Path $outputDir $resultfilename
      $errorsHTM  = Join-Path $outputDir $errorfilename

      "Remove logfile..." | LogMe -display
      rm $logfile -force -EA SilentlyContinue

      $allResults = @{}

      # Get shared printer list
      "Checking Printer status..." | LogMe -display
      Get-Printer -ComputerName $printserver | select Name, PrinterStatus, PortName, Shared, JobCount | where { $_.Shared -eq $True } | % {
        $tests = @{}	0
        $printer = $_.Name
        $printer | LogMe -display -progress

        # Check Printer Status
        if($_.PrinterStatus -eq "Error")
        {
          "Printer in Error State" | LogMe -display -error
		       $tests.PrinterStatus = "ERROR", "Error"
        }
        else
        {
		        $tests.PrinterStatus = "SUCCESS","Normal"
        }

        # Check Printer Port Status
        $printerip = Get-PrinterPort -ComputerName $printserver $_.PortName | select PrinterHostAddress
        $result = Ping $printerip.PrinterHostAddress 100
	      if ($result -ne "SUCCESS") {$tests.Ping = "ERROR", $result }
	      else {$tests.Ping = "SUCCESS", $result }

        $allResults.$printer = $tests
      }

      # Write all results to an html file
      Write-Host ("Saving results to html report: " + $resultsHTM)
      writeHtmlHeader $printserver $resultsHTM
      writeTableHeader $resultsHTM
      $allResults | sort-object -property FolderPath | % { writeData $allResults $resultsHTM }

      # Get End Time
      $endDTM = (Get-Date)

      # Echo Time elapsed
      "Elapsed Time: $(($endDTM-$startDTM).totalseconds) seconds"
    }
  }
}
