﻿<# 
.SYNOPSIS
Remove a column everwhere it appears in a site collection.

.DESCRIPTION
This script iterates through all webs in a site collection and removes a specified column wherever it is found. 
It will create date-stamped transcript and log files in the directory it is executed from.

.PARAMETER xmlfile
Refers the path to an .XML file which contains parameters for individual sites being created. Assure your parameters file
is located in the same directory as the .ps1 file.

.EXAMPLE
From the PowerShell prompt run .\RemoveColumn.ps1 -xmlfile .\[enter the name of your parameters file].xml

.NOTES
15-02-19 Assembled by Ramona Maxwell - Microsoft Public License (Ms-PL)

.LINK
http://www.microsoft.com/en-us/openness/licenses.aspx
#>

Param([string]$xmlfile) 
$ErrorActionPreference = "Inquire"
$dateStamp = Get-Date -Format "yyyy-MM-dd-hhmm"
$transcript = ".\" + $dateStamp + "_" + $MyInvocation.MyCommand.Name.Replace(".ps1", "") + ".log"
$validateFile = ".\" + $dateStamp + "_" + $MyInvocation.MyCommand.Name.Replace(".ps1", "") + "_validate.doc"
$validateCSV = ".\" + $dateStamp + "_" + $MyInvocation.MyCommand.Name.Replace(".ps1", "") + "_validate.csv"
$removeFile = ".\" + $dateStamp + "_" + $MyInvocation.MyCommand.Name.Replace(".ps1", "") + "_remove.doc"
Start-Transcript -Path $transcript

function Get-Parameters(){
	<#
	.SYNOPSIS
	Loads and validates the XML file containing the script parameters.

	.DESCRIPTION
	This function verifies the configuration file exists and is loaded. It then returns the parameters from the file at the node specified.	
	#>	
	if (-not(Test-Path $xmlfile)){ 
		Write-Warning "The file containing script parameters cannot be located."
		Write-Host "The file to create the script parameters must be located in the same directory as the calling script." -ForegroundColor:Green
		return $false
		}
	#Load and validate the configuration file
	[xml]$scriptParams = Get-Content $xmlfile
	$siteParams = $scriptParams.params.sites.site
		if($siteParams -eq $null){
			Write-Warning "Column removal parameters are not present in the parameters file."
			Write-Host "Ensure that the parameters file contains a <site> node." -ForegroundColor:DarkMagenta
			return $false
		}		
	return $siteParams	
}

function Validate-Columns () {
Write-Host "The following site collections have been specified for this operation: " `n
Get-Parameters | ForEach-Object {
$params = $_
$site = Get-SPSite ([System.Convert]::ToString([string]$params.URL))
Write-Host $site
$webs = $site.AllWebs
$columnName = ([System.Convert]::ToString([string]$params.column2.Name))
$GUID1 = $params.column1.Id
$GUID2 = $params.column2.Id
$count = 0   
#$count1 = 0
$count2 = 0
$errorCount = 0
#$errorCount1 = 0

Write-Output "Operation started at " $dateStamp
try {
foreach($web in $webs) {  
	If ($web -ne $null) {
  foreach($list in $web.Lists)  {
		If ($list -ne $null) {
		$fieldCount = $list.Fields.Count
    for ($c = 0; $c -le $fieldCount - 1; $c++) {
		If ($list.Fields[$c].InternalName -Contains $columnName -or $list.Fields[$c].Title -Contains $columnName) {              
			$count++  
			New-Object -TypeName PSCustomObject -Property @{
				Count = $count
				Website = $web.Url
				List = $list.Title
				Field = $list.Fields[$c].Title
				FieldID = $list.Fields[$c].Id
				ReadOnly = $list.Fields[$c].ReadOnlyField
				AllowDeletion = $list.Fields[$c].AllowDeletion
				Sealed = $list.Fields[$c].Sealed
				Hidden = $list.Fields[$c].Hidden
			} | Export-Csv –Path $validateCSV -Append -NoTypeInformation
	        Write-Output "The website for this operation is : `n" $web.Url " `n"
			Write-Output "Found match number " $count "in the " $list.Fields[$c].Title " field in the " $list.Title " list. `n"
			Write-Output "The field has an ID of: " $list.Fields[$c].Id " `n" " and its read-only property is " $list.Fields[$c].ReadOnlyField " `n "
			Write-Output "The Sealed property of the field is: " $list.Fields[$c].Sealed " `n "
			Write-Output "The AllowDeletion property of the field is: " $list.Fields[$c].AllowDeletion " `n "
			Write-Output "The Hidden property of the field is: " $list.Fields[$c].Hidden " `n "
		}
    }
  } 
}
}
}
}
catch  [System.Management.Automation.PSArgumentException]{
	$errorCount ++
	Write-Host $_.Exception.Message " was found in: " $list.Title " list." -ForegroundColor:Red | Format-Table -AutoSize
}
finally {
	if ($errorCount -eq 0){
		Write-Output "All column instances were located."
	}        
	echo "Finished column audit task."
	$site.Dispose()
	Sleep 3
	
}
}
Menu
}

function Remove-Columns() {
Write-Host "The following site collections have been specified for this operation: " `n
Get-Parameters | ForEach-Object {
$params = $_
$site = Get-SPSite ([System.Convert]::ToString([string]$params.URL))
Write-Host $site
$webs = $site.AllWebs
$columnName = ([System.Convert]::ToString([string]$params.column2.Name))
$GUID1 = $params.column1.Id
$GUID2 = $params.column2.Id
#$count = 0   
$count1 = 0
$count2 = 0
#$errorCount = 0
$errorCount1 = 0
try {
	$siteCount = $site.Count
	for ($z = 0; $z -le $siteCount -1; $z++) {
      $webCount = $webs.Count
      for($a = 0; $a -le $webCount - 1; $a++) {  
            If ($webs[$a] -ne $null) {
                  $listCount = $webs[$a].Lists.Count
                  for($b = 0; $b -le $listCount - 1; $b++) {
                        If ($webs[$a].Lists[$b] -ne $null) {
                              $fieldCount2 = $webs[$a].lists[$b].Fields.Count
                              for ($c = 0; $c -le $fieldCount2 - 1; $c++) { 
					$list = $webs[$a].Lists[$b]
	                            	if(($list.Fields[$c].Id -eq $GUID1) -or ($list.Fields[$c].Id -eq $GUID2)) {  
	                                        $count1++           
	                                        Write-Output "The website for this operation is : `n" $webs[$a].Url " `n"
	                                        Write-Output  "Now removing match number " $count1 " of the " $list.Fields[$c].Title " field from the " $list.Title " list. `n"
	                                        $list.Fields[$c].Hidden = $false
						$list.Fields[$c].AllowDeletion = $true
						$list.Fields[$c].Sealed = $false
	                                        $list.Update()
						$list.Fields.Delete($list.Fields[$c].InternalName)  
					}  
                                    }     
                        	}
                            $list.Update()
                        }
                  }
		        # Remove the field itself
		    	if($webs[$a].Fields.ContainsFieldWithStaticName($columnName)) {
        			Write-Output “Attempting to remove field:” $columnName -ForegroundColor DarkGreen
        			$webs[$a].Fields.Delete($columnName)
    			}
            } 
		}
      }
catch  [System.Management.Automation.PSArgumentException]{
	$errorCount1 ++
	Write-Host " Web: " $webs[$a].Url " List: " $list.Title $_.Exception.Message -ForegroundColor:Red | Format-Table -AutoSize
}
finally {
	if ($errorCount1 -eq 0){
		Write-Output "All column instances requiring removal were located."
	}
	else {
		Write-Output "Some or all of the column instances were not located and could not be removed."
		Write-Output "The error count is " $errorCount1 " `n"
	}            
	Write-Output "Finished column removal task." $count1 " fields were deleted."
	$dateStamp2 = Get-Date -Format "yyyy-MM-dd-hhmm"
	Write-Output "Operation completed at " $dateStamp2
	Sleep 3

	}
}
	Menu
}

function VerifyExit() {
		$VerifyExit = read-host "Are you sure you want to exit? (y/n)"  
        if (($VerifyExit -eq "y") -or ($VerifyExit -eq "Y")){
		Stop-Transcript
		Start-Sleep -Seconds 3
		exit
		}  
        if (($VerifyExit -eq "n") -or ($VerifyExit -eq "N")){Menu}  
        else {
			write-host -ForegroundColor:Red "Please select y to exit or n to continue."   
            VerifyExit  
        }  
} 

function VerifyValidation() {
		$VerifyValidation = read-host "This function will discover every instance of the column specified in your parameters file and record it in a text file. It may take some time to execute. Run now? (y/n)"  
        if (($VerifyValidation -eq "y") -or ($VerifyValidation -eq "Y")){
			Write-Host "Column discovery has started. Do not close the console window. You will be returned to the main menu when discovery is complete."
			Validate-Columns | Out-File -FilePath $validateFile -Append
			}  
        if (($VerifyValidation -eq "n") -or ($VerifyValidation -eq "N")){Menu}  
		else {
		write-host -foregroundcolor red "Please select y to perform column discovery or n to return to the main menu."   
		VerifyValidation
		}
          
}

function VerifyRemoval() {
		$VerifyRemoval = read-host "This script will remove the column you specify in lists and libraries throughout a given site collection. If you specify the wrong column name or GUID you can break your farm!! Are you sure you wish to execute this column removal function? (y/n)"  
        if (($VerifyRemoval -eq "y") -or ($VerifyRemoval -eq "Y")){
			Write-Host "Column removal has started. Do not close the console window. You will be returned to the main menu when removal is complete."
			Remove-Columns | Out-File -FilePath $removeFile -Append
			}  
        if (($VerifyRemoval -eq "n") -or ($VerifyRemoval -eq "N")){Menu}  
        else {
			write-host -foregroundcolor red "Please select y to permanently remove column or n to return to the main menu."   
            VerifyRemoval  
        }  
}

function Menu() {
	Write-Host "---------------------------------------------------------"   
	Write-Host ""   
	Write-Host "    1. Discover where column exists in farm." 
	Write-Host "    2. Remove column throughout the farm."
	Write-Host "    3. Exit"  
	Write-Host ""   
	Write-Host "---------------------------------------------------------"  
	$answer = read-host "Please select an option."   
	if ($answer -eq 1){VerifyValidation}  
	if ($answer -eq 2) {VerifyRemoval}
	if ($answer -eq 3){VerifyExit}  
	else {
		write-host -ForegroundColor red "Invalid Selection"  
    	sleep 5  
    	Menu
	}
}
Menu
$site.dispose()
Stop-Transcript
