# elisabeth@sqlserverland.com
# 2018-12-27
# Parses SQLIO output files in specified directory ($InputDirectory) into an Excel workbook
# Credits to Jonathan Kehayias for code that parses raw SQLIO output

# Parameters:
# $Drive - drive letter, i.e. "E" (w/o quotation marks)


# Hardcoded (in "Set paths"):
# $InputDirectory - location of result files from SQLIO

# Leaves you with one open Excel workbook at the end of execution


param(	
[Parameter(Mandatory=$TRUE)]
		[ValidateNotNullOrEmpty()]
		[string] $Drive
        )

Set-Location $PSScriptRoot

# Set path
$InputDirectory = $PSScriptRoot + "\SQLIO_ResultFiles\"


#Set up Excel Workbook to Store data
$Excel = New-Object -ComObject Excel.Application
$SaveFormat="xlExcel12"
$Excel.Visible = $true
$WorkBook = $Excel.WorkBooks.Add()
$WorkBook.WorkSheets.Item(1).Name = $Drive
$WorkSheet = $WorkBook.WorkSheets.Item($Drive)

$WorkSheet.Cells.Item(1,1) = "Threads"
$WorkSheet.Cells.Item(1,2) = "Operation"
$WorkSheet.Cells.Item(1,3) = "Duration"
$WorkSheet.Cells.Item(1,4) = "IOSize"
$WorkSheet.Cells.Item(1,5) = "IOType"
$WorkSheet.Cells.Item(1,6) = "PendingIO"
$WorkSheet.Cells.Item(1,7) = "FileSize"
$WorkSheet.Cells.Item(1,8) = "IOPS"
$WorkSheet.Cells.Item(1,9) = "MBs/Sec"
$WorkSheet.Cells.Item(1,10) = "Min_Lat(ms)"
$WorkSheet.Cells.Item(1,11) = "Avg_Lat(ms)"
$WorkSheet.Cells.Item(1,12) = "Max_Lat(ms)"
$WorkSheet.Cells.Item(1,13) = "Caption"
$WorkSheet.Cells.Item(1,14)= "Drive"

# Loop through all SQLIO output files in $InputDirectory

# Row counter
$x = 2

foreach ($f in Get-ChildItem -path $InputDirectory | sort-object -desc )
    
	{ 
		Write-Host "Processing "  $f 
		Write-Host $InputDirectory$f
        Try

        {

$filedata = [string]::Join([Environment]::NewLine,(Get-Content $InputDirectory$f))

$Results = $filedata.Split( [String[]]"sqlio v1.5.SG", [StringSplitOptions]::RemoveEmptyEntries ) | `
		 select @{Name="Threads"; Expression={[int]([regex]::Match($_, "(\d+)?\sthreads\s(reading|writing)").Groups[1].Value)}},`
				@{Name="Operation"; Expression={switch ([regex]::Match($_, "(\d+)?\sthreads\s(reading|writing)").Groups[2].Value)
												{
													"reading" {"Read"} 
													"writing" {"Write"}
												}	}},`
				@{Name="Duration"; Expression={[int]([regex]::Match($_, "for\s(\d+)?\ssecs").Groups[1].Value)}},`
				@{Name="IOSize"; Expression={[int]([regex]::Match($_, "\tusing\s(\d+)?KB\s(sequential|random)").Groups[1].Value)}},`
				@{Name="IOType"; Expression={switch ([regex]::Match($_, "\tusing\s(\d+)?KB\s(sequential|random)").Groups[2].Value)
												{
													"random" {"Random"} 
													"sequential" {"Sequential"}
												}  }},`
				@{Name="PendingIO"; Expression={[int]([regex]::Match($_, "with\s(\d+)?\soutstanding").Groups[1].Value)}},`
				@{Name="FileSize"; Expression={[int]([regex]::Match($_, "\s(\d+)?\sMB\sfor\sfile").Groups[1].Value)}},`
				@{Name="IOPS"; Expression={[decimal]([regex]::Match($_, "IOs\/sec\:\s+(\d+\.\d+)?").Groups[1].Value)}},`
				@{Name="MBs_Sec"; Expression={[decimal]([regex]::Match($_, "MBs\/sec\:\s+(\d+\.\d+)?").Groups[1].Value)}},`
				@{Name="MinLat_ms"; Expression={[int]([regex]::Match($_, "Min.{0,}?\:\s(\d+)?").Groups[1].Value)}},`
				@{Name="AvgLat_ms"; Expression={[int]([regex]::Match($_, "Avg.{0,}?\:\s(\d+)?").Groups[1].Value)}},`
				@{Name="MaxLat_ms"; Expression={[int]([regex]::Match($_, "Max.{0,}?\:\s(\d+)?").Groups[1].Value)}}`
	 | Sort-Object IOSize, IOType, Operation, Threads 

#Write data from file into spreadsheet

$Results | % {
	$WorkSheet.Cells.Item($x,1) = $_.Threads
	$WorkSheet.Cells.Item($x,2) = $_.Operation
	$WorkSheet.Cells.Item($x,3) = $_.Duration
	$WorkSheet.Cells.Item($x,4) = $_.IOSize
	$WorkSheet.Cells.Item($x,5) = $_.IOType
	$WorkSheet.Cells.Item($x,6) = $_.PendingIO
	$WorkSheet.Cells.Item($x,7) = $_.FileSize
	$WorkSheet.Cells.Item($x,8) = $_.IOPS
	$WorkSheet.Cells.Item($x,9) = $_.MBs_Sec
	$WorkSheet.Cells.Item($x,10) = $_.MinLat_ms
	$WorkSheet.Cells.Item($x,11) = $_.AvgLat_ms
	$WorkSheet.Cells.Item($x,12) = $_.MaxLat_ms
	$WorkSheet.Cells.Item($x,13) = [string]$_.IOSize + "KB " + [string]$_.IOType + " " + `
								[string]$_.Operation + " " + [string]$_.Threads + `
								" Threads " + [string]$_.PendingIO + " pending"
    $WorkSheet.Cells.Item($x,14) = $Drive

	$x++

    
    }


        }
Catch
		 {
            Set-Location $PSScriptRoot
             $myErr = "$(get-date -Format u) $_ running $f" 
             write-host $myErr
             $myErr| Out-File ScriptExec_Errors_$(get-date -f yyyy-MM-dd).txt -append
                                             
             continue
            }

	}
    
    	 #handle  failures
       		 trap
		 {
               		 "Something went wrong; $error , $_ running $f" ;
               		 continue
            		 }
       

#$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault


