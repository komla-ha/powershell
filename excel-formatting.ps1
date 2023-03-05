


$chd_hashCustom = @{
		'HourOfDay'    = 'Hour';
		'DayOfWeek'    = 'DayOfWeek';
		'DayOfYear'    = 'DayOfYear';
		'Year'		   = 'Year';
		'Month_Id'	   = 'Month';
		'Month'	       = 'CustomMMMM';
		'DayLong'	   = 'CustomMMM-dd ddd'
	}
	




#region - UTILITIES => RSPOOL - JOBS_TRACKERS


		[runspacefactory]::CreateRunspacePool()
		$script:sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
		$script:rsPool = [runspacefactory]::CreateRunspacePool(1, 4)
		
		$script:powershell = [powershell]::Create()
		$script:powershell.RunspacePool = $script:rsPool
		
		$script:rsPool.Open()
		$script:jobs_tracker = @()


#endregion



#region => EXPORT-XLSX-MACROS

function chd-gui-saveReport-ExportToXLSX
{
	param ($data,
		$sheetName,
		$filePath)
	
	
	$Excel = New-Object -ComObject excel.application
	$Excel.visible = $false
	
	#$workbook = $Excel.workbooks.Open($filePath)
	$workbook = $Excel.workbooks.Add()
	
	$tmp = [System.IO.Path]::GetTempFileName()
	$data | Export-Csv $tmp -NoTypeInformation # -Delimiter "`t"
	
	$ws = $workbook.Worksheets.Add()
	$ws.Name = $sheetName
	$ws.Activate()
	
	$TxtConnector = ("TEXT;" + $tmp)
	$Connector = $ws.QueryTables.add($TxtConnector, $ws.Range("A1"))
	$query = $ws.QueryTables.item($Connector.name)
	
	### Set the delimiter (, or ;) according to your regional settings
	$query.TextFileOtherDelimiter = $Excel.Application.International(5)
	
	$query.Refresh()
	$query.Delete()
	
	Remove-Item $tmp -Force
	
	$workbook.SaveAs($filePath)
	$Excel.Workbooks.Close()
	$Excel.Quit()
	[System.GC]::Collect()
	[System.GC]::WaitForPendingFinalizers()
	
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
	Remove-Variable Excel
	
	
}



function chd-gui-excel-Report-MainFlow-Conditional-Macros
{
	param ($filename,
		$footer,
		$percentage = $false,
		$duration = $false,
		$cellCenter = $true)
	
	
	$scriptBlock = {
		
		param ($filename,
			$footer,
			$percentage,
			$duration,
			$cellCenter)
		
		function gui-excel-color-RGB
		{
			param ($red,
				$green,
				$blue)
			
			return [System.Double]($red + $green * 256 + $blue * 256 * 256)
		}
		
		
		$Excel = New-Object -ComObject excel.application
		$Excel.visible = $true
		
		$workbook = $Excel.Workbooks.Open($filename)
		$workbook.DefaultTableStyle = "TableStyleLight16"
		
		$Excel.CalculateBeforeSave = $true
		$Excel.Calculation = [Microsoft.Office.Interop.Excel.XlCalculation]::xlCalculationManual
		
		
		#$sheetNames = $workbook.Sheets | Select-Object -ExpandProperty Name | Where-Object { $_ -notmatch "Sheet" }
		#$sheet = $workbook.Sheets.Item(1).Activate()
		
		$workbook.Sheets | Where-Object { !($_.Name -match "sheet") } | ForEach-Object {
			
			$sheet = $_
			
			#$sheet }
			#$sheet.Name | Out-Host
			$sheet.Activate()
			$Excel.Activewindow.Zoom = 80
			
			### SHEET PAGE SETUP
			$sheet.PageSetup.Orientation = 2
			$sheet.PageSetup.CenterHorizontally = $True
			$sheet.PageSetup.CenterVertically = $True
			$sheet.PageSetup.PaperSize = 5
			$sheet.PageSetup.Zoom = $false
			$sheet.PageSetup.FitToPagesWide = 1
			$sheet.PageSetup.FitToPagesTall = 1
			$sheet.PageSetup.TopMargin = 1.90
			$sheet.PageSetup.RightMargin = 0.65
			$sheet.PageSetup.BottomMargin = 1.90
			$sheet.PageSetup.LeftMargin = 0.65
			
			$sheet.PageSetup.HeaderMargin = 0.75
			$sheet.PageSetup.FooterMargin = 0.75
			
			
			### TABLE FORMATING
			$listObject = $sheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $Excel.ActiveCell.CurrentRegion, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
			#$listObject.Name = 'TableData'
			$listObject.TableStyle = "TableStyleLight16"
			
			### TABLE RANGE / AREA
			$usedRange = $sheet.UsedRange
			$columnsCount = $usedRange.Columns.Count
			$rowsCount = $usedRange.Rows.Count
			
			
			#$linen = gui-excel-color-RGB 185 181 181
			$orangeColor = gui-excel-color-RGB 255 235 156
			$linen = gui-excel-color-RGB 128 128 128
			$linenBackcolor = gui-excel-color-RGB  217 217 217
			$yellowColor = gui-excel-color-RGB 255 255 0
			
			### CONDITIONAL CELLS
			if ($sheet.Name -match "Count")
			{
				
				### RANGE 0900-1900 BORDER THICKNESS
				$innerArray = $usedRange.Range($usedRange.Cells.Item(9, 2), $usedRange.Cells.Item(20, $columnsCount))
				$innerArray.BorderAround(6, 4, 3)
				$innerArrayAddress = $innerArray.Address($false, $false)
				
				
				$sheet.Cells.Item($rowsCount + 1, $columnsCount + 1) = '=SUM(' + $innerArrayAddress + ')'
				$sheet.Cells.Item($rowsCount + 1, $columnsCount + 1).BorderAround(6, 4, 3)
				$sheet.Cells.Item($rowsCount + 1, $columnsCount + 1).Interior.Color = $yellowColor
				
				
				#"Count" | Out-Host
				$innerConditions = $usedRange.Range($usedRange.Cells.Item(2, 2), $usedRange.Cells.Item($rowsCount, $columnsCount))
				
				##### CONDITIONAL FORMATTING
				$condition = '-'
				$xlType = [Microsoft.Office.Interop.Excel.xlFormatConditionType]::xlCellValue
				$xlOperator = [Microsoft.Office.Interop.Excel.XlFormatConditionOperator]::xlEqual
				$conditionFormat = $innerConditions.FormatConditions.Add($xlType, $xlOperator, $condition)
				$conditionFormat.Interior.Color = $linenBackcolor
				
				
				$condition = 0
				$xlType = [Microsoft.Office.Interop.Excel.xlFormatConditionType]::xlCellValue
				$xlOperator = [Microsoft.Office.Interop.Excel.XlFormatConditionOperator]::xlGreater
				$conditionFormat = $innerConditions.FormatConditions.Add($xlType, $xlOperator, $condition)
				$conditionFormat.Interior.Color = $orangeColor
				
				
				
				
				##### COLUMNS
				$column = 2
				$firstCell = $usedRange.Cells.Item(2, $column).Address($false, $false)
				$lastCell = $usedRange.Cells.Item($rowsCount, $column).Address($false, $false)
				$sheet.Cells.Item($rowsCount + 2, $column) = '=SUM(' + $firstCell + ':' + $lastCell + ')'
				$sheet.Cells.Item($rowsCount + 2, $column).Font.Size = 50
				
				$sourceAddress = $sheet.Cells.Item($rowsCount + 2, $column).Address($false, $false)
				$sourceRange = $sheet.Range($sourceAddress)
				
				$destinationAddress = $sheet.Cells.Item($rowsCount + 2, $columnsCount).Address($false, $false)
				$destinationRange = $sheet.Range($sourceAddress + ':' + $destinationAddress)
				$sourceRange.AutoFill($destinationRange, 0)
				
				##### ROWS
				$row = 2
				$firstCell = $usedRange.Cells.Item($row, 2).Address($false, $false)
				$lastCell = $usedRange.Cells.Item($row, $columnsCount).Address($false, $false)
				$sheet.Cells.Item($row, $columnsCount + 2) = '=SUM(' + $firstCell + ':' + $lastCell + ')'
				$sheet.Cells.Item($row, $columnsCount + 2).Font.Size = 50
				
				$sourceAddress = $sheet.Cells.Item($row, $columnsCount + 2).Address($false, $false)
				$sourceRange = $sheet.Range($sourceAddress)
				
				$destinationAddress = $sheet.Cells.Item($rowsCount, $columnsCount + 2).Address($false, $false)
				$destinationRange = $sheet.Range($sourceAddress + ':' + $destinationAddress)
				$sourceRange.AutoFill($destinationRange, 0)
				
				
				
				### TOTAL TOTAL 
				$firstCell = $usedRange.Cells.Item(2, $columnsCount + 2).Address('+True, False+')
				$lastCell = $usedRange.Cells.Item($rowsCount, $columnsCount + 2).Address('+True, False+')
				
				$sheet.Cells.Item($rowsCount + 2, $columnsCount + 2) = '=SUM(' + $firstCell + ':' + $lastCell + ')'
				$sheet.Cells.Item($rowsCount + 2, $columnsCount + 2).Font.Size = 60
				$sheet.Cells.Item($rowsCount + 2, $columnsCount + 2).BorderAround(1, 4, 0)
				$sheet.Cells.Item($rowsCount + 2, $columnsCount + 2).Interior.Color = $yellowColor
				
				$workbook.Save()
				
			}			
			
			$usedRange = $sheet.UsedRange
			
			if ($cellCenter)
			{
				### ALIGN VERTICAL + HORIZONTAL
				$usedRange.EntireColumn.HorizontalAlignment = [Microsoft.Office.Interop.Excel.XlHAlign]::xlHAlignCenter
				$usedRange.EntireColumn.VerticalAlignment = [Microsoft.Office.Interop.Excel.XlVAlign]::xlVAlignCenter
				
				### FIT ROWS, COLUMNS, HEIGHT
				#$usedRange = $sheet.UsedRange
				$usedRange.Font.Size = 65
				$row_first = $usedRange.Rows.Item(1)
				$row_first.WrapText = $true
				$usedRange.EntireRow.AutoFit() | Out-Null
				$usedRange.Cells.RowHeight = 160
				$usedRange.EntireColumn.AutoFit() | Out-Null
				
				$Excel.Activewindow.Zoom = 15
				
				### FIT COLUMN WIDTH
				if (!($usedRange.ColumnWidth -match "\d"))
				{
					$usedRange.ColumnWidth = 55
				}
				
				
				do
				{
					$usedRange.ColumnWidth = $usedRange.ColumnWidth - 5
					#$usedRange.ColumnWidth | Out-Host
					
				}
				while ($sheet.PageSetup.Pages.Count -ne 1)
				
				### FOOTER 
				$sheet.Cells.Item($rowsCount + 5, 2) = $footer
				$sheet.Cells.Item($rowsCount + 5, 2).Font.Size = 80
				$sheet.Cells.Item($rowsCount + 5, 2).Font.ColorIndex = 5
				$sheet.Cells.Item($rowsCount + 5, 2).HorizontalAlignment = -4131
				
				$mergeCells = $sheet.Range($sheet.Cells.Item($rowsCount + 5, 2), $sheet.Cells.Item($rowsCount + 5, 10)) #$sheet.Cells.Item($rowsCount + 5, 2).BorderAround(1,2,5)
				$mergeCells.Select()
				$MergeCells.MergeCells = $true
				
				$Excel.Activewindow.Zoom = 15
			}
			else
			{
				### FIT ROWS, COLUMNS, HEIGHT
				$row_first = $usedRange.Rows.Item(1)
				$row_first.WrapText = $true
				$usedRange.EntireRow.AutoFit() | Out-Null
				$usedRange.EntireColumn.AutoFit() | Out-Null
				
				### FOOTER 
				$sheet.Cells.Item($rowsCount + 5, 2) = $footer
				$sheet.Cells.Item($rowsCount + 5, 2).Font.ColorIndex = 5
				$sheet.Cells.Item($rowsCount + 5, 2).HorizontalAlignment = -4131
				
				$mergeCells = $sheet.Range($sheet.Cells.Item($rowsCount + 5, 2), $sheet.Cells.Item($rowsCount + 5, 10)) #$sheet.Cells.Item($rowsCount + 5, 2).BorderAround(1,2,5)
				$mergeCells.Select()
				$MergeCells.MergeCells = $true
				
				$Excel.Activewindow.Zoom = 80
				
			}
			
			
			$sheet.Application.ActiveWindow.SplitColumn = 1
			$sheet.Application.ActiveWindow.FreezePanes = $true
			$sheet.Cells.Item(1, 1).Select()
			
		}
		
		
		$Excel.Calculation = [Microsoft.Office.Interop.Excel.XlCalculation]::xlCalculationAutomatic
		$workbook.Sheets.Item(1).Activate()
		$workbook.Save()
		$workbook.Close($false)
		$Excel.Quit()
		[System.GC]::Collect()
		[System.GC]::WaitForPendingFinalizers()
		
		[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
		Remove-Variable Excel
		
		Start-Process $filename
	}
	
	
	$hashArguments = @{
		filename	 = $filename
		footer	     = $footer
		percentage   = $percentage
		duration	 = $duration
		cellCenter   = $cellCenter
		
	}
	
	
	$job = [powershell]::Create().AddScript($scriptBlock).AddParameters($hashArguments)
	$job.RunspacePool = $script:rsPool
	$script:jobs_tracker += New-Object psobject -Property @{
		Category	 = 'ExcelMacros'
		Pipe		 = $job
		Result	     = $job.BeginInvoke()
	}
	
}


#endregion



#region Function => DATA-MASSAGE-OBJECT


function data-convert-DateTime
{
	
	param ($string)
	
	if ($string -match "\d")
	{
		
		$string = [datetime]$string
		
	}
	
	return $string
}


function data-convert-TimeSpan
{
	
	param ($string)
	
	if ($string -match "\d")
	{
		
		$string = [timespan]$string
		
	}
	
	return $string
	
}


function data-Massage-Object
{
	
	param ($collection,
		[array]$datetimeProp,
		[array]$durationProp,
		$progress = (New-Object System.Windows.Forms.ProgressBar))
	
	$progress.Maximum = $collection.count
	$progress.Step = 1
	$progress.Value = 0
	
	foreach ($obj in $collection)
	{
		
		foreach ($dtp in $datetimeProp)
		{
			if ($obj.$dtp -match "\w")
			{
				#"Good data : $($obj.$dtp)" | Out-Host
				$obj.$dtp = data-convert-DateTime -string $obj.$dtp
			}
			else
			{
				#"Bad Data : $($obj.$dtp)" | Out-Host
			}
		}
		
		foreach ($dp in $durationProp)
		{
			if ($obj.$dtp -match "\w")
			{
				#"Good data : $($obj.$dtp)" | Out-Host
				$obj.$dp = data-convert-TimeSpan -string $obj.$dp
			}
			else
			{
				#"Bad Data : $($obj.$dtp)" | Out-Host
			}
			
		}
		
		$progress.PerformStep()
		#"$($progress.Value) : $($progress.Maximum)" | Out-Host
	}
	
	$progress.Value = 0
}


function data-Massage-Object-Custom
{
	
	param ($collection,
		$referenceTimeProp = 'Call Start Time',
		$hashTable,
		$progress = (New-Object System.Windows.Forms.ProgressBar))
	
	$progress.Maximum = $collection.count
	$progress.Step = 1
	$progress.Value = 0
	
	
	
	foreach ($obj in $collection)
	{
		$referenceTime = $obj.$referenceTimeProp
		
		foreach ($hashKey in $hashTable.Keys)
		{
			#$hashKey }
			$hkey = $hashKey
			$vkey = $hashTable[$hkey]
			
			if ($vkey -match "Custom")
			{
				$vkey = $vkey -replace "Custom", ""
				$value = $referenceTime.ToString("$vkey")
				$obj | Add-Member -MemberType NoteProperty -Name $hkey -Value $value -Force
			}
			else
			{
				$value = $referenceTime.$vkey
				$obj | Add-Member -MemberType NoteProperty -Name $hkey -Value $value -Force
			}
			
			
			
		}
		$progress.PerformStep()
		#"$($progress.Value) : $($progress.Maximum)" | Out-Host
	}
	
	$progress.Value = 0
}


#endregion


#region Function => DATA-DAILY-MONTHLY

function chd-data-Compute-Duration-measure
{
	param ($collection,
		$property,
		$progress = $null)
	
	return @($collection | ForEach-Object{ $_.$property.Ticks }) | Measure-Object -Maximum -Average
	
}


function chd-data-Compute-Duration-format
{
	param ($ticks,
		$format = "hh':'mm':'ss")
	
	return [timespan]::FromTicks($ticks).ToString($format)
	
}


function chd-data-Monthly-reports
{
	<#
		.SYNOPSIS
			
	
		.DESCRIPTION
			
	
		.PARAMETER  .........
			
	
		.PARAMETER  ..........
			
	
		.PARAMETER ...........
			
		
		.PARAMETER ...........
			
	
		.EXAMPLE
			
	#>	
	param ($inCollection,
		$outCollection,
		$property = "Count",
        $category,
		$year,
		$month,
		$daysLong,
		$progress = (New-Object System.Windows.Forms.ToolStripProgressBar))
	
	#$analyst = $inCollection | Group-Object -Property 'Created Via Description' | Select-Object -ExpandProperty Name
	
	$unique_id = $year.ToString() + $month.ToString() + $category.ToString() + $property.ToString()
	$unique_id = $unique_id.GetHashCode()
	$id = [convert]::ToString($unique_id, 16)
	
	$outCollectionIds = @($outCollection | Select-Object -ExpandProperty id)
	
	#region Computing
	
	if (! ($outCollectionIds.contains($id)))
	{	
		$progress.Maximum = $ranges.count
		$progress.Step = 1
		$progress.Value = 0
		
		### DAILY DAILY ranges
		$collection_horizontal = New-Object System.Collections.ArrayList
		
		$inCollection_grouped = $inCollection | Group-Object -Property 'Range'
		$ranges = $inCollection_grouped | Select-Object -ExpandProperty Name
		
		### DAILY DAILY ranges
		$ranges | ForEach-Object {
			
			$range = $_
			
			$recordObject = New-Object System.Management.Automation.PSObject
			$recordObject | Add-Member -MemberType NoteProperty -Name Ranges -Value $range
			
			# monthly data
			$hourly_row_data = $inCollection_grouped | Where-Object{ $_.Name -eq $range } | Select-Object -ExpandProperty Group
			$hourly_row_data_grouped = $hourly_row_data | Group-Object -Property DayLong
			
			$daysLong | ForEach-Object {
				
				$dayLong = $_
				
				$hourly_column_data = $hourly_row_data_grouped | Where-Object{ $_.Name -match $dayLong } | Select-Object -ExpandProperty Group
				#$hourly_column_data = $hourly_row_data | Where-Object{ $_.DayLong -match $dayLong }
				
				if (@($hourly_column_data).Count -eq 0)
				{
					#Read-Host
					$value = '-'
					$recordObject | Add-Member -MemberType NoteProperty -Name $dayLong -Value $value
				}
				else
				{
					$hourly_cell_stats = $hourly_column_data.Stats
					$value = $hourly_cell_stats.$property
					$recordObject | Add-Member -MemberType NoteProperty -Name $dayLong -Value $value
				}
				
			}
			
			$collection_horizontal.Add($recordObject) | Out-Null
			
			$progress.PerformStep()
		}
		
		$dataMonth = New-Object System.Management.Automation.PSObject
		$dataMonth | Add-Member -MemberType NoteProperty -Name id -Value $id
		$dataMonth | Add-Member -MemberType NoteProperty -Name Year -Value $year
		$dataMonth | Add-Member -MemberType NoteProperty -Name Month -Value $month
		$dataMonth | Add-Member -MemberType NoteProperty -Name Category -Value $category
		$dataMonth | Add-Member -MemberType NoteProperty -Name Property -Value $property
		$dataMonth | Add-Member -MemberType NoteProperty -Name Matrix -Value $collection_horizontal
		$outCollection.Add($dataMonth) | Out-Null
		
	}
	
	#endregion
	
	
	$progress.Value = 0
	return $id
}


function chd-data-Daily-Reports
{
	param ($inCollection,
		$outCollection,
        $category,
		$year,
		$month,
		$daysLong,
		$progress = (New-Object System.Windows.Forms.ToolStripProgressBar))
	
	
	#$outCollectionIds = @($outCollection | Group-Object $id | Select-Object -ExpandProperty Name | ForEach-Object { [System.Convert]::ToString($_ , 16) })
	$outCollectionIds = @($outCollection | Select-Object -ExpandProperty id)
	
	
	$daysLong | ForEach-Object {
		
		$dayLong = $_
		
		### 
		$unique_id = $year.ToString() + $month.ToString() + $dayLong.ToString() + $category.ToString()
		$unique_id = $unique_id.GetHashCode()
		$id = [convert]::ToString($unique_id, 16)
		
		if (! ($outCollectionIds.contains($id)))
		{
			
			$day_data = $null
			$day_data = $inCollection | Group-Object -Property 'DayLong' | Where-Object { $_.Name -match $dayLong } | Select-Object -ExpandProperty Group
			$dayOfYear = $day_data | Group-Object -Property DayOfYear | Select-Object -ExpandProperty Name #-Unique
			
			$i = 0
			foreach ($i in (0 .. 23)) #{ $i }
			{

				#"Computing Ranges`t" + (Get-Date).ToString("HH:mm:ss") | Out-Host
				$rangeName = "[$i - $($i + 1)]".ToString()
				$hourly_data = $day_data | Where-Object { $_.'hourOfDay' -eq $i }
				
				$valueCount = 0
				$valueCount = $(@($hourly_data) | Measure-Object).count
				
                
				####
				$recordObj = New-Object System.Management.Automation.PSObject
				$recordObj | Add-Member -MemberType NoteProperty -Name Count -Value $valueCount

				$i = $i + 1
				
				###ADD RECORDS
				$StatsObject = New-Object System.Management.Automation.PSObject
				$StatsObject | Add-Member -MemberType NoteProperty -Name id -Value $id
				$StatsObject | Add-Member -MemberType NoteProperty -Name Category -Value $category
				$StatsObject | Add-Member -MemberType NoteProperty -Name Year -Value $year
				$StatsObject | Add-Member -MemberType NoteProperty -Name Month -Value $month
				$StatsObject | Add-Member -MemberType NoteProperty -Name DayOfyear -Value $dayOfYear
				$StatsObject | Add-Member -MemberType NoteProperty -Name DayLong -Value $dayLong
				$StatsObject | Add-Member -MemberType NoteProperty -Name Range -Value $rangeName
				$StatsObject | Add-Member -MemberType NoteProperty -Name Stats -Value $recordObj
				$outCollection.Add($StatsObject) | Out-Null
				
			}
			
		}
		
		#$progress.PerformStep()
	}


	#$progress.Value = 0
}


#endregion


