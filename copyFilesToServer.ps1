##copyFilesToServer.ps1
##author: Brandon Colley
##date: 1-23-13
##Maintains Excel Spreadsheet of server list and applications. Writes status and date performed
##Extendable script. Add new test variable and if statement. Add two new headers\columns manually to spreadsheet
Import-Module ActiveDirectory
#create instance of Excel
$myExcel = New-Object -ComObject "Excel.Application"
$myWorkbook = $myExcel.Workbooks.Open("C:\bginfo\serverList.xlsx")
$myWorksheet = $myWorkbook.ActiveSheet
$myCells = $myWorksheet.Cells
$row=2 #skip headers
$serverName = $myCells.item($row,1).Text #get cell contents

#while data exists in the first column
While ($serverName -ne "")
{
	#build tests to determine server status
	$validServer = Test-Path \\$serverName\c$
	$hasBGinfo = Test-Path \\$serverName\c$\bginfo\myBGinfo.bgi # simply change this and copy data below for new version
	
	#not a valid servername
	if(!$validServer){
		$myCells.item($row,2) = "Invalid Name" #write to cell
	}
	#does not already have BGinfo
	elseif(!$hasBGinfo){
		#check for OS - startup folder is different (supports 2003 and 2008)
		$serverADobj = Get-ADComputer $serverName -Properties OperatingSystem
		$serverOS = $serverADobj.OperatingSystem
		#copy data
		if($serverOS -eq "Windows Server 2003"){
			copy c:\bginfo\callBGinfo.bat "\\$serverName\c$\Documents and Settings\All Users\Start Menu\Programs\Startup"
		}
		else{
			copy c:\bginfo\callBGinfo.bat "\\$serverName\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
		}	
		mkdir "\\$serverName\c$\bginfo" | Out-Null
		copy c:\bginfo\bginfo.exe "\\$serverName\c$\bginfo"
		copy c:\bginfo\myBGinfo.bgi "\\$serverName\c$\bginfo"
		$myCells.item($row,2) = "Completed" #write to cell
		$myCells.item($row,3) = (Get-Date) #write to cell
	}
	#look at the next row and serverName
	$row++
	$serverName = $myCells.item($row,1).Text
}
#save, close, and exit Excel
$myWorkbook.Save()
$myWorkbook.Close()
$myExcel.Quit()