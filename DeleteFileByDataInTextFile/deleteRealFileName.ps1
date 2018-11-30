Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    Multiselect = $false # Multiple files can be chosen
    Filter = 'Excel File (*.xlsx, *.xls)|*.xls;*.xlsx' # Specified file types
    # TopMost = $true
}


 
[void]$FileBrowser.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true}))
# $result = $FileBrowser.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true}))

$file = $FileBrowser.FileName;

If($FileBrowser.FileNames -like "*\*") {

	# Do something 
    # $FileBrowser.FileName #Lists selected files (optional)
    $excel = New-Object -COM "Excel.Application"
    $excel.Visible = $false
    $workbook = $excel.Workbooks.Open($FileBrowser.FileName)
    $workbook.sheets.item(1).activate()
    $WorkbookTotal=$workbook.Worksheets.item(1)

    $pathFile = $value = $WorkbookTotal.Cells.Item(1, 1).Text

    $i = 2;
    $count = 0;
    while($true){
        $value = $WorkbookTotal.Cells.Item($i, 1)
        # $value.Text #this should give you back the Value in that Cell
        # Write-Host ($value.Text -ne '')
        if($value.Text){
            $i++
            $count++;
            # Write-Host "$($pathFile)\$($value.Text).zip"        
            try{
                Remove-Item -path "$($pathFile)\$($value.Text).zip" -Force
                Write-Host "$($pathFile)\$($value.Text).zip was deleted"
            }catch{
                Write-Host "Can't delete file $($pathFile)\$($value.Text).zip"
            }
            
        }else{
            break;
        }
    }
    Write-Host "Number of files was deleted : $($count)"
    # $value = $WorkbookTotal.Cells.Item(1, 1)
    # $value.Text #this should give you back the Value in that Cell

    #close application
    $workbook.close()
    $excel.Quit()
	
}

else {
    Write-Host "Cancelled by user"
}
Read-Host -Prompt "Press enter to exit."