Add-Type -AssemblyName System.Windows.Forms,PresentationCore,PresentationFramework
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    Multiselect = $false # Multiple files can be chosen
    Filter = 'Excel File (*.xlsx, *.xls)|*.xls;*.xlsx;' # Specified file types
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
    $countDeleteFile = 0;
    $countNotDeleteFile = 0;
    $MessageFileDelete = ""
    $MessageNotDelete = ""
    while($true){
        $value = $WorkbookTotal.Cells.Item($i, 1)
        # $value.Text #this should give you back the Value in that Cell
        # Write-Host ($value.Text -ne '')
        if($value.Text){
            $i++
            $count++
            # Write-Host "$($pathFile)\$($value.Text).zip"        
            if(Test-Path "$($pathFile)\$($value.Text).zip"){
                Remove-Item -path "$($pathFile)\$($value.Text).zip" -Force
                Write-Host "$($pathFile)\$($value.Text).zip was deleted"
                $countDeleteFile++
            }else{
                $countNotDeleteFile++
                Write-Error "Can't delete file $($pathFile)\$($value.Text).zip. Because file does not exit."
            }
            # try{                
            #     Remove-Item -path "$($pathFile)\$($value.Text).zip" -Force
            #     Write-Host "$($pathFile)\$($value.Text).zip was deleted"
            #     $countDeleteFile++

                # $MessageFileDelete = "$($MessageFileDelete) `n $($pathFile)\$($value.Text).zip was deleted"
            # }catch {
                # $countNotDeleteFile++
                # Write-Error "Can't delete file $($pathFile)\$($value.Text).zip"

                # MessageNotDelete = "$($MessageNotDelete) `n Can't delete file $($pathFile)\$($value.Text).zip"
            # }
            
        }else{
            break;
        }
    }
    Write-Host "-----------------------------------------------------------"
    Write-Host "Number of files : $($count)"
    Write-Host "Number of files was deleted : $($countDeleteFile)"
    $Host.UI.RawUI.ForegroundColor="Red"
    Write-Host "Number of files was not delete : $($countNotDeleteFile)"
    $Host.UI.RawUI.ForegroundColor="Gray"
    # $value = $WorkbookTotal.Cells.Item(1, 1)

    # Show Dialog Message Box
    # $ButtonType = [System.Windows.MessageBoxButton]::OK
    # $MessageIcon = [System.Windows.MessageBoxImage]::None
    # $MessageBody = "File delete `n $($MessageFileDelete) `n File not delete $($MessageNotDelete)"
    # $MessageTitle = "Result Execusion"
    # $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)

    #close application
    $workbook.close()
    $excel.Quit()
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
	
}

else {

    # Show Dialog Message Box
    # $ButtonType = [System.Windows.MessageBoxButton]::OK
    # $MessageIcon = [System.Windows.MessageBoxImage]::None
    # $MessageBody = "Cancelled by user"
    # $MessageTitle = "Result Execusion"
    # $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)

    Write-Host "Cancelled by user"
}
Read-Host -Prompt "Press enter to exit."