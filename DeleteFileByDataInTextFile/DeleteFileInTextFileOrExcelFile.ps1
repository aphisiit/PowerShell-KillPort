Add-Type -AssemblyName System.Windows.Forms,PresentationCore,PresentationFramework
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    Multiselect = $false # Multiple files can be chosen
    Filter = 'Text File (*.txt)|*.txt; | Excel File (*.xlsx, *.xls)|*.xls;*.xlsx;' # Specified file types
    # TopMost = $true
}
 
[void]$FileBrowser.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true}))
# $result = $FileBrowser.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true}))

$file = $FileBrowser.FileName;
# Write-Host ""



If($FileBrowser.FileNames -like "*\*") {

    # Do something 
    $pathFile = ""
    $countDeleteFile = 0;
    $countNotDeleteFile = 0;
    $MessageFileDelete = ""
    $MessageNotDelete = ""

    if([IO.Path]::GetExtension($file) -eq '.xlsx' -or [IO.Path]::GetExtension($file) -eq '.xls'){
        $Host.UI.RawUI.ForegroundColor="Blue"
        Write-Host "This is Excel File"
        $Host.UI.RawUI.ForegroundColor="Gray"

        $excel = New-Object -COM "Excel.Application"
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($FileBrowser.FileName)
        $workbook.sheets.item(1).activate()
        $WorkbookTotal=$workbook.Worksheets.item(1)

        $pathFile = $value = $WorkbookTotal.Cells.Item(1, 1).Text

        $i = 2;

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
    

    }elseif ([IO.Path]::GetExtension($file) -eq '.txt') {
        $Host.UI.RawUI.ForegroundColor="Blue"
        Write-Host "This is text file."
        $Host.UI.RawUI.ForegroundColor="Gray"

        foreach($line in Get-Content $FileBrowser.FileNames) {
            if($line -match $regex){
                if($pathFile -eq ""){
                    $pathFile = $line
                    Write-Host "path file : $($pathFile)"
                }else{
                    Write-Host "file name : $($line)"
                    if(Test-Path "$($pathFile)\$($line).zip"){
                        Remove-Item -path "$($pathFile)\$($line).zip" -Force
                        Write-Host "$($pathFile)\$($line).zip was deleted"
                        $countDeleteFile++
                    }else{
                        $countNotDeleteFile++
                        Write-Error "Can't delete file $($pathFile)\$($line).zip. Because file does not exit."
                    }
                }
                # Work here
                
            }
        }
    
        Write-Host "-----------------------------------------------------------"
        Write-Host "Number of files : $($count)"
        Write-Host "Number of files was deleted : $($countDeleteFile)"
        $Host.UI.RawUI.ForegroundColor="Red"
        Write-Host "Number of files was not delete : $($countNotDeleteFile)"
        $Host.UI.RawUI.ForegroundColor="Gray"

    }else{
        Write-Host "Unknows type file"
    }
	
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