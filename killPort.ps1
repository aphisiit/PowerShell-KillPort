$answer = "Y"
Do{
    do{
        $portNumber = Read-Host "Enter port number for find PID and kill "
    }While(!($portNumber))

    $Host.UI.RawUI.WindowTitle="Kill Port $portNumber Script"
    $Host.UI.RawUI.CursorSize=14
    $portSearch = netstat -ano | findstr :$portNumber
    if(!($portSearch.length -eq 0 -or $portSearch.length -le 0)){
        $portArray = $portSearch | ConvertFrom-String
        Write-Host "$portNumber use PID : "$portArray[1].P6
        Write-Host "Killing PID : "$portArray[1].P6
        if(!($portArray[1].P6)){
            Write-Host "Can't kill port. Because No PID"
        }elseif(taskkill /PID $portArray[1].P6 /F){
            Write-Host "Kill port sucess."
        }else{
            Write-Host "Can't kill port."
        }
    }else{
        Write-Host "Can't find PID with port $portNumber"
    }

    $answer = Read-Host "Prees any key and enter to continue or Prees 'n' or 'N' to exit script. "

    if($answer -ne "N" -or $answer -ne "n"){
        Write-Host "------------------------------------------------------------"
    }

}While($answer -ne "N" -or $answer -ne "n")