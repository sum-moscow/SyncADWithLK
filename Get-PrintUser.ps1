$DetailedName = "Гунина Наталья Олеговна"
$Departament = "Дирекция Института открытого образования"
$OU = "Дирекция Института открытого образования, Институт открытого образования, Учебные подразделения"
$UserLogin = "no_gunina"
$Password = "qwerty123"

$TempDirName = [System.Guid]::NewGuid().ToString()
New-Item -Type Directory -Name $TempDirName -Path $env:temp

$TempDir = "$env:temp\$TempDirName"

$OriginalDoc =  "C:\SyncMyGuu\PrintUser.html"
$HTML = Get-Content -Encoding UTF8 $OriginalDoc 

$HTML = $HTML -replace "DetailedName",$DetailedName
$HTML = $HTML -replace "Departament",$Departament
$HTML = $HTML -replace "OU",$OU
$HTML = $HTML -replace "UserLogin",$UserLogin
$HTML = $HTML -replace "Password",$Password

$UserDoc = "$TempDir\$UserLogin.html"
$HTML > $UserDoc

$ie = new-object -com "InternetExplorer.Application"
Start-Sleep -second 10
$ie.Navigate("$UserDoc")
Start-Sleep -second 10
#while ( $ie.busy ) { Start-Sleep -second 3 }
$ie.ExecWB(6,2)
Start-Sleep -second 10
#while ( $ie.busy ) { Start-Sleep -second 3 }
$ie.quit()