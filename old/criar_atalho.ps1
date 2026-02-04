$WshShell = New-Object -ComObject WScript.Shell
$Desktop = [Environment]::GetFolderPath('Desktop')
$Shortcut = $WshShell.CreateShortcut("$Desktop\Conversor HAP.lnk")
$Shortcut.TargetPath = "C:\Users\pedro\Downloads\Programas2\HAPPXXXX\Conversor_HAP.bat"
$Shortcut.WorkingDirectory = "C:\Users\pedro\Downloads\Programas2\HAPPXXXX"
$Shortcut.Description = "Conversor Excel para HAP 5.1"
$Shortcut.Save()
Write-Host "Atalho criado no Ambiente de Trabalho!"
