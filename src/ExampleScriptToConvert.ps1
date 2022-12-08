$volume=Get-Volume -DriveLetter C
$route=get-adsroute -local;
$session=New-TcSession -Route $route -Port 851;
Write-TcValue -Session $session -Path 'Main.sizeRemaining' -value $volume.SizeRemaining -Force; 
Close-TcSession -InputObject $session;