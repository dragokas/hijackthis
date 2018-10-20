
$packageName = 'hijackthis'
$installerType = 'exe'
$url = 'https://github.com/dragokas/hijackthis/raw/devel/binary/HiJackThis.exe'
$silentArgs = '/accepteula /install /autostart'
$validExitCodes = @(0)

Install-ChocolateyPackage -PackageName "$packageName" `
                          -FileType "$installerType" `
                          -Url "$url" `
                          -SilentArgs "$silentArgs" `
                          -ValidExitCodes $validExitCodes
