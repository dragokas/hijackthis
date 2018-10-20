
$packageName = 'hijackthis'
$installerType = 'exe'
$silentArgs = '/silentuninstall'
$validExitCodes = @(0)

$is32 = (-not (Test-Path 'env:PROCESSOR_ARCHITEW6432') -and ($env:PROCESSOR_ARCHITECTURE -eq 'x86'))
$PF32 = if ($is32) { "$env:ProgramFiles" } else { "${env:ProgramFiles(x86)}" }
$file = "$PF32\HiJackThis Fork\HiJackThis.exe"

Uninstall-ChocolateyPackage -PackageName "$packageName" `
                            -FileType "$installerType" `
                            -SilentArgs "$silentArgs" `
                            -File "$file" `
                            -ValidExitCodes $validExitCodes
