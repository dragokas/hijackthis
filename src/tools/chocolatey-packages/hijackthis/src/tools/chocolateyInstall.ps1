$ErrorActionPreference = 'Stop'

$packageName= 'hijackthis'
$toolsDir   = "$(Split-Path -parent $MyInvocation.MyCommand.Definition)"
$url        = 'https://dragokas.com/tools/HiJackThis.zip'
$setupName  = 'HiJackThis.exe'

$packageArgs = @{
  packageName   = $packageName
  unzipLocation = $toolsDir
  fileType      = 'EXE'
  url           = $url

  softwareName  = 'HiJackThis+'
  silentArgs    = '/accepteula /install /autostart'
  validExitCodes= @(0)
}

Install-ChocolateyZipPackage @packageArgs

$packageArgs.file = Join-Path -Path $toolsDir -ChildPath $setupName
Install-ChocolateyInstallPackage @packageArgs
