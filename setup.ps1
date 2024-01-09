Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.4.0
Set-ExecutionPolicy RemoteSigned
winrm quickconfig
winrm set winrm/config/client/auth '@{Basic="false"}'
Install-Module PowerShellGet -Force -AllowClobber
Install-Module Microsoft.PowerShell.PSResourceGet -Repository PSGallery
Install-PackageProvider -Name NuGet -Force
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
Import-Module ExchangeOnlineManagement