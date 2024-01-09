
# Install the Exchange Online PowerShell module #
# Ref: https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps #
Install-Module -Name ExchangeOnlineManagement -RequiredVersion 3.4.0

# Sets the PowerShell execution policies for Windows computers #
# Ref: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7.4 #
Set-ExecutionPolicy RemoteSigned

# Configure WinRM #
# Ref: https://learn.microsoft.com/en-us/windows/win32/winrm/installation-and-configuration-for-windows-remote-management #
winrm quickconfig
winrm set winrm/config/client/auth '@{Basic="false"}'

# Configure PowerShellGet #
# Ref: https://learn.microsoft.com/en-us/powershell/gallery/powershellget/overview?view=powershellget-3.x #
Install-Module PowerShellGet -Force -AllowClobber
Install-Module Microsoft.PowerShell.PSResourceGet -Repository PSGallery
Install-PackageProvider -Name NuGet -Force
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

# Load the Exchange Online PowerShell module #
Import-Module ExchangeOnlineManagement