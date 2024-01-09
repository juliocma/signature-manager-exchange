# Author: Julio Cesar M. Amaral #
# Contact: julio@outlook.com.br #

# Before using the script, make sure to run setup.ps1 #

# TEMPLATE HTML #
$fileHtml = "./template/signature.html"

# TEMPLATE TXT #
$fileTxt = "./template/signature.txt"

$count = 0
Write-Host "Signature Generator for Exchange Online" -ForegroundColor Yellow

Write-Host "Author: " -NoNewline
Write-Host "Julio Cesar M. Amaral" -ForegroundColor Blue
Write-Host "Contact: " -NoNewline
Write-Host "julio@outlook.com.br" -ForegroundColor Blue

Write-Host "Before using the script, make sure to run setup.ps1" -ForegroundColor Red


# Start connecting to Exchange using Connect-ExchangeOnline #

try {
  if (!(Get-PSSession | Where-Object { $_.Name -match 'ExchangeOnline' -and $_.Availability -eq 'Available' })) {  
    $adminAccount = Read-Host -Prompt 'Enter administrator email'
    Connect-ExchangeOnline -UserPrincipalName $adminAccount
    Write-Host "Authentication successfully with the user " -NoNewline  -ForegroundColor Yellow
    Write-Host $adminAccount -ForegroundColor Yellow
  }

}
catch {
  Write-Host "Authentication error." -ForegroundColor Red
  Exit
}

# Loading HTML and TXT templates #
$signatureHTML = Get-Content -Path $fileHtml -ReadCount 0 | Out-String
$signatureTXT = Get-Content -Path $fileTxt -encoding utf8 -ReadCount 0 | Out-String

# Select a filter for email boxes, e.g. for everyone: *@email.com #
$mailBoxFilter = Read-Host -Prompt 'Enter the email to update, leave it blank for everyone.'

# Carregar a lista das caixas de email #
$listMailbox = Get-Mailbox -Filter "UserPrincipalName -like '$mailBoxFilter'" | Sort-Object  -Property UserPrincipalName

$totalMailBox = $listMailbox.count

if (totalMailBox -gt 0) {

  Write-Host "$totalMailBox" -ForegroundColor Cyan -NoNewline
  if ($count -gt 1) {
    Write-Host " email accounts were found" -NoNewlin
  }
  else {
    Write-Host " email account was found" -NoNewlin
  }
  Write-Host " for update, please wait..."

  # Inserting/Updating signatures #

  foreach ($mailBox in $listMailbox) {

    # Searching for variables in exchange and AD, reference: https://learn.microsoft.com/en-us/entra/identity/hybrid/connect/reference-connect-sync-attributes-synchronized #
    $userPrincipalName = $mailBox.UserPrincipalName 
    $displayName = $mailBox.DisplayName
    $mobilePhone = (Get-User $userPrincipalName).MobilePhone
    $title = (Get-User $userPrincipalName).Title 

    # Opening the file and changing the variables for the data #
    $signatureHTML = $signatureHTML.Replace("[DisplayName]", $displayName)
    $signatureHTML = $signatureHTML.Replace("[JobTitle]", $title)
    $signatureHTML = $signatureHTML.Replace("[EmailAddress]", $userPrincipalName)
    $signatureHTML = $signatureHTML.Replace("[MobileNumber]", $mobilePhone)

    $signatureTXT = $signatureTXT.Replace("[DisplayName]", $DisplayName)
    $signatureTXT = $signatureTXT.Replace("[JobTitle]", $JobTitle)
    $signatureTXT = $signatureTXT.Replace("[EmailAddress]", $userPrincipalName)


    # Applying the signature #
    Set-MailboxMessageConfiguration -Identity $userPrincipalName -AutoAddSignature $True -AutoAddSignatureOnReply $True -SignatureText $signatureTXT -signatureHTML $signatureHTML
  
    $count ++
    Write-Host "$count de $totalMailBox - " -ForegroundColor Cyan -NoNewline
    Write-Host "signature added to " -NoNewline
    Write-Host $userPrincipalName -ForegroundColor Green

  }

  Write-Host "Updated " -NoNewline
  Write-Host "$count of $totalMailBox" -ForegroundColor Cyan -NoNewline
  if ($count -gt 1) {
    Write-Host " email signatures"
  }
  else {
    Write-Host " email signature"
  }
}
else {
  Write-Host "No email accounts were found" -ForegroundColor Red
}
get-PSSession | remove-PSSession
