param(
  [int]$Limit = 40,
  [string]$Mailbox = ""
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Convert-SafeText {
  param($Value)
  if ($null -eq $Value) { return "" }
  return ([string]$Value -replace "[\x00-\x08\x0B\x0C\x0E-\x1F]", " ").Trim()
}

function Convert-TextPreview {
  param([string]$Text, [int]$Length)
  if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
  $clean = ((Convert-SafeText $Text) -replace "\s+", " ").Trim()
  if ($clean.Length -le $Length) { return $clean }
  return $clean.Substring(0, $Length).Trim() + "..."
}

function Get-SenderAddress {
  param($Item)
  try {
    if ($Item.SenderEmailType -eq "EX") {
      $exchangeUser = $Item.Sender.GetExchangeUser()
      if ($exchangeUser -and $exchangeUser.PrimarySmtpAddress) {
        return $exchangeUser.PrimarySmtpAddress
      }
    }
  } catch {}
  return [string]$Item.SenderEmailAddress
}

function Test-TextContains {
  param([string]$Value, [string]$Needle)
  if ([string]::IsNullOrWhiteSpace($Value) -or [string]::IsNullOrWhiteSpace($Needle)) { return $false }
  return $Value.IndexOf($Needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
}

function Get-VisibleMailboxes {
  param($Namespace)
  $names = New-Object System.Collections.Generic.List[string]
  try {
    foreach ($account in $Namespace.Accounts) {
      $parts = @(
        (Convert-SafeText $account.DisplayName),
        (Convert-SafeText $account.UserName),
        (Convert-SafeText $account.SmtpAddress)
      ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
      if ($parts.Count -gt 0) {
        $names.Add("Account: " + (($parts | Select-Object -Unique) -join " / "))
      }
    }
  } catch {}

  try {
    foreach ($store in $Namespace.Stores) {
      $names.Add("Store: " + (Convert-SafeText $store.DisplayName))
    }
  } catch {}

  return $names
}

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$inbox = $null

if (-not [string]::IsNullOrWhiteSpace($Mailbox)) {
  foreach ($account in $namespace.Accounts) {
    $displayName = Convert-SafeText $account.DisplayName
    $userName = Convert-SafeText $account.UserName
    $smtpAddress = Convert-SafeText $account.SmtpAddress

    if (
      (Test-TextContains $smtpAddress $Mailbox) -or
      (Test-TextContains $userName $Mailbox) -or
      (Test-TextContains $displayName $Mailbox)
    ) {
      try {
        $inbox = $account.DeliveryStore.GetDefaultFolder(6)
        break
      } catch {}
    }
  }

  $recipient = $namespace.CreateRecipient($Mailbox)
  $null = $recipient.Resolve()
  if (-not $inbox -and $recipient.Resolved) {
    try {
      $inbox = $namespace.GetSharedDefaultFolder($recipient, 6)
    } catch {}
  }

  if (-not $inbox) {
    foreach ($store in $namespace.Stores) {
      $root = $store.GetRootFolder()
      $storeName = Convert-SafeText $store.DisplayName
      $rootName = Convert-SafeText $root.Name
      if ((Test-TextContains $storeName $Mailbox) -or (Test-TextContains $rootName $Mailbox)) {
        try {
          $inbox = $store.GetDefaultFolder(6)
          break
        } catch {}
      }
    }
  }
}

if (-not $inbox -and -not [string]::IsNullOrWhiteSpace($Mailbox)) {
  $visible = Get-VisibleMailboxes $namespace
  $visibleText = if ($visible.Count -gt 0) { $visible -join "; " } else { "no accounts or stores visible to Outlook automation" }
  throw "Could not find or access Outlook mailbox '$Mailbox'. Outlook automation can currently see: $visibleText. If your school account is only in New Outlook, add it to classic Outlook Desktop as well."
}

if (-not $inbox) {
  $inbox = $namespace.GetDefaultFolder(6)
}

$items = $inbox.Items
$items.Sort("[ReceivedTime]", $true)

$messages = New-Object System.Collections.Generic.List[object]
$index = 1

while ($messages.Count -lt $Limit -and $index -le $items.Count) {
  $item = $items.Item($index)
  $index += 1

  if (-not $item -or -not $item.EntryID -or -not $item.Subject) {
    continue
  }

  $importance = switch ([int]$item.Importance) {
    0 { "low" }
    2 { "high" }
    default { "normal" }
  }

  $flagStatus = if ([int]$item.FlagStatus -gt 0) { "flagged" } else { "notFlagged" }
  $body = Convert-SafeText $item.Body

  $messages.Add([pscustomobject]@{
    id = [string]$item.EntryID
    subject = (Convert-SafeText $item.Subject)
    from = [pscustomobject]@{
      emailAddress = [pscustomobject]@{
        name = (Convert-SafeText $item.SenderName)
        address = (Convert-SafeText (Get-SenderAddress $item))
      }
    }
    receivedDateTime = ([datetime]$item.ReceivedTime).ToUniversalTime().ToString("o")
    bodyPreview = Convert-TextPreview $body 360
    bodyText = Convert-TextPreview $body 6000
    importance = $importance
    isRead = -not [bool]$item.UnRead
    flag = [pscustomobject]@{
      flagStatus = $flagStatus
    }
    conversationId = [string]$item.ConversationID
    webLink = ""
    categories = @($item.Categories)
  })
}

[pscustomobject]@{
  value = $messages
} | ConvertTo-Json -Depth 8
