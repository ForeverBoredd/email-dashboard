$ErrorActionPreference = "Stop"

$inputJson = [Console]::In.ReadToEnd()
if ([string]::IsNullOrWhiteSpace($inputJson)) {
  throw "Missing draft payload"
}

$payload = $inputJson | ConvertFrom-Json
if (-not $payload.id) {
  throw "Missing Outlook message id"
}
if (-not $payload.content) {
  throw "Missing draft content"
}

$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$original = $namespace.GetItemFromID([string]$payload.id)
$reply = $original.Reply()

$reply.Body = ([string]$payload.content).Trim() + "`r`n`r`n" + $reply.Body
$reply.Save()

[pscustomobject]@{
  id = [string]$reply.EntryID
  subject = [string]$reply.Subject
  saved = $true
} | ConvertTo-Json -Depth 4
