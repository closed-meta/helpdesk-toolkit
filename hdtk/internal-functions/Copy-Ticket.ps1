# [TODO] Add manual.

param (
  [string]$Subject,

  [string]$Body,

  [string]$Fulfillment
)

if ($Subject) {
  $quiet = Read-Host 'Press <enter> to copy the ticket subject'
  Set-Clipboard $Subject
}
if ($Body) {
  $quiet = Read-Host 'Press <enter> to copy the ticket body'
  Set-Clipboard $Body
}
if ($Fulfillment) {
  $quiet = Read-Host 'Press <enter> to copy the ticket fulfillment comment'
  Set-Clipboard $Fulfillment
}
