# [TODO] Add manual.

param (
  [string]$Subject,
  
  [string]$Body,
  
  [string]$Fullfillment
)

if ($Subject) {
  Read-Host 'Press <enter> to copy the ticket subject'
  Set-Clipboard $Subject
}
if ($Body) {
  Read-Host 'Press <enter> to copy the ticket body'
  Set-Clipboard $Body
}
if ($Fulfillment) {
  Read-Host 'Press <enter> to copy the ticket fulfillment comment'
  Set-Clipboard $Fulfillment
}
