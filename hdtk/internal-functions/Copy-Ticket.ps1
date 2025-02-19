<#
  .SYNOPSIS
    Helps the user quickly copy the ticket subject, body and fulfillment comment to their clipboard one at a time.

  .PARAMETER Subject
    Represents the ticket subject to be copied to the clipboard.

  .PARAMETER Body
    Represents the ticket body to be copied to the clipboard.

  .PARAMETER Fulfillment
    Represents the ticket fulfillment comment to be copied to the clipboard.
#>

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
