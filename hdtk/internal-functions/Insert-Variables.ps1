<#
  .SYNOPSIS
    Performs variable substitution on a string (with variables being written in the "$name" format used with PowerShell).

  .PARAMETER String
    Represents the string that variable substitution will be performed on. Placeholders for variables appear in this initial string using the "$name" format also used with PowerShell.

  .EXAMPLE
    Insert-Variables.ps1 'They live at $address.'

    In this example, it will be assumed that there is a PowerShell variable called "$address" representing the string "308 Negra Arroyo Lane, Albuquerque, New Mexico 87104".

    This command would return the string "They live at 308 Negra Arroyo Lane, Albuquerque, New Mexico 87104".

  .EXAMPLE
    Insert-Variables.ps1 'Please contact $person.'

    In this example, it will be assumed that there is a PowerShell variable called "$person" representing the string "$name ($email)"; another called "$name" representing the string "John Doe"; and a third called "$email" representing the string "john-doe@example.com".

    This command would return the string "Please contact John Doe (john-doe@example.com).".
#>

param (
  [Parameter(
    Mandatory=$true,
    Position=0,
    ValueFromPipeline=$true
  )]
  [string]$String
)

$extension = '.txt'
$copyPasteDirectory = 'copypastes'
$copyPastePath = [System.IO.Path]::Combine(
  [System.IO.Directory]::GetParent($PSScriptRoot),
  $copyPasteDirectory
)
$names = ([System.IO.Path]::Combine($copyPastePath, "*$extension") `
    | Get-ChildItem -Name).ForEach({
  $_.SubString(0, ($_.Length - $extension.Length))
})
$paths = ([System.IO.Path]::Combine($copyPastePath, "*$extension") `
    | Get-ChildItem -Name).ForEach({
      [System.IO.Path]::Combine($copyPastePath, $_)
    })
$contents = $paths.ForEach({
  (Get-Content $_ -Encoding 'utf8') -join "`n"
})
$numFiles = $names.Count

for ($i = 0; $i -lt $numFiles; $i += 1) {
  $String = $String -replace "\`$$($names[$i])", $contents[$i]
}

return $String
