<#
  .SYNOPSIS
    Allows you to select an object from an array of objects using a table containing an index number and object metadata.

  .PARAMETER Objects
    Represents the array of objects to allow you to choose from. If only one argument is provided, that object is automatically returned without asking you for a selection.

  .PARAMETER Properties
    Represents an array of dictionaries (hashtables).

    Every dictionary has a key-value pair called "Header" whose value is a string that acts as the column headers for the table.

    All but one of the dictionaries have a key-value pair called "CanonName" whose value represents the object property to populate that column with. The one exception is the dictionary for the first column ("#") for the row index.

  .EXAMPLE
    Select-ObjectFromTable.ps1 -Objects $objs -Properties $props

    $objs = @(Get-ADUser 'JS022877', Get-ADUser 'OB024534')

    $props = @(
      @{
        Header = 'USERNAME'
        CanonName = 'SamAccountName'
      },
      @{
        Header = 'DISPLAY NAME'
        CanonName = 'DisplayName'
      }
    )

    Prints the following table:

    # USERNAME DISPLAY NAME
    - -------- ------------
    0 JS022877 James Smith
    1 OB024534 Olivia Brown

    Then asks for a selection and returns the corresponding object.
#>

param (
  [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
  [SupportsWildcards()]
  [object[]]$Objects,

  [Parameter(Mandatory=$true, Position=1)]
  [SupportsWildcards()]
  [hashtable[]]$Properties
)

if ($Objects.Count -eq 1) {
  return $Objects
}

# Populates table of results.
$table = New-Object System.Data.DataTable
foreach ($property in $Properties) {
  $quiet = $table.Columns.Add($property['Header'])
}
$i = 0
foreach ($object in $Objects) {
  $i += 1
  $row = $table.NewRow()
  foreach ($property in $Properties) {
    if ($property['Header'] -eq '#') {
      $row.'#' = $i
      continue
    }
    $value = $object.($property['CanonName'])
    if ($value -is `
        [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
      $value = $value.Value
    }  # [HACK]
    $row.($property['Header']) = "$value"
  }
  $table.Rows.Add($row)
}

# Displays options to user and requests a selection from the results.
Out-String -InputObject ($table | Format-Table -Wrap) | Write-Host
$selection = $null
$selection = Read-Host "[0] Return  [1-$i] Make selection"
if (-not $selection) {
  Write-Error "Invalid selection. Expected a number 0-$i."
  return
}
try {
  $selection = [int]$selection
} catch {
  Write-Error "Invalid selection. Expected a number 0-$i."
  return
}
if ($selection -eq 0) {
  return
} elseif ($selection -gt $i) {
  Write-Error "Invalid selection. Expected a number 0-$i."
  return
}
return $Objects[$selection - 1]
