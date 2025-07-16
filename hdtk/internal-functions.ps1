function Copy-Ticket {
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
}

function Display-VerticalMenu {
  <#
    .SYNOPSIS
      Displays a menu with a list options displayed in a vertical manner and facilitates the navigation of this menu and returns the index of the option selected.

    .PARAMETER Options
      Represents the list of options to appear in the menu.

    .PARAMETER Index
      Represents the index of the option to be highlighted when the menu is loaded.

    .PARAMETER MenuStart
      Represents the row position of the cursor within the buffer when it first prints the menu and where it will reprint the menu.
  #>

  param (
    [Parameter(Mandatory=$true)]
    [string[]]$Options,

    [int]$Index = 0,

    [int]$MenuStart = [Console]::CursorTop
  )

  :menu while ($true) {
    # Draws actions menu.
    [Console]::SetCursorPosition(0, $MenuStart)
    for ($i = 0; $i -lt $Options.Length; $i += 1) {
      if ($i -eq $Index) {
        Write-Host "[$($Options[$i])]" `
            -NoNewLine `
            -ForegroundColor ([Console]::BackgroundColor).ToString() `
            -BackgroundColor ([Console]::ForegroundColor).ToString()
      } else {
        Write-Host " $($Options[$i]) " -NoNewLine
      }
    }

    # Moves console cursor back to the start of the menu.
    [Console]::CursorTop -= [Math]::Floor((($Options.ForEach({ " $_ " }) -join '').Length) / [Console]::WindowWidth)
    [Console]::CursorLeft = 0

    # Handles keyboard input.
    $key = [Console]::ReadKey($true)
    switch ($key.Key) {
      'Escape' {
        [Console]::CursorTop += [Math]::Floor((($Options.ForEach({ " $_ " }) -join '').Length) / [Console]::WindowWidth)
        [Console]::CursorVisible = $true
        break menu
      } 'Enter' {
        [Console]::CursorTop += [Math]::Floor((($Options.ForEach({ " $_ " }) -join '').Length) / [Console]::WindowWidth)
        [Console]::CursorVisible = $true
        return $Index
      } 'LeftArrow' {
        if ($Index -gt 0) {
          $Index -= 1
        } else {
          $Index = $Options.Length - 1
        }
        continue menu
      } 'RightArrow' {
        if ($Index -lt ($Options.Length - 1)) {
          $Index += 1
        } else {
          $Index = 0
        }
        continue menu
      }
    }
  }
}

function Insert-Variables {
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
}

function Search-Objects {
  <#
    .SYNOPSIS
      Allows you to search for objects in Active Directory that match the filters given.

    .PARAMETER Filters
      Represents an array of dictionaries (hashtables), each containing two key-value pairs. (1) The first is called "Arguments" and its values represent the values to be searched for according to the values of the corresponding "Properties" key. (2) The second one is called "Properties" and its values represent the properties to match the arguments ("Arguments") against.

    .PARAMETER MaxResults
      Represents the maximum number of results to pull from Active Directory. Argument is passed to the ResultSetSize parameter of Get-ADUser / Get-ADGroup / Get-ADComputer.

    .PARAMETER Properties
      Represents the list of properties needed from the object. Arguments are passed to the Properties parameter of Get-ADUser / Get-ADGroup / Get-ADComputer.

    .PARAMETER SearchBase
      Represents the LDAP distinguished name to search under.

    .PARAMETER Type
      Represents whether this search is for users, groups or computers.

    .PARAMETER Literal
      Signifies that the arguments for every filter should be treated as literals, not wildcard patterns.

    .EXAMPLE
      Search-Objects.ps1 -Filters $filters -Type user -Properties 'Name', 'DisplayName'

      $filters = @(
        @{
          Arguments = @('arg1', 'arg2')
          Properties = @('Name', 'DisplayName')
        },
        @{
          Arguments = @('arg3', 'arg4')
          Properties = @('Title', 'Description')
        }
      )

      Returns all users in Active Directory whose name or display name matches either "arg1" or "arg2" while also having a title or description matching either "arg3" or "arg4".

      Filter string buiilt for Get-ADUser / Get-ADGroup / Get-ADComputer:

      ((Name -like "arg1") -or (Name -like "arg2") -or (DisplayName -like "arg1") -or (DisplayName -like "arg2")) -and ((Title -like "arg3") -or (Title -like "arg4") -or (Description -like "arg3") -or (Description -like "arg4"))
  #>

  param (
    [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
    [hashtable[]]$Filters,

    [ValidateRange(1, 100)]
    [int]$MaxResults = 20,

    [Parameter(Mandatory=$true, Position=1)]
    [ValidateSet('user', 'group', 'computer')]
    [string]$Type,

    [Parameter(Position=2)]
    [string[]]$Properties = '*',

    [string]$SearchBase,

    [switch]$Literal
  )

  # Builds filter string.
  $filterStrings = @()
  foreach ($filter in $Filters) {
    $conditions = @()
    if ($Literal) {
      $filter['Arguments'] = $filter['Arguments'] `
          | ForEach-Object { [wildcardpattern]::Escape($_) }
    }
    foreach ($property in $filter['Properties']) {
      if ($property -eq 'EmployeeID') {
        $SearchBase = 'OU=SMHEmployees,OU=SMHUsers,DC=smh,DC=com'
      }
      foreach ($argument in $filter['Arguments']) {
        $conditions += '({0} -like "{1}")' -f $property, $argument
      }
    }
    $filterStrings += '({0})' -f ($conditions -join ' -or ')
  }
  $filterString = '{0}' -f ($filterStrings -join ' -and ')

  # Returns object(s) matching built filter string.
  $GetAdArguments = @{
    Filter = $filterString
    ResultSetSize = $MaxResults
    Properties = $Properties
  }
  if ($SearchBase) {
    $GetAdArguments['SearchBase'] = $SearchBase
  }
  switch ($Type) {
    'user' {
      return Get-ADUser @GetAdArguments
    } 'group' {
      return Get-ADGroup @GetAdArguments
    } 'computer' {
      return Get-ADComputer @GetAdArguments
    }
  }
}

function Select-ObjectFromTable {
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
  [Console]::CursorTop = [Console]::CursorTop - 2
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
}
