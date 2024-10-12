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
