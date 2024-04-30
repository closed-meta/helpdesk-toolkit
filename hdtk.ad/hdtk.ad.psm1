# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

# scripting guidelines #
<#
  - "Why have scripting conventions?":
    <https://www.oracle.com/java/technologies/cc-java-programming-language.html>.
  - 80-character column limit.
  - 2-space indentation with 4-space indentation for continuation lines.
  - Use 1TB-style bracing.
      - <https://en.wikipedia.org/wiki/Indentation_style#One_True_Brace>.
      - Resembles K&R style bracing with a few differences:
          - Never omit braces (even for single-statement blocks).
          - Place opening-braces on the same line, rather than on their
            own line.
          - Use chaining (a practice where something is placed on the same
            line as the closing brace, unless it is the first statement in the
            block) for the following blocks: (1) if-else statements;
            (2) switch statements; (3) exception handling statements.

            ```
            if (conditon) {
              statement
              statement
            } elseif (condition) {
              statement
              statement
            } else {
              statement
              statement
            }
            switch (condition) {
              case {
                statement
                statement
              } case {
                statement
                statement
              } default {
                statement
                statement
              }
            }
            try {
              statement
              statement
            } catch {
              statement
              statement
            } finally {
              statement
              statement
            }
            ```

  - Identifier naming conventions:
      - whispercase for modules.
      - PascalCase for classes, methods and properties.
      - camelCase for variables (aside from properties).
      - Do not use prefixes (such as `s_`, `_`, `I`, et cetera).
      - Avoid abbreviations.
  - Don't use double-quotes for strings unless there's a special character 
    or you need to do variable substitution. Use single-quotes instead.
      - Special characters: 
        <https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_special_characters>.
  - Don't use quotes when accessing the properties of an onject unless you 
    expect the object to have a property containing a special character.
      - Special characters: 
        <https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_special_characters>.
  - Commenting:
      - Comments may begin with tags within brackets such as: `[BUG]`, 
        `[FIXME]`, `[HACK]`, `[TODO]`.
  - Do not quote strings in examples for documentation unless the string 
    contains any of the following characters:
      - ```<space> ' " ` , ; ( ) { } | & < > @ # -```
  - Do not use continuation lines in the comment-based manuals/help.
  - Writing parameters with "param" blocks:
      - Use a blank line as the delimiter between parameter definitions.
      - The attributes (minus the parameter's type) of each parameter should 
        preceed the parameter name, each starting on its own line.
      - Parameter attributes should be written in alphabetical order.
        Their arguments (if the attribute has any) should also be written
        in alphabetical order.
      - The parameter type should immediately preceed the parameter name with
        a space separating the type and the name.
      - Switch parameters should be included at the end of the block.
#>

$ILLEGAL_GROUPS = @()

function Get-User {
  <#
    .SYNOPSIS
      Allows you to search Active Directory for users using various parameters such as Username, Employee ID, Name, Phone and Email. At least one argument of every parameter used must match with a user for the user to be considered a match.

      By default, after the user's property list has been printed, you are then presented with a variety of actions that can be performed, such as unlocking the account, resetting the password, reloading the list, et cetera. This can be disabled using the DisableActions switch.

      ALIAS: guser

    .PARAMETER Username
      Represents the username(s) to search Active Directory users for. If used, at least one of the arguments passed must match a user's username for the user to be considered a match.

      ALIAS: user

    .PARAMETER EmployeeId
      Represents the employee ID(s) to search Active Directory users for. If used, at least one of the arguments passed must match a user's employee ID for the user to be considered a match.

      ALIAS: eid

    .PARAMETER Name
      Represents the name(s) to search Active Directory users for. If used, at least one of the arguments passed must match a user's name for the user to be considered a match.

    .PARAMETER Phone
      Represents the phone number(s) to search Active Directory users for. If used, at least one of the arguments passed must match a user's phone number for the user to be considered a match.

    .PARAMETER Email
      Represents the email address(es) to search Active Directory users for. If used, at least one of the arguments passed must match a user's email address for the user to be considered a match.

    .PARAMETER Properties
      Represents an array of dictionaries (hashtables), each containing two key-value pairs. (1) The first is called "Title" and its value represents the string that will be used as the displayed name of the property in the property list. (2) The second one is called "CanonName" and its value represents the canonical name of the property in Active Directory.

    .PARAMETER DisableActions
      Instructs the function to end after displaying the properties of the user instead of providing the user with follow-up actions.

      ALIAS: noactions

    .PARAMETER Literal
      Signifies that the arguments for every filter should be treated as literals, not wildcard patterns.

    .EXAMPLE
      Get-User RL097898

      Displays a list of properties associated with the user with the username "RL097898".

    .EXAMPLE
      guser RL097898

      Displays a list of properties associated with the user with the username "RL097898".

    .EXAMPLE
      Get-User -EmployeeId 12*45

      Retrieves all users in Active Directory whose employee ID matches the wildcard pattern "12*45" (such as those under the employee ID "12345", "12445", et cetera) and displays a list of properties assosicated with the retrieved user you select from the table of matches.

    .EXAMPLE
      Get-User -Name 'j* doe'

      Retrieves all users in Active Directory whose name matches the wildcard pattern "j* doe" (such as those with the employee ID "12345", "12445", et cetera) and displays a list of properties assosicated with the retrieved user you select from the table of matches.

    .EXAMPLE
      Get-User -Phone *123*456*7890*

      Retrieves all users in Active Directory whose phone number (either personal or work phone numbers) matches the wildcard pattern "*123*456*7890*" (such as those with the phone number "+1 123 456 7890", "+11234567890", "123/456-7890", "123-456-7890", et cetera) and displays a list of properties assosicated with the retrieved user you select from the table of matches.

    .EXAMPLE
      Get-User -Email 'john-doe@example.com'

      Retrieves all users in Active Directory whose email address (either personal or work email addresses) matches the wildcard pattern "john-doe@example.com" and displays a list of properties assosicated with the retrieved user you select from the table of matches.

    .EXAMPLE
      Get-User -Name 'Jo*n Doe' -Phone +1*123*456*, +1*123*789* -Email *example.com

      Retrieves all users in Active Directory whose name matches the wildcard pattern "Jo*n Doe" **and** whose phone number(s) match either "+1*123*456*" or "+1*123*789*" **and** whose email address matches "*example.com". Then displays a list of properties assosicated with the retrieved user you select from the table of matches.
  #>

  [Alias('guser')]
  [CmdletBinding()]

  param (
    [Alias('user')]
    [Parameter(Position=0, ValueFromPipeline=$true)]
    [SupportsWildcards()]
    [string[]] $Username,

    [Alias('eid')]
    [SupportsWildcards()]
    [string[]] $EmployeeId,

    [SupportsWildcards()]
    [string[]] $Name,

    [SupportsWildcards()]
    [string[]] $Phone,

    [SupportsWildcards()]
    [string[]] $Email,

    [hashtable[]] $Properties = @(
      @{ Title = 'Name (display)';  CanonName = 'DisplayName' },
      @{ Title = '       (legal)';  CanonName = 'Name' },
      @{ Title = 'Username';        CanonName = 'SamAccountName' },
      @{ Title = 'Employee ID';     CanonName = 'EmployeeID' },
      @{ Title = 'Email';           CanonName = 'EmailAddress' },
      @{ Title = 'Phone (desk)';    CanonName = 'ipPhone' },
      @{ Title = '  (personal)';    CanonName = 'otherMobile' },
      @{ Title = '  (personal)';    CanonName = 'telephoneNumber' },
      @{ Title = '  (personal)';    CanonName = 'otherHomePhone' },
      @{ Title = 'Department';      CanonName = 'Department' },
      @{ Title = 'Job title';       CanonName = 'Title' },
      @{ Title = 'Manager';         CanonName = 'Manager' },
      @{ Title = 'Description';     CanonName = 'Description' },
      @{ Title = 'Expires';         CanonName = 'accountExpires' },
      @{ Title = 'User enabled';    CanonName = 'Enabled' },
      @{ Title = 'User locked';     CanonName = 'LockedOut' },
      @{ Title = 'Lockout date';    CanonName = 'AccountLockoutTime' },
      @{ Title = 'Password set';    CanonName = 'PasswordLastSet' },
      @{ Title = 'Last bad logon';  CanonName = 'LastBadPasswordAttempt' },
      @{ Title = 'Last good logon'; CanonName = 'LastLogonDate' },
      @{ Title = 'User created';    CanonName = 'Created' },
      @{ Title = 'User modified';   CanonName = 'Modified' },
      @{ Title = 'Home directory';  CanonName = 'HomeDirectory' },
      @{ Title = 'X profile';       CanonName = 'ProfilePath' }
    ),

    [Alias('noactions')]
    [switch] $DisableActions,

    [switch] $Literal
  )

  $searchFilters = @()
  $selectProperties = @(
    @{ Header = '#' },
    @{
      Header = 'USERNAME'
      CanonName = 'SamAccountName'
    },
    @{
      Header = 'DISPLAY NAME'
      CanonName = 'DisplayName'
    }
  )

  if ($Username) {
    $searchFilters += @{
      Arguments = $Username
      Properties = @('SamAccountName')
    }
  }
  if ($EmployeeId) {
    $searchFilters += @{
      Arguments = $EmployeeId
      Properties = @('EmployeeID')
    }
    $selectProperties += @{
      Header = 'EMPLOYEE ID'
      CanonName = 'EmployeeID'
    }
  }
  if ($Name) {
    $searchFilters += @{
      Arguments = $Name
      Properties = @('Name', 'DisplayName')
    }
  }
  if ($Phone) {
    $searchFilters += @{
      Arguments = $Phone
      Properties = @('ipPhone', 'otherMobile', 'telephoneNumber', 'otherHomePhone')
    }
    $selectProperties += @{
      Header = 'DESK PHONE'
      CanonName = 'ipPhone'
    }
    $selectProperties += @{
      Header = 'PERSONAL PHONE'
      CanonName = 'otherMobile'
    }
  }
  if ($Email) {
    $searchFilters += @{
      Arguments = $Email
      Properties = @('EmailAddress')
    }
    $selectProperties += @{
      Header = 'EMAIL ADDRESS'
      CanonName = 'EmailAddress'
    }
  }

  $searchArguments = @{
    Filters = $searchFilters
    Type = 'user'
    Properties = $Properties.ForEach({ $_['CanonName'] })
    Literal = $Literal
  }

  $domainObjects = Search-Objects @searchArguments
  if ($domainObjects) {
    $user = Select-ObjectFromTable `
        -Objects $domainObjects `
        -Properties $selectProperties
  } else {
    Write-Error 'No users found.'
    return
  }

  # Prepares the "actions" menu.
  <# Prepares the menu here to reduce lag between when the "information"
     section is printed and when the "actions" menu is printed. #>
  $actions = @(
    'End',
    'Reload',
    'Reset password',
    'List groups',
    'Add groups'
  )
  if (($user.LockedOut) -or ($user.AccountLockoutTime)) {
    $actions += 'Unlock Account'
  }
  if ($domainObjects.Count -ne 1) {
    $actions += 'Return to search'
  }
  if ($user.Manager) {
    $actions += 'Search manager'
  }
  if ($user.EmailAddress) {
    $actions += 'Send email'
  }

  # Determines max property display name length for later padding.
  $maxLength = 0
  foreach ($property in $Properties) {
    if ($property['Title'].Length -gt $maxLength) {
      $maxLength = $property['Title'].Length
    }
  }

  # Writes properties for user.
  $timeSince = '{0}D : {1}H : {2}M ago'
  Write-Host "`n# INFORMATION #"
  foreach ($property in $Properties) {
    $canonName = $property['CanonName']
    $displayName = $property['Title'].PadRight($maxLength)
    $value = $user.$canonName
    if ($value -is `
        [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
      $value = $value.Value
    }  # [HACK]
    switch ($canonName) {
      'accountExpires' {
        if ($value) {
          if (($value -eq 0) -or ($value -gt [DateTime]::MaxValue.Ticks)) {
            continue
          } else {
            $value = ([DateTime]($user.accountExpires)).AddYears(1600).ToLocalTime()
          }
          $diff = (Get-Date) - $value
          if ($value -lt (Get-Date)) {
            Write-Host ("$displayName : $value ($timeSince)" `
                -f $diff.Days, $diff.Hours, $diff.Minutes) `
                -ForegroundColor 'red'
          } else {
            Write-Host "$displayName : $value" -ForegroundColor 'green'
          }
        }
      } 'AccountLockoutTime' {
        if ($value) {
          $diff = (Get-Date) - $value
          Write-Host ("$displayName : $value ($timeSince)" `
              -f $diff.Days, $diff.Hours, $diff.Minutes) `
              -ForegroundColor 'red'
        }
      } 'Created' {
        if ($value) {
          $diff = (Get-Date) - $value
          Write-Host ("$displayName : $value ($timeSince)" `
              -f $diff.Days, $diff.Hours, $diff.Minutes)
        }
      } 'EmployeeID' {
        if ($value) {
          Write-Host "$displayName : $value" -ForegroundColor 'green'
        } else {
          Write-Host "$displayName : $value" -ForegroundColor 'red'
        }
      } 'Enabled' {
        if ($value) {
          Write-Host "$displayName : $value" -ForegroundColor 'green'
        } else {
          Write-Host "$displayName : $value" -ForegroundColor 'red'
        }
      } 'LastBadPasswordAttempt' {
        if ($value) {
          $diff = (Get-Date) - $value
          Write-Host ("$displayName : $value ($timeSince)" `
              -f $diff.Days, $diff.Hours, $diff.Minutes)
        }
      } 'LastLogonDate' {
        if ($value) {
          $diff = (Get-Date) - $value
          Write-Host ("$displayName : $value ($timeSince)" `
              -f $diff.Days, $diff.Hours, $diff.Minutes)
        }
      } 'LockedOut' {
        if ($value) {
          Write-Host "$displayName : $value" -ForegroundColor 'red'
        } else {
          Write-Host "$displayName : $value" -ForegroundColor 'green'
        }
      } 'Manager' {
        if ($value) {
          $value = ($value -split ',')[0].Substring(3)
          Write-Host "$displayName : $value"
        } else {
          Write-Host "$displayName : "
        }
      } 'Modified' {
        $diff = (Get-Date) - $value
        Write-Host ("$displayName : $value ($timeSince)" `
            -f $diff.Days, $diff.Hours, $diff.Minutes)
      } 'PasswordLastSet' {
        if ($value) {
          $diff = (Get-Date) - $value
          if ($diff.Days -ge 90) {
            Write-Host ("$displayName : $value ($timeSince)" `
                -f $diff.Days, $diff.Hours, $diff.Minutes) `
                -ForegroundColor 'red'
          } else {
            Write-Host ("$displayName : $value ($timeSince)" `
                -f $diff.Days, $diff.Hours, $diff.Minutes) `
                -ForegroundColor 'green'
          }
        } else {
          Write-Host "$displayName : change at next sign-in" `
              -ForegroundColor 'yellow'
        }
      } 'ProfilePath' {
        if ($value) {
          Write-Host "$displayName : $value"
        }
      } 'otherHomePhone' {
        if ("$value".Trim()) {
          Write-Host "$displayName : $value" -ForegroundColor 'green'
        }
      } 'otherMobile' {
        if ("$value".Trim()) {
          Write-Host "$displayName : $value" -ForegroundColor 'green'
        }
      } 'telephoneNumber' {
        if ("$value".Trim()) {
          Write-Host "$displayName : $value" -ForegroundColor 'green'
        }
      } default {
        Write-Host "$displayName : $value"
      }
    }
  }
  Write-Host ''

  if ($DisableActions) {
    return
  }

  # Prints "actions" menu.
  Write-Host '# ACTIONS #'
  $i = 1
  foreach ($action in $actions) {
    Write-Host "[$i] $action  " -NoNewLine
    $i += 1
  }
  Write-Host ''

  Write-Host ''
  try {
    $selection = ([int] (Read-Host 'Action')) - 1
    if ($selection -lt 0) {
      Write-Error ('Invalid selection. Expected a number 1-' `
          + $actions.Count + '.')
      return
    }
    $selection = $actions[$selection]
  } catch {
    Write-Host ''
    Write-Error ('Invalid selection. Expected a number 1-' `
        + $actions.Count + '.')
    return
  }

  # Executes selected action.
  switch ($selection) {
    'End' {
      Write-Host ''
      return
    } 'Reload' {
      Get-User -Username $user.SamAccountName
    } 'Return to search' {
      $user = Select-ObjectFromTable `
          -Objects (Search-Objects @searchArguments) `
          -Properties $selectProperties
      Get-User $user.SamAccountName
    } 'Unlock account' {
      Unlock-ADAccount -Identity $user.SamAccountName
      Get-User -Username $user.SamAccountName
    } 'Reset password' {
      Write-Host ''
      Reset-Password -Users $user.SamAccountName
      Unlock-ADAccount -Identity $user.SamAccountName
      Write-Host ''
      Get-User -Username $user.SamAccountName
    } 'List groups' {
      Write-Host ''
      Get-ADPrincipalGroupMembership -Identity $user.SamAccountName `
          | ForEach-Object { $_.Name } `
          | Sort-Object `
          | Write-Host
      Write-Host ''
    } 'Add groups' {
      Write-Host ''
      Write-Host ('***You may add multiple groups by separating them with ' `
          + "a comma (no space).***`n")
      $groups = (Read-Host 'Groups to add') -split ','
      foreach ($group in $groups) {
        if ($ILLEGAL_GROUPS -contains $group) {
          Write-Error ("Skipping `"$group`". Modifying this group is " `
              + 'restricted by Data Security.')
          continue
        }
        Add-ADGroupMember -Identity $group -Members $user
      }
      Write-Host ''
    } 'Search manager' {
      Get-User -Name (($user.Manager) -split ',')[0].Substring(3)
    } 'Send email' {
      Start-Process "mailto:$($user.EmailAddress)"
    } default {
      Write-Host ''
      Write-Error 'Undefined action.'
      return
    }
  }
}

function Get-Group {
  <#
    .SYNOPSIS
      Allows you to search Active Directory for groups using various parameters such as Name and Description. At least one argument of every parameter used must match with a group for the group to be considered a match.

      ALIAS: ggroup

    .PARAMETER Name
      Represents the name(s) to search Active Directory groups for. If used, at least one of the arguments passed must match a group's name for the group to be considered a match.

    .PARAMETER Description
      Represents the description(s) to search Active Directory groups for. If used, at least one of the arguments passed must match a group's description for the group to be considered a match.

      ALIAS: desc

    .PARAMETER Properties
      Represents an array of dictionaries (hashtables), each containing two key-value pairs. (1) The first is called "Title" and its value represents the string that will be used as the displayed name of the property in the property list. (2) The second one is called "CanonName" and its value represents the canonical name of the property in Active Directory.

    .PARAMETER DisableActions
      Instructs the function to end after displaying the properties of the user instead of providing the user with follow-up actions.

    .PARAMETER Literal
      Signifies that the arguments for every filter should be treated as literals, not wildcard patterns.

    .EXAMPLE
      Get-Group -Name GROUP_NAME

      Retrieves all groups in Active Directory whose name matches "GROUP_NAME" and displays a list of properties assosicated with the retrieved group you select from the table of matches.

    .EXAMPLE
      Get-Group -Name GROUP_*

      Retrieves all groups in Active Directory whose name begins with "GROUP_" and displays a list of properties assosicated with the retrieved group you select from the table of matches.

    .EXAMPLE
      Get-Group -Description '*email-group@example.com*'

      Retrieves all groups in Active Directory whose description contains "email-group@example.com" and displays a list of properties assosicated with the retrieved group you select from the table of matches.

    .EXAMPLE
      Get-Group -Description *\\company\department\unit*

      Retrieves all groups in Active Directory whose description contains "\\company\department\unit" and displays a list of properties assosicated with the retrieved group you select from the table of matches.
  #>

  [Alias('ggroup')]
  [CmdletBinding()]

  param (
    [Parameter(Position=0, ValueFromPipeline=$true)]
    [SupportsWildcards()]
    [string[]] $Name,

    [Alias('desc')]
    [SupportsWildcards()]
    [string[]] $Description,

    [SupportsWildcards()]
    [string[]] $Email,

    [hashtable[]] $Properties = @(
      @{ Title = 'Name';        CanonName = 'Name' },
      @{ Title = '    ';        CanonName = 'SamAccountName' },
      @{ Title = 'Description'; CanonName = 'Description' },
      @{ Title = 'Notes';       CanonName = 'info' },
      @{ Title = 'Email';       CanonName = 'mail' },
      @{ Title = 'Managed by';  CanonName = 'ManagedBy' },
      @{ Title = 'Created';     CanonName = 'Created' },
      @{ Title = 'Modified';    CanonName = 'Modified' }
    ),

    [Alias('noactions')]
    [switch] $DisableActions,

    [switch] $Literal
  )

  if (-not ($Name -or $Description -or $Email)) {
    Write-Error ('At least one search parameter (that is: Name, ' `
        + 'Description, Email) must be used.')
    return
  }

  $searchFilters = @()
  $selectProperties = @(
    @{ Header = '#' },
    @{
      Header = 'NAME'
      CanonName = 'Name'
    },
    @{
      Header = 'DESCRIPTION'
      CanonName = 'Description'
    }
  )
  if ($Name) {
    $searchFilters += @{
      Arguments = $Name
      Properties = @('Name', 'DisplayName', 'SamAccountName')
    }
  }
  if ($Description) {
    $searchFilters += @{
      Arguments = $Description
      Properties = @('Description')
    }
  }
  if ($Email) {
    $searchFilters += @{
      Arguments = $Email
      Properties = @('mail')
    }
    $selectProperties += @{
      Header = 'EMAIL ADDRESS'
      CanonName = 'mail'
    }
  }

  $searchArguments = @{
    Filters = $searchFilters
    Type = 'group'
    Properties = $Properties.ForEach({ $_['CanonName'] })
    Literal = $Literal
  }

  $domainObjects = Search-Objects @searchArguments
   if (-not $domainObjects) {
    Write-Error 'No groups found.'
    return
   }
  $group = Select-ObjectFromTable `
      -Objects $domainObjects `
      -Properties $selectProperties

  # Prepares the "actions" menu.
  <# Prepares the menu here to reduce lag between when the "information"
     section is printed and when the "actions" menu is printed. #>
  $actions = @(
    'End',
    'Return to search'
  )
  if ($ILLEGAL_GROUPS -notcontains $group.Name) {
    $actions += 'Add users'
  }
  if ($group.ManagedBy) {
    $actions += 'Search manager'
  }

  # Determines max property display name length for later padding.
  $maxLength = 0
  foreach ($property in $Properties) {
    if ($property['Title'].Length -gt $maxLength) {
      $maxLength = $property['Title'].Length
    }
  }

  # Writes properties for user.
  Write-Host ''
  Write-Host '# INFORMATION #'
  foreach ($property in $Properties) {
    $canonName = $property['CanonName']
    $displayName = $property['Title'].PadRight($maxLength)
    $value = $group.$canonName
    if ($value -is `
        [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
      $value = $value.Value
    }  # [HACK]
    switch ($canonName) {
      'info' {
        if ($value) {
          Write-Host "$displayName : $value" -ForegroundColor 'red'
        }
      } 'mail' {
        if ($value) {
          Write-Host "$displayName : $value"
        }
      } 'ManagedBy' {
        if ($value) {
          $value = ($value -split ',')[0].Substring(3)
          Write-Host "$displayName : $value"
        } else {
          Write-Host "$displayName : "
        }
      } default {
          Write-Host "$displayName : $value"
      }
    }
  }
  Write-Host ''

  if ($DisableActions) {
    return
  }

  # Prints "actions" menu.
  Write-Host '# ACTIONS #'
  $i = 1
  foreach ($action in $actions) {
    Write-Host "[$i] $action  " -NoNewLine
    $i += 1
  }
  Write-Host ''

  # Requests selection from the user.
  Write-Host ''
  $selection = $actions[([int] (Read-Host 'Action')) - 1]

  # Executes selection.
  Write-Host ''
  switch ($selection) {
    'End' {
      return
    } 'Return to search' {
      $group = Select-ObjectFromTable `
          -Objects (Search-Objects @searchArguments) `
          -Properties $selectProperties
      Get-Group $group.SamAccountName
    } 'Add users' {
      Write-Host ('You may add multiple users by separating them with a ' `
          + "comma (no space).`n")
      $users = (Read-Host 'Users to add') -split ','
      Add-ADGroupMember -Identity $group.Name -Members $users
      Write-Host ''
    } 'Search manager' {
      Get-User `
          -Name (($group.ManagedBy) -split ',')[0].Substring(3)
    } default {
      Write-Error 'Invalid selection. Expected a number 1-' `
          + $actions.Count + '.'
      return
    }
  }
}

function Reformat-Names {
  <#
    .SYNOPSIS
      Takes a list of names written in the "last name, first name" format and returns a table of the referenced users with each row containg, the status (disabled account, name not found, et cetera), a number representing the index when there is a list of accounts found under the given name, the name searched (if the LooseSearch or LooserSearch parameter is used), the account's  display name and the account's username.

      ALIAS: reformat

    .PARAMETER List
      Represents the closing you would like to be used in the reply (such as "Best", "Sincerely", "Thank you", et cetera). Use the string "none" to have the closing part excluded.

    .PARAMETER EntryDelimiter
      Represents the name of the customer you are building the reply for.

      ALIAS: entrybreak

    .PARAMETER NameDelimiter
      Represents the greeting you would like to be used in the reply (such as "Hi", "Hello", "Hey", et cetera). Use the string "none" to have the greeting part excluded.

      ALIAS: namebreak

    .PARAMETER LooseSearch
      Signifies that an account will be counted as a match as long as the display name begins with the first name and ends with the last name. This means that an account with a display name like "John Smith Doe" will be counted as a match for the name "john doe".

      Using this switch will cause a minor dip in performance.

      ALIAS: loose

    .PARAMETER LooserSearch
      Signifies that an account will be counted as a match as long as the display name contains both the first name and the last name somewhere after the first name in the display name. This means that an account with a display name like "Doctor John Smith Doe II" will be counted as a match for the name "john doe".

      Using this switch will cause a more significant dip in performance and is not generally recommended unless it is expected that many names in the list will not be found unless searched for in this way.

      ALIAS: looser

    .EXAMPLE
      Reformat-Names /home/username/Documents/names.txt

      Retrieves a list of names written in the "<last name>, <first name>" format, with each name on a separate line, from a file located at "/home/username/Documents/names.txt". Then creates a table with columns labled "status", "#", "name" and "username", where each row represents an account associated with a name from the file.

      A number will appear in the "#" column for a row if it is amoung a list of accounts found for a name.

      The display names of the accounts must match "<first name> <last name>" exactly.

    .EXAMPLE
      Reformat-Names /home/username/Documents/names.txt -LooseSearch

      Retrieves a list of names written in the "<last name>, <first name>" format, with each name on a separate line, from a file located at "/home/username/Documents/names.txt". Then creates a table with columns labled "status", "#", "name" and "username", where each row represents an account associated with a name from the file.

      A number will appear in the "#" column for a row if it is amoung a list of accounts found for a name.

      The display names of the accounts must begin with the first name and end with the last name. This means that an account with a display name like "John Smith Doe" will be counted as a match for the name "john doe".

    .EXAMPLE
      Reformat-Names /home/username/Documents/names.txt -LooserSearch

      Retrieves a list of names written in the "<last name>, <first name>" format, with each name on a separate line, from a file located at "/home/username/Documents/names.txt". Then creates a table with columns labled "status", "#", "name" and "username", where each row represents an account associated with a name from the file.

      A number will appear in the "#" column for a row if it is amoung a list of accounts found for a name.

      The accounts must contain the first name somewhere in the display name, as well as the last name somewhere after the first name in the display name. This means that an account with a display name like "Doctor John Smith Doe II" will be counted as a match for the name "john doe".
  #>

  [Alias('reformat')]
  [CmdletBinding()]

  param (
    [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
    [string] $List,

    [Alias('entrybreak')]
    [string] $EntryDelimiter = "`n",

    [Alias('namebreak')]
    [string] $NameDelimiter = ', ',

    [Alias('loose')]
    [switch] $LooseSearch,

    [Alias('looser')]
    [switch] $LooserSearch
  )

  $entries = @()
  if (Test-Path -Path $List -IsValid) {
    $entries = Get-Content -Delimiter $EntryDelimiter -Path $Path
  } else {
    $entries = $List -split $EntryDelimiter
  }

  $table = New-Object System.Data.DataTable
  $quiet = $table.Columns.Add('STATUS')
  $quiet = $table.Columns.Add('#')
  $quiet = $table.Columns.Add('SEARCH')
  $quiet = $table.Columns.Add('NAME')
  $quiet = $table.Columns.Add('USERNAME')
  foreach ($entry in $entries) {
    $names = $entry -split $NameDelimiter
    $filter = ''
    if ($LooseSearch) {
      $filter = "(DisplayName -like ""$($names[1])*$($names[0])"") " `
            + "-or (Name -like ""$($names[1])*$($names[0])"")"
    } elseif ($LooserSearch) {
      $filter = "(DisplayName -like ""*$($names[1])*$($names[0])*"") " `
            + "-or (Name -like ""*$($names[1])*$($names[0])*"")"
    } else {
      $filter = "(DisplayName -like ""$($names[1]) $($names[0])"") " `
            + "-or (Name -like ""$($names[1]) $($names[0])"")"
    }
    $results = Get-ADUser -Filter $filter `
        -Properties 'DisplayName', 'Enabled', 'Name', 'SamAccountName'
    if (-not $results) {
      $row = $table.NewRow()
      $row.'STATUS' = 'NOT FOUND'
      $row.'#' = ' '
      $row.'SEARCH' = "$($names[1]) $($names[0])"
      $row.'NAME' = ' '
      $row.'USERNAME' = ' '
      $table.Rows.Add($row)
    }
    $i = 1
    foreach ($user in $results) {
      $row = $table.NewRow()
      if ($user.Enabled) {
        $row.'STATUS' = ' '
      } else {
        $row.'STATUS' = 'DISABLED'
      }
      if ($results.Count -gt 1) {
        $row.'#' = $i
      } else {
        $row.'#' = ' '
      }
      $row.'SEARCH' = "$($names[1]) $($names[0])"
      $row.'NAME' = $user.Name
      $row.'USERNAME' = $user.SamAccountName
      $table.Rows.Add($row)
      $i += 1
    }
  }
  $table | Format-Table
}

function Remote-Computer {
  <#
    .SYNOPSIS
      Allows you to search Active Directory for a computer and to then offer that computer remote assistance using the Quick Assist / Windows Remote Assistance / Microsoft Remote Assistance (msra.exe).

      This command is only available for computers running Windows.

      ALIAS: remote

    .PARAMETER Name
      Represents the name of the computer to search Active Directory for and offer remote assistance to.

    .PARAMETER Literal
      Signifies that the string passed to the Name parameter (the computer name) should be treated as a literal, not a wildcard pattern.

    .EXAMPLE
      Remote-Computer -Name COMPUTER_1

      Sends an offer to the computer ("COMPUTER_1") for remote assistance / access.

    .EXAMPLE
      Remote-Computer COMPUTER_1

      Sends an offer to the computer ("COMPUTER_1") for remote assistance / access.
  #>

  [Alias('remote')]
  [CmdletBinding()]

  param (
    [Alias('pc')]
    [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
    [SupportsWildcards()]
    [string] $Name,

    [switch] $Literal
  )

  <# The IsWindows automatic variable not yet supported by this version of 
     PowerShell. #>
  <#
    if (-not $IsWindows) {
      Write-Error 'Remote-Computer is only available on Windows.'
      return
    }
  #>

  $searchFilters += @{
    Arguments = $Name
    Properties = @('Name')
  }
  $selectProperties = @(
    @{ Header = '#' },
    @{
      Header = 'NAME'
      CanonName = 'Name'
    }
  )

  $searchArguments = @{
    Filters = $searchFilters
    Type = 'computer'
    Properties = 'Name'
    Literal = $Literal
  }

  $domainObjects = Search-Objects @searchArguments
  if (-not $domainObjects) {
    Write-Error 'No computers found.'
    return
  }
  $computer = Select-ObjectFromTable `
      -Objects $domainObjects `
      -Properties $selectProperties
  $Name = $computer.Name

  Set-Clipboard -Value $Name
  Write-Host ''
  Write-Host 'Copied...' -ForegroundColor 'green'
  Write-Host $Name

  Write-Host ''
  Write-Host "Connecting to `"$Name`"..."

  Start-Process -FilePath 'msra.exe' -ArgumentList "/offerra $Name"
  Write-Host ''
}

function Reset-Password {
  <#
    .SYNOPSIS
      Allows the user to force a password reset for one or more users in Active Directory.

      ALIAS: pwreset

    .PARAMETER Subjects
      Represents the users to force a password reset for.

    .PARAMETER Force
      Signifies that the function should not ask the user for confirmation before executing a password reset for multiple users.

    .EXAMPLE
      Reset-Password JD012345

      Changes the password for a user under the username "JD012345" and forces the user to change their password at next log-on.

    .EXAMPLE
      Reset-Password 'JS033549', 'FU033549'

      Changes the password for the users under the usernames "JS033549" and "FU033549" and forces the users to change their password at next log-on. Asks the user for confirmation before executing the password reset.
  #>

  [Alias('pwreset')]
  [CmdletBinding()]

  param (
    [Parameter(
      HelpMessage=('Enter the username(s) of the user(s) you would like to ' `
          + 'force a password reset for.'),
      Mandatory=$true,
      Position=0,
      ValueFromPipeline=$true
    )]
    [string] $Users,

    [switch] $Force
  )

  if (($Users.Count -gt 1) -and (-not $Force)) {
    Write-Host ('Are you sure that you want to force a password reset on ' `
        + $Users.Count + 'accounts?')
    $confirmation = Read-Host '[Y] Yes  [N] No'
    if (($confirmation -ine 'y') -and ($confirmation -ine 'yes')) {
      return
    }
  }

  # Resets password to one provided by the user.
  foreach ($user in $Users) {
    Set-ADAccountPassword `
        -Identity $user `
        -Reset `
        -NewPassword (ConvertTo-SecureString `
            -AsPlainText (Read-Host 'New password') `
            -Force)
    Set-ADUser -Identity $user -ChangePasswordAtLogon $true
  }
}

function Search-Objects {
  <#
    .SYNOPSIS
      Allows you to search Active Directory for objects that match the filters given.

    .PARAMETER Filters
      Represents an array of dictionaries (hashtables), each containing two key-value pairs. (1) The first is called "Arguments" and its values represent the values to be searched for according to the values of the corresponding "Properties" key. (2) The second one is called "Properties" and its values represent the properties to match the arguments ("Arguments") against.

    .PARAMETER MaxResults
      Represents the maximum number of results to pull from Active Directory. Argument is passed to the ResultSetSize parameter of Get-ADUser / Get-ADGroup / Get-ADComputer.

    .PARAMETER Properties
      Represents the list of properties needed from the object. Arguments are passed to the Properties parameter of Get-ADUser / Get-ADGroup / Get-ADComputer.

    .PARAMETER SearchBase
      Represents the Active Directory path to search under.

    .PARAMETER Type
      Represents whether this search is for users, groups or computers.

    .PARAMETER Literal
      Signifies that the arguments for every filter should be treated as literals, not wildcard patterns.

    .EXAMPLE
      Search-Objects -Filters $filters -Type user -Properties 'Name', 'DisplayName'

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
    [hashtable[]] $Filters,

    [ValidateRange(1, 100)]
    [int] $MaxResults = 20,

    [Parameter(Mandatory=$true, Position=1)]
    [ValidateSet('user', 'group', 'computer')]
    [string] $Type,

    [Parameter(Position=2)]
    [string[]] $Properties = '*',

    [string] $SearchBase,

    [switch] $Literal
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
      Select-ObjectFromTable -Objects $objs -Properties $props

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
    [object[]] $Objects,

    [Parameter(Mandatory=$true, Position=1)]
    [SupportsWildcards()]
    [hashtable[]] $Properties
  )

  if ($Objects.Count -eq 1) {
    return $Objects
  }

  # Populates table of results.
  $table = New-Object System.Data.DataTable
  foreach ($property in $Properties) {
    $quiet = $table.Columns.Add($property['Header'])
  }
  $i = 1
  foreach ($object in $Objects) {
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
    $i += 1
  }
  $i -= 1

  # Displays options to user and requests a selection from the results.
  $table | Format-Table | Out-String | Write-Host
  $selection = [int](Read-Host "Enter index # of your selection (1-$i)")
  return $Objects[$selection - 1]
}
