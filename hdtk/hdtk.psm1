# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

# scripting guidelines #
<#
  - "Why have scripting conventions?":
    <https://www.oracle.com/java/technologies/cc-java-programming-language.html>.
  - 80-character column limit.
  - 2-space indentation with 4-space indentation for continuation lines.
  - Write in the [1TB style](https://en.wikipedia.org/wiki/Indentation_style#One_True_Brace).
  - Identifier naming conventions:
      - whispercase for modules.
      - PascalCase for classes, methods, and properties.
      - camelCase for variables (aside from properties).
      - Do not use prefixes (such as `s_`, `_`, `I`, et cetera).
      - Avoid abbreviations.
      - Use the plural form for a parameter identifier if it can accept multiple values, otherwise, use the singular form.
  - Don't use double-quotes for strings unless there's a special character 
    or you need to do variable substitution. Use single-quotes instead.
      - Special characters: 
        <https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_special_characters>.
  - Don't use quotes when accessing the properties of an onject unless you 
    expect the object to have a property containing a special character.
      - Special characters: 
        <https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_special_characters>.
  - Don't escape double-quotes with the typical escape character (`"). Escape 
    them with another double-quote instead ("").
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
      - The parameter type should immediately preceed the parameter name.
      - Switch parameters should be included at the end of the block.
#>

$ILLEGAL_GROUPS = @()

<# Generates a dictionary with an entry for each PowerShell 
   script in the "internal-functions" directory, deriving 
   each entry's key from the name of the corresponding 
   script, and the entry's value being the script's path. #>
$children = [System.IO.Path]::Combine(
  $PSScriptRoot,
  'internal-functions',
  '*.ps1'
) | Get-ChildItem
$internal = @{}
foreach ($child in $children) {
  $internal[$child.BaseName] = $child.FullName
}
Remove-Variable -Name 'children'

function Access-CopyPastes {
  <#
    .SYNOPSIS
      Opens the directory containing the copy-paste documents that the copy-paste variables are built from.
  #>

  $path = [System.IO.Path]::Combine($PSScriptRoot, 'copypastes')
  Invoke-Item $path
}

function Copy-AccountUnlockTicket {
  <#
    .SYNOPSIS
      Builds and copies the body, subject, and fulfillment comment of a ticket for an account unlock. Also unlocks the account if the account is a domain / Active Directory account.

      ALIAS: unlock

    .PARAMETER Type
      Represents the platform that this account is for.

    .PARAMETER Phone
      Represents the phone number that the customer called from to get the acount unlocked.

    .PARAMETER Username
      Represents the domain username for the account and only applies to domain account unlocks. This username is used to pull the lockout date and time for the ticket body and to unlock the account.

      ALIAS: user

    .PARAMETER DisableSubjectCopy
      Signifies that the function should not offer to copy the ticket subject after waiting for you to press <enter>.

      ALIAS: nosubject

    .PARAMETER DisableFulfillmentCopy
      Signifies that the function should not offer to copy the ticket fulfillment comment after waiting for you to press <enter>.

      ALIAS: nocomment

    .EXAMPLE
      Copy-AccountUnlockTicket -Type domain -Phone '+1 012 345 6789' -Username JD012345

      Copies, then prints the ticket body, subject, and fulfillment comment.

    .EXAMPLE
      Copy-AccountUnlockTicket -Type Imprivata

      Copies, then prints the ticket body, subject, and fulfillment comment.
  #>

  [Alias('unlock')]
  [CmdletBinding()]

  param (
    [Parameter(
      HelpMessage='Enter the platform that this account is on.',
      Mandatory=$true,
      Position=0
    )]
    [string]$Type,

    [string]$Phone,

    [Alias('user')]
    [string]$Username,

    [Alias('nosubject')]
    [switch]$DisableSubjectCopy,

    [Alias('nocomment')]
    [switch]$DisableFulfillmentCopy
  )

  $subject = "unlock account ($Type)"
  $body = "Caller requested that their account ($Type) be unlocked."
  $fulfillment = "Unlocked account ($Type)."

  if ($Phone) {
    if ($body) {
      $body += "`n`n"
    }
    $body += "Called from: $Phone"
  }

  switch ($Type) {
    'domain' {
      if ($Username) {
        $user = Get-ADUser $Username `
            -Properties 'AccountLockoutTime', 'LockedOut', 'SamAccountName'
        if ($user.LockedOut -or $user.AccountLockoutTime) {
          if ($user.AccountLockoutTime) {
            $date = ($user.AccountLockoutTime).ToString('yyyy-MM-dd')
            $time = ($user.AccountLockoutTime).ToString('HH:mm:ss')
            if ($body) {
              $body += "`n`n"
            }
            $body += "Account locked on $date at $time."
          }
          Unlock-ADAccount -Identity $user.SamAccountName
        }
      }
    }
  }

  $parameters = @{}
  if (-not $DisableSubjectCopy) {
    $parameters['Subject'] = $subject
  }
  $parameters['Body'] = $body
  if (-not $DisableFulfillmentCopy) {
    $parameters['Fulfillment'] = $fulfillment
  }
  & $internal['Copy-Ticket'] @parameters
}

function Copy-ConnectPrinterTicket {
  <#
    .SYNOPSIS
      Builds and copies the body of a ticket for connecting one or more computers to one or more printers.

      ALIAS: addprinter

    .PARAMETER Computers
      Represents the name(s) of the computer(s) being connected to the printer(s).

      ALIAS: pc

    .PARAMETER Printers
      Represents the path(s) of the printer(s) being connected to.

    .PARAMETER DisableSubjectCopy
      Signifies that the function should not offer to copy the ticket subject after waiting for you to press <enter>.

      ALIAS: nosubject

    .PARAMETER DisableFulfillmentCopy
      Signifies that the function should not offer to copy the ticket fulfillment comment after waiting for you to press <enter>.

      ALIAS: nocomment

    .EXAMPLE
      Copy-ConnectPrinterTicket COMPUTER_NAME PRINTER_NAME

      Copies, then prints the ticket body, subject, and fulfillment comment.

    .EXAMPLE
      Copy-ConnectPrinterTicket 'COMPUTER_0', 'COMPUTER_1' 'PRINTER_0', 'PRINTER_1'

      Copies, then prints the ticket body, subject, and fulfillment comment.

    .EXAMPLE
      Copy-ConnectPrinterTicket COMPUTER_NAME PRINTER_NAME -DisableSubjectCopy

      Copies, then prints the ticket body and fulfillment comment.

    .EXAMPLE
      Copy-ConnectPrinterTicket COMPUTER_NAME PRINTER_NAME -DisableFulfillmentCopy

      Copies, then prints the ticket body fulfillment comment.
  #>

  [Alias('addprinter')]
  [CmdletBinding()]

  param (
    [Alias('pc')]
    [Parameter(
      HelpMessage='Enter the name(s) of the computer(s) that you are connecting.',
      Mandatory=$true,
      Position=0
    )]
    [string[]]$Computers,

    [Parameter(
      HelpMessage='Enter the path(s) of the printer(s) that you are connecting.',
      Mandatory=$true,
      Position=1
    )]
    [string[]]$Printers,

    [Alias('nosubject')]
    [switch]$DisableSubjectCopy,

    [Alias('nocomment')]
    [switch]$DisableFulfillmentCopy
  )

  $listOfComputers = '"{0}"' -f ($Computers -join '", "')
  $listOfPrinters = '"{0}"' -f ($Printers -join '", "')

  $subject = 'connect computer(s) to printer(s)'
  $body = 'Customer requested to have '
  $fulfillment = 'Connected '

  if ($Computers.Count -gt 1) {
    $body += "some computers ($listOfComputers) and "
    $fulfillment += "the computers ($listOfComputers) to "
  } else {
    $body += "a computer ($listOfComputers) and "
    $fulfillment += "the computer ($listOfComputers) to "
  }
  if ($Printers.Count -gt 1) {
    $body += "some printers ($listOfPrinters) connected."
    $fulfillment += "the printers ($listOfPrinters)."
  } else {
    $body += "a printer ($listOfPrinters) connected."
    $fulfillment += "the printer ($listOfPrinters)."
  }
  $parameters = @{}
  if (-not $DisableSubjectCopy) {
    $parameters['Subject'] = $subject
  }
  $parameters['Body'] = $body
  if (-not $DisableFulfillmentCopy) {
    $parameters['Fulfillment'] = $fulfillment
  }
  & $internal['Copy-Ticket'] @parameters
}

function Copy-MapDriveTicket {
  <#
    .SYNOPSIS
      Builds and copies the body and subject of a ticket for mapping network drive(s).

      ALIAS: mapdrive

    .PARAMETER Computers
      Represents the name(s) of the computer(s) having drives mapped.

      ALIAS: pc

    .PARAMETER Paths
      Represents the path(s) of the network fileshare(s) being mapped for the computer(s).

    .PARAMETER DisableSubjectCopy
      Signifies that the function should not offer to copy the ticket subject after waiting for you to press <enter>.

      ALIAS: nosubject

    .PARAMETER DisableFulfillmentCopy
      Signifies that the function should not offer to copy the ticket fulfillment comment after waiting for you to press <enter>.

      ALIAS: nocomment

    .PARAMETER Remap
      Signifies that you are re-mapping a drive that was somehow un-mapped as opposed to mapping the path to a drive for the first time on that computer.

    .EXAMPLE
      Copy-MapDriveTicket -Computers COMPUTER_1 -Paths PATH_1

      Copies, then prints the ticket body, subject, and fulfillment comment.

    .EXAMPLE
      Copy-MapDriveTicket 'COMPUTER_1', 'COMPUTER_2' 'PATH_1', 'PATH_2'

      Copies, then prints the ticket body, subject, and fulfillment comment.

    .EXAMPLE
      Copy-MapDriveTicket COMPUTER_1 PATH_1 -DisableSubjectCopy

      Copies, then prints the ticket body and fulfillment comment.
  #>

  [Alias('mapdrive')]
  [CmdletBinding()]

  param (
    [Alias('pc')]
    [Parameter(
      HelpMessage='Enter the computer name(s).',
      Mandatory=$true,
      Position=0
    )]
    [string[]]$Computers,

    [Parameter(
      HelpMessage='Enter the network path(s).',
      Mandatory=$true,
      Position=1
    )]
    [string[]]$Paths,

    [Alias('nosubject')]
    [switch]$DisableSubjectCopy,

    [Alias('nocomment')]
    [switch]$DisableFulfillmentCopy,

    [switch]$Remap
  )

  $listOfComputers = '"{0}"' -f ($Computers -join '", "')
  $listOfPaths = '"{0}"' -f ($Paths -join '", "')

  $subject = ''
  $body = ''
  $fulfillment = ''

  if ($Remap) {
    $subject = 're-map drive(s)'
  } else {
    $subject = 'map drive(s)'
  }

  if ($Paths.Count -gt 1) {
    $body = "Customer requested to have drives ($listOfPaths) "
    if ($Remap) {
      $fulfillment = "Re-mapped the paths ($listOfPaths) to some drives for the "
    } else {
      $fulfillment = "Mapped the paths ($listOfPaths) to some drives for the "
    }
  } else {
    $body = "Customer requested to have a drive ($listOfPaths) "
    if ($Remap) {
      $fulfillment = "Re-mapped the path ($listOfPaths) to a drive for the "
    } else {
      $fulfillment = "Mapped the path ($listOfPaths) to a drive for the "
    }
  }
  if ($Computers.Count -gt 1) {
    if ($Remap) {
      $body += "re-mapped for some computers ($listOfComputers)."
    } else {
      $body += "mapped for some computers ($listOfComputers)."
    }
    $fulfillment += "computers ($listOfComputers)."
  } else {
    if ($Remap) {
      $body += "re-mapped for a computer ($listOfComputers)."
    } else {
      $body += "mapped for a computer ($listOfComputers)."
    }
    $fulfillment += "computer ($listOfComputers)."
  }

  Set-Clipboard -Value $body
  Write-Host ''
  Write-Host 'Copied...'
  Write-Host $body -ForegroundColor 'green'
  Write-Host ''

  if (-not $DisableSubjectCopy) {
    $parameters['Subject'] = $subject
  }
  $parameters['Body'] = $body
  if (-not $DisableFulfillmentCopy) {
    $parameters['Fulfillment'] = $fulfillment
  }
  & $internal['Copy-Ticket'] @parameters
}

function Format-Quote {
  <#
    .SYNOPSIS
      Prefixes every line of the provided string with a Markdown blockquote indicator ("> "). The level of blockquote nesting for the text may be specified with the Level parameter; and, to add blockquote syntax to a copied (Outlook-style) email containing a block of headers and a new line before the actual text, you may use the Email parameter.

      ALIAS: quote

    .PARAMETER Text
      Represents either the input text itself or the path to a text file containing the input string. If an invalid path is provided, the string will be treated as the input text itself.

    .PARAMETER Level
      Represents the level of nesting for the blockquote. For example, one level of nesting would look like...

      > Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.
      > 
      > Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.

      ... and two levels of nesting would look like...

      > > Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.
      > > 
      > > Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.

      ... and so on.

    .PARAMETER Email
      Represents that the input text should not include the block of email headers as part of the blockquote, and that the blank line separating the block of headers of the blockquote should be removed.

      Assumes that the text is the in the Outlook-style format (see below).

      > From: sender@example.com
      > Sent: Thursday, January 1st, 1970 12:01 AM
      > To: recipient@example.com
      > Subject: Lorem Ipsum
      > 
      > Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.

    .EXAMPLE
      Format-Quote "Paragraph 1`n`nParagraph 2"

      Would return the below:

      > Paragraph 1
      > 
      > Paragraph 2

    .EXAMPLE
      Format-Quote

      Would return the below:

      > From: sender@example.com
      > Sent: Thursday, January 1st, 1970 12:01 AM
      > To: recipient@example.com
      > Subject: Lorem Ipsum
      > 
      > Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.
      > 
      > Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.

    .EXAMPLE
      Format-Quote -Email

      Would return the below:

      From: sender@example.com
      Sent: Thursday, January 1st, 1970 12:01 AM
      To: recipient@example.com
      Subject: Lorem Ipsum
      > Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.
      > 
      > Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.
  #>

  [Alias('quote')]

  param (

    [Parameter(
      Position=0,
      ValueFromPipeline=$true
    )]
    [string]$Text = [System.IO.Path]::Combine($HOME, 'Desktop', 'text.txt'),

    [int]$Level = 1,

    [switch]$Email
  )

  $lines = @()
  if (Test-Path -Path $Text) {
    $lines = Get-Content -Path $Text -Encoding 'utf8'
  } elseif (Test-Path -Path $Text -IsValid) {
    Write-Error "Unable to locate a file at the address provided to the Text " `
        + "parameter (""$Text""). "
    $lines = $Text -split "`n"
  } else {
    $lines = $Text -split "`n"
  }
  $Text = ''
  $prefix = '> ' * $Level
  if ($Email) {
    $inHeaders = $true
    $notInBody = $true
    foreach ($line in $lines) {
      if ($inHeaders -and ($line -notmatch '[^ ]:')) {
        $inHeaders = $false
      }
      if ($notInBody -and (-not $inHeaders) -and $line.Trim()) {
        $notInBody = $false
      }
      if ($inHeaders) {
        if ($Text) {
          $Text = $Text, $line -join "`n"
        } else {
          $Text = $Text + $line
        }
      } elseif (-not $notInBody) {
        $Text = $Text, "$prefix$line" -join "`n"
      }
    }
  } else {
    foreach ($line in $lines) {
      $Text = $Text, "$prefix$line" -join "`n"
    }
  }

  return $Text
}

function Get-Computer {
  <#
    .SYNOPSIS
      Allows you to search for computers in Active Directory by Name. At least one argument must match the value of the corresponding property of a computer for it to be considered a match.

      ALIAS: gcom

    .PARAMETER Names
      Represents the name(s) to search computers in Active Directory for. If used, at least one of the arguments passed must match a computer's name for the computer to be considered a match.

    .PARAMETER Properties
      Represents an array of dictionaries (hashtables), each containing two key-value pairs. (1) The first is called "Title" and its value represents the string that will be used as the displayed name of the property in the property list. (2) The second one is called "CanonName" and its value represents the canonical name of the property in Active Directory.

    .PARAMETER DisableActions
      Instructs the function to end after displaying the properties of the user instead of providing the user with follow-up actions.

    .PARAMETER Literal
      Signifies that the arguments for every filter should be treated as literals, not wildcard patterns.

    .EXAMPLE
      Get-Computer -Names COMPUTER_NAME

      Retrieves all computers in Active Directory whose name matches "COMPUTER_NAME" and displays a list of properties assosicated with the retrieved computer you select from the table of matches.

    .EXAMPLE
      Get-Computer -Names COMPUTER_*

      Retrieves all computers in Active Directory whose name begins with "COMPUTER_" and displays a list of properties assosicated with the retrieved computer you select from the table of matches.
  #>

  [Alias('gcom')]
  [CmdletBinding()]

  param (
    [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
    [SupportsWildcards()]
    [string]$Names,

    [hashtable[]]$Properties = @(
      @{ Title = 'Name';         CanonName = 'Name' },
      @{ Title = 'IPv4 address'; CanonName = 'IPv4Address' },
      @{ Title = 'IPv6 address'; CanonName = 'IPv6Address' },
      @{ Title = 'Domain name';  CanonName = 'DNSHostName' },
      @{ Title = 'Last logon';   CanonName = 'LastLogonDate' },
      @{ Title = 'Created';      CanonName = 'whenCreated' }
    ),

    [switch]$DisableActions,

    [switch]$Literal
  )

  $searchFilters += @{
    Arguments = $Names
    Properties = @('Name')
  }
  $selectProperties = @(
    @{ Header = '#' },
    @{ Header = 'NAME'; CanonName = 'Name' }
  )

  $searchArguments = @{
    Filters = $searchFilters
    Type = 'computer'
    Properties = $Properties.ForEach({ $_['CanonName'] })
    Literal = $Literal
  }

  $domainObjects = & $internal['Search-Objects'] @searchArguments
  if (-not $domainObjects) {
    Write-Error 'No computers found.'
    return
  }
  $computer = & $internal['Select-ObjectFromTable'] `
      -Objects $domainObjects `
      -Properties $selectProperties
  if (-not $computer) {
    Write-Error 'No selection made.'
    return
  }

  # Records computer's connection status ahead of time.
  $connected = Test-Connection `
      -ComputerName $computer.Name `
      -Count 1 `
      -Quiet

  # Determines max property display name length for later padding.
  $maxLength = 0
  foreach ($property in ($Properties + 'Connection')) {
    if ($property['Title'].Length -gt $maxLength) {
      $maxLength = $property['Title'].Length
    }
  }

  # Writes properties for user.
  Write-Host "`n# INFORMATION #"
  foreach ($property in $Properties) {
    $canonName = $property['CanonName']
    $displayName = $property['Title'].PadRight($maxLength)
    $value = $computer.$canonName
    if ($value -is [datetime]) {
      $date = $value.ToString('yyyy-MM-dd HH:mm:ss')
    }
    switch ($canonName) {
      'LastLogonDate' {
        Write-Host "$displayName : $date"
      } 'whenCreated' {
        Write-Host "$displayName : $date"
      } default {
        if ($value -is [datetime]) {
          Write-Host "$displayName : $date"
        } else {
          Write-Host "$displayName : $value"
        }
      }
    }
  }
  $displayName = 'Connection'.PadRight($maxLength)
  if ($connected) {
    Write-Host "$displayName : online" -ForegroundColor 'green'
  } else {
    Write-Host "$displayName : offline" -ForegroundColor 'red'
  }

  if ($DisableActions) {
    return
  }

  # Prepares the "actions" menu.
  $actions = @(
    'End',
    'Reload'
  )
  if ($connected) {
    $actions += 'Ping (until offline)'
    if ($PSVersionTable.PSVersion.Major -ge 6) {
      if ($IsWindows) {
        $actions += 'Remote'
      }
    } else {
      Write-Host ''
      Write-Warning '"Remote" action only functional on Windows.'
      $actions += 'Remote'
    }
    $actions += 'Restart'
    $actions += 'Power off'
  } else {
    $actions += 'Ping (until online)'
  }

  # Prints "actions" menu.
  :actionLoop while ($true) {
    Write-Host ''
    Write-Host '# ACTIONS #'
    $actionMenu = ''
    $i = 1
    foreach ($action in $actions) {
      $actionMenu += "[$i] $action  "
      $i += 1
    }
    Write-Host "$actionMenu"
    Write-Host ''

    # Validates and processes selection response.
    $selection = $null
    $selection = Read-Host 'Action'
    if (-not $selection) {
      break actionLoop
    }
    try {
      $selection = [int]$selection
    } catch {
      Write-Error ("Invalid selection. Expected a number 1-$($actions.Count).")
      continue actionLoop
    }
    if (($selection -lt 1) -or ($selection -gt $actions.Count)) {
      Write-Error ("Invalid selection. Expected a number 1-$($actions.Count).")
      continue actionLoop
    }
    $selection = $actions[$selection - 1]

    # Executes selected action.
    switch ($selection) {
      'End' {
        Write-Host ''
        break actionLoop
      } 'Ping (until offline)' {
        Write-Host ''
        Write-Host '***Press <control>+<C> to escape.***' -ForegroundColor 'red'
        Write-Host ''
        $passes = 0
        while ($passes -lt 2) {
          $connected = Test-Connection `
              -ComputerName $computer.IPv4Address `
              -Count 2 `
              -Delay 3 `
              -Quiet
          $time = (Get-Date).ToString('HH:mm:ss')
          if ($connected) {
            Write-Host "[$time] Response received (online)." `
                -ForegroundColor 'green'
          } else {
            $passes += 1
            Write-Host "[$time] No response received (offline)." `
                -ForegroundColor 'red'
          }
        }
        Get-Computer $computer.Name
        break actionLoop
      } 'Ping (until online)' {
        Write-Host ''
        Write-Host '***Press <control>+<C> to escape.***' -ForegroundColor 'red'
        Write-Host ''
        $passes = 0
        while ($passes -lt 2) {
          $connected = Test-Connection `
              -ComputerName $computer.IPv4Address `
              -Count 2 `
              -Delay 3 `
              -Quiet
          $time = (Get-Date).ToString('HH:mm:ss')
          if ($connected) {
            $passes += 1
            Write-Host "[$time] Response received (online)." `
                -ForegroundColor 'green'
          } else {
            Write-Host "[$time] No response received (offline)." `
                -ForegroundColor 'red'
          }
        }
        Get-Computer $computer.Name
        break actionLoop
      } 'Power off' {
        Write-Host ''
        Write-Host "Powering ""$($computer.Name)"" off..."
        Stop-Computer -ComputerName $computer.IPv4Address -Force
        continue actionLoop
      } 'Reload' {
        Get-Computer $computer.Name
        break actionLoop
      } 'Remote' {
        Set-Clipboard -Value $($computer.Name)
        Write-Host ''
        Write-Host 'Copied...'
        Write-Host $($computer.Name) -ForegroundColor 'green'
        Write-Host ''
        Write-Host "Connecting to ""$($computer.Name)""..."
        & 'msra.exe' '/offerra' $computer.IPv4Address
        continue actionLoop
      } 'Restart' {
        Write-Host ''
        Write-Host "Restarting ""$($computer.Name)""..."
        Restart-Computer -ComputerName $computer.IPv4Address -Force
        continue actionLoop
      } default {
        Write-Host ''
        Write-Error 'Undefined action.'
        continue actionLoop
      }
    }
  }
}

function Get-Group {
  <#
    .SYNOPSIS
      Allows you to search for groups in Active Directory using various parameters such as Names and Descriptions. At least one argument of every parameter used must match with a group for the group to be considered a match.

      ALIAS: ggroup

    .PARAMETER Names
      Represents the name(s) to search Active Directory groups for. If used, at least one of the arguments passed must match a group's name for the group to be considered a match.

    .PARAMETER Descriptions
      Represents the description(s) to search Active Directory groups for. If used, at least one of the arguments passed must match a group's description for the group to be considered a match.

      ALIAS: desc

    .PARAMETER Properties
      Represents an array of dictionaries (hashtables), each containing two key-value pairs. (1) The first is called "Title" and its value represents the string that will be used as the displayed name of the property in the property list. (2) The second one is called "CanonName" and its value represents the canonical name of the property in Active Directory.

    .PARAMETER DisableActions
      Instructs the function to end after displaying the properties of the user instead of providing the user with follow-up actions.

    .PARAMETER Literal
      Signifies that the arguments for every filter should be treated as literals, not wildcard patterns.

    .EXAMPLE
      Get-Group -Names GROUP_NAME

      Retrieves all groups in Active Directory whose name matches "GROUP_NAME" and displays a list of properties assosicated with the retrieved group you select from the table of matches.

    .EXAMPLE
      Get-Group -Names GROUP_*

      Retrieves all groups in Active Directory whose name begins with "GROUP_" and displays a list of properties assosicated with the retrieved group you select from the table of matches.

    .EXAMPLE
      Get-Group -Descriptions '*email-group@example.com*'

      Retrieves all groups in Active Directory whose description contains "email-group@example.com" and displays a list of properties assosicated with the retrieved group you select from the table of matches.

    .EXAMPLE
      Get-Group -Descriptions '*\\company\department\unit*'

      Retrieves all groups in Active Directory whose description contains "\\company\department\unit" and displays a list of properties assosicated with the retrieved group you select from the table of matches.
  #>

  [Alias('ggroup')]
  [CmdletBinding()]

  param (
    [Parameter(Position=0, ValueFromPipeline=$true)]
    [SupportsWildcards()]
    [string[]]$Names,

    [Alias('desc')]
    [SupportsWildcards()]
    [string[]]$Descriptions,

    [SupportsWildcards()]
    [string[]]$Emails,

    [hashtable[]]$Properties = @(
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
    [switch]$DisableActions,

    [switch]$Literal
  )

  if (-not ($Names -or $Descriptions -or $Emails)) {
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
  if ($Names) {
    $searchFilters += @{
      Arguments = $Names
      Properties = @('Name', 'DisplayName', 'SamAccountName')
    }
  }
  if ($Descriptions) {
    $searchFilters += @{
      Arguments = $Descriptions
      Properties = @('Description')
    }
  }
  if ($Emails) {
    $searchFilters += @{
      Arguments = $Emails
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

  $domainObjects = & $internal['Search-Objects'] @searchArguments
  if (-not $domainObjects) {
    Write-Error 'No groups found.'
    return
  }
  $group = & $internal['Select-ObjectFromTable'] `
      -Objects $domainObjects `
      -Properties $selectProperties
  if (-not $group) {
    Write-Error 'No selection made.'
    return
  }

  # Prepares the "actions" menu.
  <# Prepares the menu here to reduce lag between when the "information"
     section is printed and when the "actions" menu is printed. #>
  $actions = @(
    'End'
  )
  if ($domainObjects.Count -ne 1) {
    $actions += 'Return to search'
  }
  if ($ILLEGAL_GROUPS -notcontains $group.Name) {
    $actions += 'Add users'
    $actions += 'Remove users'
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
    if ($value -is [datetime]) {
      $date = $value.ToString('yyyy-MM-dd HH:mm:ss')
      if ($value -lt (Get-Date)) {
        $diff = (Get-Date) - $value
        $timeSince = '{0}D : {1}H : {2}M ago' `
            -f $diff.Days, $diff.Hours, $diff.Minutes
      }
    }
    switch ($canonName) {
      'Created' {
        Write-Host "$displayName : $date ($timeSince)"
      } 'info' {
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
      } 'Modified' {
        Write-Host "$displayName : $date ($timeSince)"
      } default {
        if ($value -is [datetime]) {
          Write-Host "$displayName : $date"
        } else {
          Write-Host "$displayName : $value"
        }
      }
    }
  }
  Write-Host ''

  if ($DisableActions) {
    return
  }

  # Prints "actions" menu.
  :actionLoop while ($true) {
    Write-Host '# ACTIONS #'
    $actionMenu = ''
    $i = 1
    foreach ($action in $actions) {
      $actionMenu += "[$i] $action  "
      $i += 1
    }
    Write-Host "$actionMenu"
    Write-Host ''

    # Requests selection from the user.
    $selection = $null
    $selection = Read-Host 'Action'
    if (-not $selection) {
      break actionLoop
    }
    try {
      $selection = [int]$selection
    } catch {
      Write-Error ("Invalid selection. Expected a number 1-$($actions.Count).")
      continue actionLoop
    }
    if (($selection -lt 1) -or ($selection -gt $actions.Count)) {
      Write-Error "Invalid selection. Expected a number 1-$($actions.Count)."
      continue actionLoop
    }
    $selection = $actions[$selection - 1]

    # Executes selection.
    Write-Host ''
    switch ($selection) {
      'End' {
        break actionLoop
      } 'Return to search' {
        $group = & $internal['Select-ObjectFromTable'] `
            -Objects (& $internal['Search-Objects'] @searchArguments) `
            -Properties $selectProperties
        Get-Group $group.SamAccountName
        break actionLoop
      } 'Add users' {
        Write-Host ('You may add multiple users by separating them with a ' `
            + "comma (no space).`n")
        $users = (Read-Host 'Users to add') -split ','
        Add-ADGroupMember -Identity $group -Members $users
        Write-Host ''
        continue actionLoop
      } 'Remove users' {
        Write-Host ('You may remove multiple users by separating them with a ' `
            + "comma (no space).`n")
        $users = (Read-Host 'Users to remove') -split ','
        Remove-ADGroupMember -Identity $group -Members $users -Confirm:$false
        Write-Host ''
        continue actionLoop
      } 'Search manager' {
        Get-User `
            -Name (($group.ManagedBy) -split ',')[0].Substring(3)
        break actionLoop
      } default {
        Write-Error 'Invalid selection. Expected a number 1-' `
            + $actions.Count + '.'
        continue actionLoop
      }
    }
  }
}

function Get-User {
  <#
    .SYNOPSIS
      Allows you to search for users in Active Directory using various parameters such as Username, Employee ID, Name, Phone, and Email. At least one argument of every parameter used must match with a user for the user to be considered a match.

      By default, after the user's property list has been printed, you are then presented with a variety of actions that can be performed, such as unlocking the account, resetting the password, reloading the list, et cetera. This can be disabled using the DisableActions switch.

      ALIAS: guser

    .PARAMETER Usernames
      Represents the username(s) to search Active Directory users for. If used, at least one of the arguments passed must match a user's username for the user to be considered a match.

      ALIAS: user

    .PARAMETER EmployeeIDs
      Represents the employee ID(s) to search Active Directory users for. If used, at least one of the arguments passed must match a user's employee ID for the user to be considered a match.

      ALIAS: eid

    .PARAMETER Names
      Represents the name(s) to search Active Directory users for. If used, at least one of the arguments passed must match a user's name for the user to be considered a match.

    .PARAMETER Phones
      Represents the phone number(s) to search Active Directory users for. If used, at least one of the arguments passed must match a user's phone number for the user to be considered a match.

    .PARAMETER Emails
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
      Get-User -EmployeeIDs 12*45

      Retrieves all users in Active Directory whose employee ID matches the wildcard pattern "12*45" (such as those under the employee ID "12345", "12445", et cetera) and displays a list of properties assosicated with the retrieved user you select from the table of matches.

    .EXAMPLE
      Get-User -Names 'j* doe'

      Retrieves all users in Active Directory whose name matches the wildcard pattern "j* doe" (such as those with the employee ID "12345", "12445", et cetera) and displays a list of properties assosicated with the retrieved user you select from the table of matches.

    .EXAMPLE
      Get-User -Phones *123*456*7890*

      Retrieves all users in Active Directory whose phone number (either personal or work phone numbers) matches the wildcard pattern "*123*456*7890*" (such as those with the phone number "+1 123 456 7890", "+11234567890", "123/456-7890", "123-456-7890", et cetera) and displays a list of properties assosicated with the retrieved user you select from the table of matches.

    .EXAMPLE
      Get-User -Emails 'john-doe@example.com'

      Retrieves all users in Active Directory whose email address (either personal or work email addresses) matches the wildcard pattern "john-doe@example.com" and displays a list of properties assosicated with the retrieved user you select from the table of matches.

    .EXAMPLE
      Get-User -Names 'Jo*n Doe' -Phones +1*123*456*, +1*123*789* -Emails *example.com

      Retrieves all users in Active Directory whose name matches the wildcard pattern "Jo*n Doe" **and** whose phone number(s) match either "+1*123*456*" or "+1*123*789*" **and** whose email address matches "*example.com". Then displays a list of properties assosicated with the retrieved user you select from the table of matches.
  #>

  [Alias('guser')]
  [CmdletBinding()]

  param (
    [Alias('user')]
    [Parameter(Position=0, ValueFromPipeline=$true)]
    [SupportsWildcards()]
    [string[]]$Usernames,

    [Alias('eid')]
    [SupportsWildcards()]
    [string[]]$EmployeeIDs,

    [SupportsWildcards()]
    [string[]]$Names,

    [SupportsWildcards()]
    [string[]]$Phones,

    [SupportsWildcards()]
    [string[]]$Emails,

    [hashtable[]]$Properties = @(
      @{ Title = 'Name (display)';  CanonName = 'DisplayName' },
      @{ Title = '       (legal)';  CanonName = 'Name' },
      @{ Title = 'Username';        CanonName = 'SamAccountName' },
      @{ Title = 'Employee ID';     CanonName = 'EmployeeID' },
      @{ Title = 'Email';           CanonName = 'EmailAddress' },
      @{ Title = 'Phone (ipPhone)'; CanonName = 'ipPhone' },
      @{ Title = '  (otherMobile)'; CanonName = 'otherMobile' },
      @{ Title = '       (mobile)'; CanonName = 'mobile' },
      @{ Title = '  (MobilePhone)'; CanonName = 'MobilePhone' },
      @{ Title = '  (~teleNumber)'; CanonName = 'telephoneNumber' },
      @{ Title = '   (~otherHome)'; CanonName = 'otherHomePhone' },
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
    [switch]$DisableActions,

    [switch]$Literal
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

  if ($Usernames) {
    $searchFilters += @{
      Arguments = $Usernames
      Properties = @('SamAccountName')
    }
  }
  if ($EmployeeIDs) {
    $searchFilters += @{
      Arguments = $EmployeeIDs
      Properties = @('EmployeeID')
    }
    $selectProperties += @{
      Header = 'EMPLOYEE ID'
      CanonName = 'EmployeeID'
    }
  }
  if ($Names) {
    $searchFilters += @{
      Arguments = $Names
      Properties = @('Name', 'DisplayName')
    }
  }
  if ($Phones) {
    $searchFilters += @{
      Arguments = $Phones
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
  if ($Emails) {
    $searchFilters += @{
      Arguments = $Emails
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

  $domainObjects = & $internal['Search-Objects'] @searchArguments
  if (-not $domainObjects) {
    Write-Error 'No users found.'
    return
  }
  $user = & $internal['Select-ObjectFromTable'] `
      -Objects $domainObjects `
      -Properties $selectProperties
  if (-not $user) {
    Write-Error 'No selection made.'
    return
  }

  # Prepares the "actions" menu.
  <# Prepares the menu here to reduce lag between when the "information"
     section is printed and when the "actions" menu is printed. #>
  $actions = @(
    'End',
    'Reload',
    'Summarize',
    'Reset password',
    'List groups',
    'Add groups',
    'Remove groups'
  )
  if (($user.LockedOut) -or ($user.AccountLockoutTime)) {
    $actions += 'Unlock'
    $actions += 'Unlock (copy ticket)'
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
  Write-Host "`n# INFORMATION #"
  foreach ($property in $Properties) {
    $canonName = $property['CanonName']
    $displayName = $property['Title'].PadRight($maxLength)
    $value = $user.$canonName
    if ($value -is `
        [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
      $value = $value.Value
    }  # [HACK]
    if ($value -is [datetime]) {
      $date = $value.ToString('yyyy-MM-dd HH:mm:ss')
      if ($value -lt (Get-Date)) {
        $diff = (Get-Date) - $value
        $timeSince = '{0}D : {1}H : {2}M ago' `
            -f $diff.Days, $diff.Hours, $diff.Minutes
      }
    }
    switch ($canonName) {
      'accountExpires' {
        if ($value) {
          if (($value -eq 0) -or ($value -gt [DateTime]::MaxValue.Ticks)) {
            continue
          } else {
            $value = ([DateTime]$value).AddYears(1600).ToLocalTime()
          }
          $date = $value.ToString('yyyy-MM-dd HH:mm:ss')
          if ($value -lt (Get-Date)) {
            $diff = (Get-Date) - $value
            $timeSince = '{0}D : {1}H : {2}M ago' `
                -f $diff.Days, $diff.Hours, $diff.Minutes
          }
          if ($value -lt (Get-Date)) {
            Write-Host "$displayName : $date ($timeSince)" `
                -ForegroundColor 'red'
          } else {
            Write-Host "$displayName : $date" -ForegroundColor 'green'
          }
        }
      } 'AccountLockoutTime' {
        if ($value) {
          Write-Host "$displayName : $date ($timeSince)" `
              -ForegroundColor 'red'
        }
      } 'Created' {
        if ($value) {
          Write-Host "$displayName : $date ($timeSince)"
        }
      } 'EmployeeID' {
        if ($value) {
          Write-Host "$displayName : $value" -ForegroundColor 'green'
        } else {
          Write-Host "$displayName : $value" -ForegroundColor 'red'
        }
      } 'Enabled' {
        if (-not $value) {
          Write-Host "$displayName : $value" -ForegroundColor 'red'
        }
      } 'LastBadPasswordAttempt' {
        if ($value) {
          Write-Host "$displayName : $date ($timeSince)"
        }
      } 'LastLogonDate' {
        if ($value) {
          Write-Host "$displayName : $date ($timeSince)"
        }
      } 'LockedOut' {
        if ($value) {
          Write-Host "$displayName : $value" -ForegroundColor 'red'
        }
      } 'Manager' {
        if ($value) {
          $value = ($value -split ',')[0].Substring(3)
          Write-Host "$displayName : $value"
        } else {
          Write-Host "$displayName : "
        }
      } 'mobile' {
        if ("$value".Trim()) {
          Write-Host "$displayName : $value" -ForegroundColor 'green'
        }
      } 'MobilePhone' {
        if ("$value".Trim()) {
          Write-Host "$displayName : $value" -ForegroundColor 'green'
        }
      } 'Modified' {
        Write-Host "$displayName : $date ($timeSince)"
      } 'PasswordLastSet' {
        if ($value) {
          if ($diff.Days -ge 90) {
            Write-Host "$displayName : $date ($timeSince)" `
                -ForegroundColor 'red'
          } else {
            Write-Host "$displayName : $date ($timeSince)" `
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
        if ($value -is [datetime]) {
          Write-Host "$displayName : $date"
        } else {
          Write-Host "$displayName : $value"
        }
      }
    }
  }

  if ($DisableActions) {
    return
  }

    # Prints "actions" menu.
  :actionLoop while ($true) {
    Write-Host ''
    Write-Host '# ACTIONS #'
    $actionMenu = ''
    $i = 1
    foreach ($action in $actions) {
      $actionMenu += "[$i] $action  "
      $i += 1
    }
    Write-Host "$actionMenu"
    Write-Host ''

    $selection = $null
    $selection = Read-Host 'Action'
    if (-not $selection) {
      break actionLoop
    }
    try {
      $selection = [int]$selection
    } catch {
      Write-Error ("Invalid selection. Expected a number 1-$($actions.Count).")
      continue actionLoop
    }
    if (($selection -lt 1) -or ($selection -gt $actions.Count)) {
      Write-Error ("Invalid selection. Expected a number 1-$($actions.Count).")
      continue actionLoop
    }
    $selection = $actions[$selection - 1]

    # Executes selected action.
    switch ($selection) {
      'End' {
        Write-Host ''
        break actionLoop
      } 'Reload' {
        Get-User -Usernames $user.SamAccountName
        break actionLoop
      } 'Summarize' {
        $summary = Get-UserSummary -Usernames $user.SamAccountName
        Set-Clipboard $summary
        Write-Host ''
        Write-Host 'Copied...'
        Write-Host $summary -ForegroundColor 'green'
        Write-Host ''
        continue actionLoop
      } 'Return to search' {
        $user = & $internal['Select-ObjectFromTable'] `
            -Objects (& $internal['Search-Objects'] @searchArguments) `
            -Properties $selectProperties
        Get-User $user.SamAccountName
        break actionLoop
      } 'Unlock' {
        Unlock-ADAccount -Identity $user.SamAccountName
        Write-Host ''
        continue actionLoop
      } 'Unlock (copy ticket)' {
        Copy-AccountUnlockTicket -Type 'domain' -Username $user.SamAccountName
        continue actionLoop
      } 'Reset password' {
        Write-Host ''
        Reset-Password -Users $user.SamAccountName
        Unlock-ADAccount -Identity $user.SamAccountName
        Write-Host ''
        continue actionLoop
      } 'List groups' {
        $groups = Get-ADPrincipalGroupMembership -Identity $user.SamAccountName `
            | Get-ADGroup -Properties 'Name', 'Description' | Sort-Object 'Name'
        $table = New-Object System.Data.DataTable
        $headers = @('#', 'NAME', 'DESCRIPTION')
        foreach ($header in $headers) {
          $quiet = $table.Columns.Add($header)
        }
        $i = 0
        foreach ($group in $groups) {
          $i += 1
          $row = $table.NewRow()
          foreach ($header in $headers) {
            switch ($header) {
              '#' {
                $row.'#' = $i
              } 'NAME' {
                $row.'NAME' = $group.Name
              } 'DESCRIPTION' {
                $row.'DESCRIPTION' = $group.Description
              } default {
                Write-Error "Unrecognized header provided (""$header"")."
                return
              }
            }
          }
          $table.Rows.Add($row)
        }
        $table | Format-Table -Wrap
        :groupActions while ($true) {
          Write-Host '# ACTIONS #'
          $selection = $null
          $selection = Read-Host "[0] Return  [1-$i] Load group by #"
          if (-not $selection) {
            Get-User $user.SamAccountName
            break groupActions
          }
          try {
            $selection = [int]$selection
          } catch {
            Write-Error "Invalid selection. Expected a number 0-$i."
            continue groupActions
          }
          if ($selection -eq 0) {
            Get-User $user.SamAccountName
            break groupActions
          } elseif (($selection -lt 0) -or ($selection -gt $i)) {
            Write-Error "Invalid selection. Expected a number 0-$i."
            continue groupActions
          } else {
            Get-Group -Names ($table.Rows[$selection - 1]['NAME'])
            break groupActions
          }
        }
        break actionLoop
      } 'Add groups' {
        Write-Host ''
        Write-Host ('***You may add multiple groups by separating them with ' `
            + "a comma (no space).***`n")
        $groups = (Read-Host 'Groups to add') -split ','
        foreach ($group in $groups) {
          if ($ILLEGAL_GROUPS -contains $group) {
            Write-Error ("Skipping ""$group"". Modifying this group is " `
                + 'restricted by Data Security.')
          } else {
            Add-ADGroupMember -Identity $group -Members $user
          }
        }
        Write-Host ''
        continue actionLoop
      } 'Remove groups' {
        Write-Host ''
        Write-Host ('***You may remove multiple groups by separating them with ' `
            + "a comma (no space).***`n")
        $groups = (Read-Host 'Groups to remove') -split ','
        foreach ($group in $groups) {
          if ($ILLEGAL_GROUPS -contains $group) {
            Write-Error ("Skipping ""$group"". Modifying this group is " `
                + 'restricted by Data Security.')
          } else {
            Remove-ADGroupMember -Identity $group -Members $user -Confirm:$false
          }
        }
        Write-Host ''
        continue actionLoop
      } 'Search manager' {
        Get-User -Names (($user.Manager) -split ',')[0].Substring(3)
        break actionLoop
      } 'Send email' {
        Start-Process "mailto:$($user.EmailAddress)"
        continue actionLoop
      } default {
        Write-Host ''
        Write-Error 'Undefined action.'
        continue actionLoop
      }
    }
  }
}

function Get-UserSummary {
  <#
    .SYNOPSIS
      Returns a short list of properties for one or more users in Active Directory.

      ALIAS: summary

    .PARAMETER Usernames
      Represents the username(s) of one or more Active Directory users to be included in the list.

      ALIAS: users

    .PARAMETER Properties
      Represents an array of dictionaries (hashtables), each containing two key-value pairs. (1) The first is called "Title" and its value represents the string that will be used as the displayed name of the property in the property list. (2) The second one is called "CanonName" and its value represents the canonical name of the property in Active Directory.

    .EXAMPLE
      Get-UserSummary JD012345

      Returns a list of various properties of the user with the username "JD012345" omitting any properties the user lacks a value for.
  #>

  [Alias('summary')]

  param (
    [Alias('users')]
    [Parameter(
      HelpMessage=('Enter the username(s) of the user(s).'),
      Mandatory=$true,
      Position=0,
      ValueFromPipeline=$true
    )]
    [string[]]$Usernames,

    [hashtable[]]$Properties = @(
      @{ Title = 'Username';    CanonName = 'SamAccountName' },
      @{ Title = 'Employee ID'; CanonName = 'EmployeeID' },
      @{ Title = 'Email';       CanonName = 'EmailAddress' },
      @{ Title = 'Department';  CanonName = 'Department' },
      @{ Title = 'Job title';   CanonName = 'Title' }
    )
  )

  # Retries all relevant users from Active Directory.
  $users = $Usernames.ForEach({
    Get-ADUser `
        -Identity $_ `
        -Properties ($Properties.ForEach({ $_['CanonName'] }) + 'DisplayName')
  })

  $delimiter = @{
    User = "`n`n"
    Property = "`n- "
    TitleValue = ': '
  }
  $summary = ''
  foreach ($user in $users) {
    if ($summary) {
      $summary += $delimiter['User']
    }
    $summary += $user.DisplayName
    foreach ($property in $Properties) {
      $canonName = $property['CanonName']
      $title = $property['Title']
      $value = $user.$canonName
      if ($value) {
        $summary += $delimiter['Property'] + $title `
            + $delimiter['TitleValue'] + $value
      }
    }  
  }

  return $summary
}

function Reformat-Names {
  <#
    .SYNOPSIS
      Takes a list of names written in the "last name, first name" format and returns a table of the referenced users with each row containg, the status (disabled account, name not found, et cetera), a number representing the index when there is a list of accounts found under the given name, the name searched (if the LooseSearch or LooserSearch parameter is used), the account's  display name, and the account's username.

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
    [string]$List,

    [Alias('entrybreak')]
    [string]$EntryDelimiter = "`n",

    [Alias('namebreak')]
    [string]$NameDelimiter = ', ',

    [Alias('loose')]
    [switch]$LooseSearch,

    [Alias('looser')]
    [switch]$LooserSearch
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
    [string]$Users,

    [switch]$Force
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

function Set-CopyPastes {
  <#
    .SYNOPSIS
      Generates various global string variables in PowerShell representing copy-pastes built using the copy-paste documents (determines the name and content of the variable) (see Access-CopyPastes) and variable substitution (see Insert-Variables.ps1).
  #>

  $scope = 'Global'
  $extension = '.txt'
  $copyPasteDirectory = 'copypastes'
  $copyPastePath = [System.IO.Path]::Combine($PSScriptRoot, $copyPasteDirectory)
  $names = ([System.IO.Path]::Combine($copyPastePath, "*$extension") `
      | Get-ChildItem -Name).ForEach({
    $_.SubString(0, ($_.Length - $extension.Length))
  })
  $paths = ([System.IO.Path]::Combine($copyPastePath, "*$extension") `
      | Get-ChildItem -Name).ForEach({
        [System.IO.Path]::Combine($copyPastePath, $_)
      })
  $contents = $paths.ForEach({
    ((Get-Content $_ -Encoding 'utf8') -join "`n").TrimEnd("`n")
  })
  $numFiles = $names.Count

  for ($i = 0; $i -lt $numFiles; $i += 1) {
    New-Variable `
        -Name $names[$i] `
        -Value (& $internal['Insert-Variables'] $contents[$i]) `
        -Scope $scope `
        -Option 'ReadOnly'`
        -ErrorAction 'SilentlyContinue'
  }
}

function Update-GroupMemberships {
  <#
    .SYNOPSIS
      Allows you to add/remove one or more users to/from one or more groups in Active Directory.

      ALIAS: update

    .PARAMETER GroupNames
      Represents the name(s) of the groups you would like to add/remove the users to/from.

      ALIAS: groups

    .PARAMETER Usernames
      Represents the username(s) of the users you would like to add/remove to/from the group(s).

      ALIAS: users

    .PARAMETER Add
      Signals that the function should add the specified users to the specified groups.

    .PARAMETER Remove
      Signals that the function should remove the specified users from the specified groups.

    .EXAMPLE
      Update-GroupMemberships -GroupNames 'GROUP_1', 'GROUP_2' -Usernames 'USER_1', 'USER_2' -Add

      Adds the users with the usernames "USER_1" and "USER_2" to the groups named "GROUP_1" and "GROUP_2".

    .EXAMPLE
      Update-GroupMemberships -GroupNames 'GROUP_1', 'GROUP_2' -Usernames 'USER_1', 'USER_2' -Remove

      Removes the users with the usernames "USER_1" and "USER_2" from the groups named "GROUP_1" and "GROUP_2".

    .EXAMPLE
      Update-GroupMemberships -groups 'GROUP_1', 'GROUP_2' -users 'USER_1', 'USER_2' -Add

      Adds the users with the usernames "USER_1" and "USER_2" to the groups named "GROUP_1" and "GROUP_2".
  #>

  [Alias('update')]
  [CmdletBinding()]

  param (
    [Alias('groups')]
    [Parameter(Position=0)]
    [string[]]$GroupNames,

    [Alias('users')]
    [Parameter(Position=1)]
    [string[]]$Usernames,

    [switch]$Add,

    [switch]$Remove
  )

  if (-not ($Add -or $Remove)) {
    Write-Error 'You must specify whether this is to add members or remove ' `
        + 'members using the Add and Remove switches.'
    return
  } elseif ($Add -and $Remove) {
    Write-Error 'You cannot use the Add and Remove switches simultaneously.'
    return
  }

  $groups = $GroupNames.ForEach({
    if ($ILLEGAL_GROUPS -contains $_) {
      Write-Error ('Skipping group ("' + $_ + '"). Modifying this group is ' `
          + 'restricted by Data Security.')
      continue
    }
    $results = Get-ADGroup `
        -Filter "(SamAccountName -eq ""$_"") -or (Name -eq ""$_"")" `
        -Properties 'ObjectGUID', 'SamAccountName'
    if ($results) {
      if ($results.Count -gt 1) {
        Write-Error 'Skipping group ("' + $_ `
            + '"). Multiple groups found with this name.'
      } else {
        $results
      }
    } else {
      Write-Error "Skipping group (""$_""). No groups found with this name."
    }
  })
  $users = $Usernames.ForEach({
    $results = Get-ADUser $_ -Properties 'ObjectGUID', 'SamAccountName'
    if ($results) {
      $results
    } else {
      Write-Error "Skipping user (""$_""). No users found with this username."
    }
  })

  if (-not $groups) {
    Write-Error 'No groups found.'
    return
  }
  if (-not $users) {
    Write-Error 'No users found.'
    return
  }

  foreach ($group in $groups) {
    if ($Add) {
      Add-ADGroupMember -Identity $group.ObjectGUID -Members $users
    } elseif ($Remove) {
      Remove-ADGroupMember -Identity $group.ObjectGUID -Members $users
    }
  }
}
