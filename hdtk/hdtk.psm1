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

<# Imports all of the functions defined in internal-functions.ps1. 
   These functions are defined there as they are only intended 
   as internal/private functions. #>
. "$([System.IO.Path]::Combine($PSScriptRoot, 'internal-functions.ps1'))"

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
  Copy-Ticket @parameters
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

  $i = 1
  $footnotes = $Computers.ForEach({ "[$i]: ""$_"""; $i += 1 })
  $footnotes += $Printers.ForEach({ "[$i]: ""$_"""; $i += 1 })
  $footnotes = $footnotes -join "`n"
  $i = 1
  $references = @{
    Computers = $Computers.ForEach({ "[$i]"; $i += 1 }) -join ' '
    Printers = $Printers.ForEach({ "[$i]"; $i += 1 }) -join ' '
  }

  $subject = 'connect computer(s) to printer(s)'
  $body = 'Customer requested to have '
  $fulfillment = 'Connected '

  if ($Computers.Count -gt 1) {
    $body += "some computers $($references.Computers) and "
    $fulfillment += "the computers $($references.Computers) to "
  } else {
    $body += "a computer $($references.Computers) and "
    $fulfillment += "the computer $($references.Computers) to "
  }
  if ($Printers.Count -gt 1) {
    $body += "some printers $($references.Printers) connected."
    $fulfillment += "the printers $($references.Printers)."
  } else {
    $body += "a printer $($references.Printers) connected."
    $fulfillment += "the printer $($references.Printers)."
  }
  $body += "`n`n$footnotes"
  $fulfillment += "`n`n$footnotes"

  $parameters = @{}
  if (-not $DisableSubjectCopy) {
    $parameters['Subject'] = $subject
  }
  $parameters['Body'] = $body
  if (-not $DisableFulfillmentCopy) {
    $parameters['Fulfillment'] = $fulfillment
  }
  Copy-Ticket @parameters
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
    [switch]$DisableFulfillmentCopy
  )

  $subject = ''
  $body = ''
  $fulfillment = ''

  $i = 1
  $footnotes = (
    $Paths.ForEach({ "[$i]: ""$_"""; $i += 1 }) `
        + $Computers.ForEach({ "[$i]: ""$_"""; $i += 1 })
  ) -join "`n"
  $i = 1
  $references = @{
    Paths = $Paths.ForEach({ "[$i]"; $i += 1 }) -join ' '
    Computers = $Computers.ForEach({ "[$i]"; $i += 1 }) -join ' '
  }

  $subject = 'map drive(s)'

  if ($Paths.Count -gt 1) {
    $body = "Customer requested to have drives $($references.Paths) "
    $fulfillment = "Mapped the paths $($references.Paths) to some drives for the "
  } else {
    $body = "Customer requested to have a drive $($references.Paths) "
    $fulfillment = "Mapped the path $($references.Paths) to a drive for the "
  }
  if ($Computers.Count -gt 1) {
    $body += "mapped for some computers $($references.Computers)."
    $fulfillment += "computers $($references.Computers)."
  } else {
    $body += "mapped for a computer $($references.Computers)."
    $fulfillment += "computer $($references.Computers)."
  }
  $body += "`n`n$footnotes"
  $fulfillment += "`n`n$footnotes"

  $parameters = @{}
  if (-not $DisableSubjectCopy) {
    $parameters['Subject'] = $subject
  }
  $parameters['Body'] = $body
  if (-not $DisableFulfillmentCopy) {
    $parameters['Fulfillment'] = $fulfillment
  }
  Copy-Ticket @parameters
}

function Copy-PinFolderTicket {
  <#
    .SYNOPSIS
      Builds and copies the body and subject of a ticket for pinning folder(s) to File Explorer's Quick Access.

      ALIAS: pin

    .PARAMETER Computers
      Represents the name(s) of the computer(s) having the folder(s) pinned.

      ALIAS: pc

    .PARAMETER Paths
      Represents the path(s) of the folder(s) being pinned to Quick Access in File Explorer.

    .PARAMETER DisableSubjectCopy
      Signifies that the function should not offer to copy the ticket subject after waiting for you to press <enter>.

      ALIAS: nosubject

    .PARAMETER DisableFulfillmentCopy
      Signifies that the function should not offer to copy the ticket fulfillment comment after waiting for you to press <enter>.

      ALIAS: nocomment

    .EXAMPLE
      Copy-PinFolderTicket -Computers COMPUTER_1 -Paths PATH_1

      Copies, then prints the ticket body, subject, and fulfillment comment.

    .EXAMPLE
      Copy-PinFolderTicket 'COMPUTER_1', 'COMPUTER_2' 'PATH_1', 'PATH_2'

      Copies, then prints the ticket body, subject, and fulfillment comment.

    .EXAMPLE
      Copy-PinFolderTicket COMPUTER_1 PATH_1 -DisableSubjectCopy

      Copies, then prints the ticket body and fulfillment comment.
  #>

  [Alias('pin')]
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
      HelpMessage='Enter the folder path(s).',
      Mandatory=$true,
      Position=1
    )]
    [string[]]$Paths,

    [Alias('nosubject')]
    [switch]$DisableSubjectCopy,

    [Alias('nocomment')]
    [switch]$DisableFulfillmentCopy
  )

  $subject = ''
  $body = ''
  $fulfillment = ''

  $i = 1
  $footnotes = (
    $Paths.ForEach({ "[$i]: ""$_"""; $i += 1 }) `
        + $Computers.ForEach({ "[$i]: ""$_"""; $i += 1 })
  ) -join "`n"
  $i = 1
  $references = @{
    Paths = $Paths.ForEach({ "[$i]"; $i += 1 }) -join ' '
    Computers = $Computers.ForEach({ "[$i]"; $i += 1 }) -join ' '
  }

  $subject = 'File Explorer: pin folder(s) to Quick Access'

  if ($Paths.Count -gt 1) {
    $body = "Customer requested to have some folders $($references.Paths) "
    $fulfillment = "Pinned the folders $($references.Paths) to Quick " `
        + "Access (File Explorer) for the "
  } else {
    $body = "Customer requested to have a folder $($references.Paths) "
    $fulfillment = "Pinned the folder $($references.Paths) to Quick " `
        + "Access (File Explorer) for the "
  }
  if ($Computers.Count -gt 1) {
    $body += "pinned to Quick Access (File Explorer) for some computers " `
        + "$($references.Computers)."
    $fulfillment += "computers $($references.Computers)."
  } else {
    $body += "pinned to Quick Access (File Explorer) for a computer $($references.Computers)."
    $fulfillment += "computer $($references.Computers)."
  }
  $body += "`n`n$footnotes"
  $fulfillment += "`n`n$footnotes"

  $parameters = @{}
  if (-not $DisableSubjectCopy) {
    $parameters['Subject'] = $subject
  }
  $parameters['Body'] = $body
  if (-not $DisableFulfillmentCopy) {
    $parameters['Fulfillment'] = $fulfillment
  }
  Copy-Ticket @parameters
}

function Format-Quote {
  <#
    .SYNOPSIS
      Prefixes every line of the provided string with a Markdown blockquote indicator ("> "). The level of blockquote nesting for the text may be specified with the Level parameter; and, to add blockquote syntax to a copied (Outlook-style) email containing a block of headers and a new line before the actual text, you may use the Email parameter.

      ALIAS: quote

    .PARAMETER Text
      Represents either the input text itself or the path to a text file containing the input string. If an invalid path is provided, the string will be treated as the input text itself.

    .PARAMETER Email
      Represents that the input text should not include the block of email headers as part of the blockquote, and that the blank line separating the block of headers of the blockquote should be removed.

      Assumes that the text is the in the Outlook-style format (see below).

      > From: Sender Name <sender-name@example.com>
      > Sent: Thursday, January 1st, 1970 12:01 AM
      > To: Recipient Name <recipient-name@example.com>
      > Subject: Lorem Ipsum
      > 
      > Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.

    .PARAMETER Concise
      Represents that certain header information should be reformatted to be more concise. For example--in the to, from, CC and BCC headers--email address written in the "First Last <first-last@example.com>" format will be rewritten in the more basic "first-last@example.com" format instead.

      Assumes that the text is the in the Outlook-style format (see below).

      > From: Sender Name <sender-name@example.com>
      > Sent: Thursday, January 1st, 1970 12:01 AM
      > To: Recipient Name <recipient-name@example.com>
      > Subject: Lorem Ipsum
      > 
      > Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.

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
      Format-Quote -Email -Concise

      Would return the below:

      From: sender@example.com
      Sent: 1970-01-31, 00:01
      To: recipient@example.com
      Subject: Lorem Ipsum
      > Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.
      > 
      > Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.
  #>

  [Alias('fqt')]
  [CmdletBinding()]

  param (
    [Parameter(
      Position=0,
      ValueFromPipeline=$true
    )]
    [string]$Text = [System.IO.Path]::Combine($HOME, 'Desktop', 'text.txt'),

    [switch]$Concise,

    [switch]$Email
  )

  $lines = @()
  if (Test-Path -Path $Text) {
    $lines = Get-Content -Path $Text -Encoding 'utf8'
  } elseif (Test-Path -Path $Text -IsValid) {
    Write-Error "Unable to locate a text file at the address " `
        + "provided to the Text parameter (""$Text""). "
    $lines = $Text -split "`n"
  } else {
    $lines = $Text -split "`n"
  }
  $Text = ''
  $prefix = '> '
  $inHeaders = $true
  $inBody = $true
  if ($Email) {
    $inHeaders = $true
    $inBody = $false
  } else {
    $inHeaders = $false
    $inBody = $true
  }
  foreach ($line in $lines) {
    if ($inHeaders -and ($line -notmatch '[^ ]:')) {
      $inHeaders = $false
    }
    if ((-not $inBody) -and (-not $inHeaders) -and $line.Trim()) {
      $inBody = $true
    }
    if ($inHeaders) {
      if ((-not $Email) -and $Concise) {
        throw 'The Concise switch cannot be used without the Email switch.'
      }
      $title = ($line -split ': ')[0]
      if ($Concise) {
        if ($title -in @('To', 'From', 'CC', 'BCC')) {
          $emails = $line.Remove(0, ($title.Length + 2)) -split '; '
          $emails = $emails.ForEach({
            $_ -replace '.*<([^>]+)>.*', '$1'
          })
          $emails = $emails -join '; '
          $line = "$title`: $emails"
        } elseif ($title -eq 'Sent') {
          try {
            $date = $line.Remove(0, ($title.Length + 2))
            $date = [DateTime]::Parse($date)
            $date = $date.ToString('yyyy-MM-dd, HH:mm')
            $line = "$title`: $date"
          } catch {}
        }
      }
      $Text += "$line`n"
    } elseif ($inBody) {
      $Text += "$prefix$line`n"
    }
  }

  return $Text.Trim()
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

  $domainObjects = Search-Objects @searchArguments
  if (-not $domainObjects) {
    Write-Error 'No computers found.'
    return
  }
  $computer = Select-ObjectFromTable `
      -Objects $domainObjects `
      -Properties $selectProperties
  [Console]::CursorVisible = $false
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

  :rewrite while ($true) {
    # Clears the window.
    $start = [Console]::CursorStart + 1
    for ($i = $start; $i -le [Console]::WindowHeight; $i += 1) {
      Write-Host (' ' * [Console]::WindowWidth)
    }
    [Console]::SetCursorPosition(0, 0)

    # Writes properties for user.
    Write-Host '# INFORMATION #'
    foreach ($property in $Properties) {
      $canonName = $property['CanonName']
      $displayName = $property['Title'].PadRight($maxLength)
      $value = $computer.$canonName
      if ($value -is [datetime]) {
      $date = $value.ToString('yyyy-MM-dd, HH:mm:ss')
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
      [Console]::CursorVisible = $true
      break rewrite
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

    # Displays actions menu.
    Write-Host '# ACTIONS #'
    $selection = Display-VerticalMenu -Options $actions
    Write-Host ''
    if ($selection -eq $null) {
      $selection = 'End'
    } else {
      $selection = $actions[$selection]
    }

    # Executes selection.
    switch ($selection) {
      'End' {
        Write-Host ''
        break rewrite
      } 'Ping (until offline)' {
        Write-Host "`n"
        Write-Host '***Press <control>+<C> to escape.***' -ForegroundColor 'Red'
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
                -ForegroundColor 'Green'
          } else {
            $passes += 1
            Write-Host "[$time] No response received (offline)." `
                -ForegroundColor 'Red'
          }
        }
        Get-Computer $computer.Name
        break rewrite
      } 'Ping (until online)' {
        Write-Host "`n"
        Write-Host '***Press <control>+<C> to escape.***' -ForegroundColor 'Red'
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
                -ForegroundColor 'Green'
          } else {
            Write-Host "[$time] No response received (offline)." `
                -ForegroundColor 'Red'
          }
        }
        Get-Computer $computer.Name
        break rewrite
      } 'Power off' {
        Write-Host "`n"
        Write-Host "Powering ""$($computer.Name)"" off..."
        Stop-Computer -ComputerName $computer.IPv4Address -Force
        continue rewrite
      } 'Reload' {
        Get-Computer $computer.Name
        break rewrite
      } 'Remote' {
        Set-Clipboard -Value $($computer.Name)
        Write-Host "`n"
        Write-Host 'Copied...'
        Write-Host $($computer.Name) -ForegroundColor 'Green'
        Write-Host ''
        Write-Host "Connecting to ""$($computer.Name)""..."
        & 'msra.exe' '/offerra' $computer.IPv4Address
        continue rewrite
      } 'Restart' {
        Write-Host "`n"
        Write-Host "Restarting ""$($computer.Name)""..."
        Restart-Computer -ComputerName $computer.IPv4Address -Force
        continue rewrite
      } default {
        Write-Host ''
        Write-Error 'Undefined action.'
        continue rewrite
      }
    }
  }
}

function Get-ExclusiveGroups {
  <#
    .SYNOPSIS
      Allows you to get the names of every group that one user in Active Directory is a member of that another is not a member of. Includes an optional Filter paramter that can be used to further filter the groups returned.

    .PARAMETER User
      Represents the account to compare the group memberships of Reference against. The argument for the User parameter may either be an [ADUser object] (such as those returned by Get-ADUser) or a string representing an identifier for an account in Active Directory (corresponds with the identifiers accepted by [Get-ADUser's Identity parameter]).

      [ADUser object]: https://learn.microsoft.com/en-us/dotnet/api/microsoft.activedirectory.management.aduser

      [Get-ADUser's Identity parameter]: https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-aduser#-identity

    .PARAMETER Reference
      Represents the account to compare User's group memberships against. The argument for the Reference parameter may either be an [ADUser object] (such as those returned by Get-ADUser) or a string representing an identifier for an account in Active Directory (corresponds with the identifiers accepted by [Get-ADUser's Identity parameter]).

      [ADUser object]: https://learn.microsoft.com/en-us/dotnet/api/microsoft.activedirectory.management.aduser

      [Get-ADUser's Identity parameter]: https://learn.microsoft.com/en-us/powershell/module/activedirectory/get-aduser#-identity

    .PARAMETER Filter
      Represents the [wilcard expression] to filter out the groups returned. Only groups with names matching this wilcard expression will be returned.

      [wilcard expression]: https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_wildcards

    .EXAMPLE
      Get-ExclusiveGroups -User 'JD012345' -Reference '0012345'

      Returns the names of all groups that 0012345 is a member of in Active Directory that JD012345 is not.

    .EXAMPLE
      Get-ExclusiveGroups 'JD012345' '0012345' -Filter 'Citrix*'

      Returns the names of all groups starting with "Citrix" that 0012345 is a member of in Active Directory that JD012345 is not.

    .EXAMPLE
      Get-ExclusiveGroups 'JD012345' '0012345' '*Microsoft*'

      Returns the names of all groups containing "Microsoft" that 0012345 is a member of in Active Directory that JD012345 is not.
  #>

  [CmdletBinding()]

  param (
    [Parameter(Mandatory=$true, Position=0)]
    [object]$User,

    [Alias('ref')]
    [Parameter(Mandatory=$true, Position=1)]
    [object]$Reference,

    [Parameter(Position=2)]
    [SupportsWildcards()]
    [string]$Filter = '*'
  )

  if ($User.GetType().Name -eq 'String') {
    try {
      $User = Get-ADUser $User
    } catch {
      throw $_
    }
  } elseif ($User.GetType().Name -ne 'ADUser') {
    Write-Error """$($User.GetType().Name)"" is not an acceptable type for the User parameter. Argument must be either an ADUser object or a string identifying an ADUser object (correspond's with Get-ADUser's Identity parameter)."
    return
  }

  if ($Reference.GetType().Name -eq 'String') {
    try {
      $Reference = Get-ADUser $Reference
    } catch {
      throw $_
    }
  } elseif ($Reference.GetType().Name -ne 'ADUser') {
    Write-Error """$($Reference.GetType().Name)"" is not an acceptable type for the Reference parameter. Argument must be either an ADUser object or a string identifying an ADUser object (correspond's with Get-ADUser's Identity parameter)."
    return
  }

  $userGroups = Get-ADPrincipalGroupMembership `
      -Identity $User.SamAccountName | Get-ADGroup -Properties 'Name' `
      | Select-Object -ExpandProperty 'Name'
  $referenceGroups = Get-ADPrincipalGroupMembership `
      -Identity $Reference.SamAccountName | Get-ADGroup -Properties 'Name' `
      | Select-Object -ExpandProperty 'Name'

  $exclusives = @()
  foreach ($group in $referenceGroups) {
    if (($userGroups -notcontains $group) -and ($group -like $Filter)) {
      $exclusives += $group
    }
  }

  return $exclusives | Sort-Object
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

  $domainObjects = Search-Objects @searchArguments
  if (-not $domainObjects) {
    Write-Error 'No groups found.'
    return
  }
  $group = Select-ObjectFromTable `
      -Objects $domainObjects `
      -Properties $selectProperties
  [Console]::CursorVisible = $false
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

  :rewrite while ($true) {
    # Clears the window.
    $start = [Console]::CursorStart + 1
    for ($i = $start; $i -le [Console]::WindowHeight; $i += 1) {
      Write-Host (' ' * [Console]::WindowWidth)
    }
    [Console]::SetCursorPosition(0, 0)

    # Writes properties for user.
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
        $date = $value.ToString('yyyy-MM-dd, HH:mm:ss')
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
      [Console]::CursorVisible = $true
      break rewrite
    }

    # Displays actions menu.
    Write-Host '# ACTIONS #'
    $selection = Display-VerticalMenu -Options $actions
    Write-Host ''
    if ($selection -eq $null) {
      $selection = 'End'
    } else {
      $selection = $actions[$selection]
    }

    # Executes selection.
    switch ($selection) {
      'End' {
        Write-Host ''
        break rewrite
      } 'Return to search' {
        Write-Host ''
        $group = Select-ObjectFromTable `
            -Objects (Search-Objects @searchArguments) `
            -Properties $selectProperties
        Get-Group $group.SamAccountName
        break rewrite
      } 'Add users' {
        Write-Host ''
        Write-Host ('You may add multiple users by separating them with a ' `
            + "comma.`n")
        $users = ((Read-Host 'Users to add') -split ',').Trim()
        Add-ADGroupMember -Identity $group -Members $users
        Write-Host ''
        break rewrite
      } 'Remove users' {
        Write-Host ''
        Write-Host ('You may remove multiple users by separating them with a ' `
            + "comma.`n")
        $users = ((Read-Host 'Users to remove') -split ',').Trim()
        Remove-ADGroupMember -Identity $group -Members $users -Confirm:$false
        Write-Host ''
        break rewrite
      } 'Search manager' {
        Write-Host ''
        Get-User `
            -Name (($group.ManagedBy) -split ',')[0].Substring(3)
        break rewrite
      } default {
        Write-Host ''
        Write-Error 'Unrecognized action.'
        break rewrite
      }
    }
    [Console]::CursorVisible = $true
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
      @{ Title = 'X profile';       CanonName = 'ProfilePath' },
      @{ Title = 'Object address';  CanonName = 'CanonicalName' }
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
    if ($Literal) {
      $searchFilters += @{
        Arguments = $Names
        Properties = @('Name', 'DisplayName')
      }
    } else {
      $searchFilters += @{
        Arguments = $Names.ForEach({ $_.Replace(' ', '*').Replace('-', '*') })
        Properties = @('Name', 'DisplayName')
      }
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

  $domainObjects = Search-Objects @searchArguments
  if (-not $domainObjects) {
    Write-Error 'No users found.'
    return
  }
  $user = Select-ObjectFromTable `
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

  :rewrite while ($true) {
    # Clears the window.
    $start = [Console]::CursorStart + 1
    for ($i = $start; $i -le [Console]::WindowHeight; $i += 1) {
      Write-Host (' ' * [Console]::WindowWidth)
    }
    [Console]::SetCursorPosition(0, 0)

    # Writes properties for user.
    Write-Host '# INFORMATION #'
    foreach ($property in $Properties) {
      $canonName = $property['CanonName']
      $displayName = $property['Title'].PadRight($maxLength)
      $value = $user.$canonName
      if ($value -is `
          [Microsoft.ActiveDirectory.Management.ADPropertyValueCollection]) {
        $value = $value.Value
      }  # [HACK]
      if ($value -is [datetime]) {
        $date = $value.ToString('yyyy-MM-dd, HH:mm:ss')
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
            $date = $value.ToString('yyyy-MM-dd, HH:mm:ss')
            if ($value -lt (Get-Date)) {
              $diff = (Get-Date) - $value
              $timeSince = '{0}D : {1}H : {2}M ago' `
                  -f $diff.Days, $diff.Hours, $diff.Minutes
            }
            if ($value -lt (Get-Date)) {
              Write-Host "$displayName : $date ($timeSince)" `
                  -ForegroundColor 'Red'
            } else {
              Write-Host "$displayName : $date" -ForegroundColor 'Green'
            }
          }
        } 'AccountLockoutTime' {
          if ($value) {
            Write-Host "$displayName : $date ($timeSince)" `
                -ForegroundColor 'Red'
          }
        } 'Created' {
          if ($value) {
            Write-Host "$displayName : $date ($timeSince)"
          }
        } 'EmployeeID' {
          if ($value) {
            Write-Host "$displayName : $value" -ForegroundColor 'Green'
          } else {
            Write-Host "$displayName : $value" -ForegroundColor 'Red'
          }
        } 'Enabled' {
          if (-not $value) {
            Write-Host "$displayName : $value" -ForegroundColor 'Red'
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
            Write-Host "$displayName : $value" -ForegroundColor 'Red'
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
            Write-Host "$displayName : $value" -ForegroundColor 'Green'
          }
        } 'MobilePhone' {
          if ("$value".Trim()) {
            Write-Host "$displayName : $value" -ForegroundColor 'Green'
          }
        } 'Modified' {
          Write-Host "$displayName : $date ($timeSince)"
        } 'PasswordLastSet' {
          if ($value) {
            if ($diff.Days -ge 90) {
              Write-Host "$displayName : $date ($timeSince)" `
                  -ForegroundColor 'Red'
            } else {
              Write-Host "$displayName : $date ($timeSince)" `
                  -ForegroundColor 'Green'
            }
          } else {
            Write-Host "$displayName : change at next sign on" `
                -ForegroundColor 'Yellow'
          }
        } 'ProfilePath' {
          if ($value) {
            Write-Host "$displayName : $value"
          }
        } 'otherHomePhone' {
          if ("$value".Trim()) {
            Write-Host "$displayName : $value" -ForegroundColor 'Green'
          }
        } 'otherMobile' {
          if ("$value".Trim()) {
            Write-Host "$displayName : $value" -ForegroundColor 'Green'
          }
        } 'telephoneNumber' {
          if ("$value".Trim()) {
            Write-Host "$displayName : $value" -ForegroundColor 'Green'
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
    Write-Host ''

    if ($DisableActions) {
      [Console]::CursorVisible = $true
      break rewrite
    }

    # Displays actions menu.
    Write-Host '# ACTIONS #'
    $selection = Display-VerticalMenu -Options $actions
    Write-Host ''
    if ($selection -eq $null) {
      $selection = 'End'
    } else {
      $selection = $actions[$selection]
    }

    # Executes selection.
    switch ($selection) {
      'End' {
        Write-Host ''
        break rewrite
      } 'Reload' {
        Get-User -Usernames $user.SamAccountName
        break rewrite
      } 'Summarize' {
        $summary = Get-UserSummary -Usernames $user.SamAccountName
        Set-Clipboard $summary
        Write-Host ''
        Write-Host 'Copied...'
        Write-Host $summary -ForegroundColor 'green'
        Write-Host ''
        break rewrite
      } 'Return to search' {
        Write-Host ''
        $group = Select-ObjectFromTable `
            -Objects (Search-Objects @searchArguments) `
            -Properties $selectProperties
        Get-User $group.SamAccountName
        break rewrite
      } 'Unlock' {
        Unlock-ADAccount -Identity $user.SamAccountName
        Write-Host ''
        [Console]::CursorVisible = $false
        break rewrite
      } 'Unlock (copy ticket)' {
        Write-Host ''
        Copy-AccountUnlockTicket -Type 'domain' -Username $user.SamAccountName
        [Console]::CursorVisible = $false
        break rewrite
      } 'Reset password' {
        Write-Host ''
        Reset-Password -Users $user.SamAccountName
        Unlock-ADAccount -Identity $user.SamAccountName
        Write-Host ''
        [Console]::CursorVisible = $false
        break rewrite
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
                break rewrite
              }
            }
          }
          $table.Rows.Add($row)
        }
        $table | Format-Table -Wrap
        [Console]::CursorTop = [Console]::CursorTop - 1
        :groupActions while ($true) {
          Write-Host '# ACTIONS #'
          $selection = $null
          $selection = Read-Host "[0] Return  [1-$i] Make selection"
          if (-not $selection) {
            Get-User $user.SamAccountName
            [Console]::CursorVisible = $false
            break rewrite
          }
          try {
            $selection = [int]$selection
          } catch {
            Write-Error "Invalid selection. Expected a number 0-$i."
            continue groupActions
          }
          if ($selection -eq 0) {
            Get-User $user.SamAccountName
            break rewrite
          } elseif (($selection -lt 0) -or ($selection -gt $i)) {
            Write-Error "Invalid selection. Expected a number 0-$i."
            continue groupActions
          } else {
            Get-Group -Names ($table.Rows[$selection - 1]['NAME'])
            break rewrite
          }
        }
        break rewrite
      } 'Add groups' {
        Write-Host ''
        Write-Host ('***You may add multiple groups by separating them with ' `
            + "a comma.***`n")
        $groups = ((Read-Host 'Groups to add') -split ',').Trim()
        foreach ($group in $groups) {
          if ($ILLEGAL_GROUPS -contains $group) {
            Write-Error ("Skipping ""$group"". Modifying this group is " `
                + 'restricted by Data Security.')
          } else {
            Add-ADGroupMember -Identity $group -Members $user
          }
        }
        Write-Host ''
        break rewrite
      } 'Remove groups' {
        Write-Host ''
        Write-Host ('***You may remove multiple groups by separating them with ' `
            + "a comma.***`n")
        $groups = ((Read-Host 'Groups to remove') -split ',').Trim()
        foreach ($group in $groups) {
          if ($ILLEGAL_GROUPS -contains $group) {
            Write-Error ("Skipping ""$group"". Modifying this group is " `
                + 'restricted by Data Security.')
          } else {
            Remove-ADGroupMember -Identity $group -Members $user -Confirm:$false
          }
        }
        Write-Host ''
        [Console]::CursorVisible = $false
        break rewrite
      } 'Search manager' {
        Write-Host ''
        Get-User -Names (($user.Manager) -split ',')[0].Substring(3)
        break rewrite
      } 'Send email' {
        Start-Process "mailto:$($user.EmailAddress)"
        Write-Host ''
        break rewrite
      } default {
        Write-Host ''
        Write-Error 'Unrecognized action.'
        break rewrite
      }
    }
  }
  [Console]::CursorVisible = $true
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

function Open-CopypastesDirectory {
  <#
    .SYNOPSIS
      Opens the directory containing the copy-paste documents that the copy-paste variables are built from.
  #>

  $path = [System.IO.Path]::Combine($PSScriptRoot, 'copypastes')
  Invoke-Item $path
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
      Generates various global string variables in PowerShell representing copy-pastes built using the copy-paste documents (determines the name and content of the variable) (see Open-CopypastesDirectory) and variable substitution (see Insert-Variables.ps1).
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
        -Value (Insert-Variables $contents[$i]) `
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
