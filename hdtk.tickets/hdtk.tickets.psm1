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
      - The parameter type should immediately preceed the parameter name with
        a space separating the type and the name.
      - Switch parameters should be included at the end of the block.
#>

function Copy-AccountUnlockTicket {
  # [TODO] Write manual.

  [Alias('unlock')]
  [CmdletBinding()]

  param (
    [Parameter(
      HelpMessage='Enter the platform that this account is on',
      Mandatory=$true,
      Position=0
    )]
    <# ArgumentCompletions attribute not yet supported by this version of 
      PowerShell. #>
    <# 
      [ArgumentCompletions(
        'domain'
      )]
    #>
    [string] $Type,

    [string] $Phone,

    [Alias('user')]
    [string] $Username,

    [Alias('nosubject')]
    [switch] $DisableSubjectCopy
  )

  $body = "unlock account ($Type)"
  $subject = "Customer requested that their account ($Type) be unlocked."

  if ($Phone) {
    $body += "`n`nCalled from: $Phone"
  }

  switch ($Type) {
    'domain' {
      $user = Get-ADUser $Username -Properties 'AccountLockoutTime'
      if ($Username -and $user.LockedOut) {
        $date = ($user.AccountLockoutTime).ToString("yyyy-MM-dd")
        $time = ($user.AccountLockoutTime).ToString("HH:mm:ss")
        $body += "`n`nAccount locked on $date at $time."
      }
    }
  }

  Set-Clipboard -Value $body
  Write-Host ''
  Write-Host 'Copied...' -ForegroundColor 'green'
  Write-Host $body
  Write-Host ''

  if (-not $DisableSubjectCopy) {
    Read-Host 'Press <enter> to copy the ticket subject'
    Set-Clipboard -Value $subject
    Write-Host 'Copied...' -ForegroundColor 'green'
    Write-Host $subject
    Write-Host ''
  }
}

function Copy-CallerIdTicket {
  <#
    .SYNOPSIS
      Builds and copies the body of a ticket for a misconfigured caller ID.

      ALIAS: callerid

    .PARAMETER Phone
      Represents the phone number with the misconfigured caller ID.

    .PARAMETER Location
      Represents an array of two strings: (1) the first being the current, incorrect location listed in the caller ID; and (2) the second being the correct location for the caller ID.

    .PARAMETER Name
      Represents an array of two strings: (1) the first being the current, incorrect name listed in the caller ID; and (2) the second being the correct name for the caller ID.

    .PARAMETER DisableSubjectCopy
      Signifies that the function should end after copying the ticket body, rather than also copying the ticket subject after waiting for you to press <enter>.

      ALIAS: nosubject

    .EXAMPLE
      Copy-CallerIdTicket '+1 012 345 6789' -Location 'East Tower', 'Central Tower'

      Copies the following ticket body to your clipboard:

      > phone number: +1 012 345 6789
      > 
      > * Caller ID says that it is at "East Tower" when it should say that it is at "Central Tower".

      Then writes "Copied..." plus the ticket body to the command-line interface.

      Then waits for you to hit <enter> before copying the ticket subject to your clipboard and writing "Copied..." plus the ticket subject to the command-line interface.

    .EXAMPLE
      Copy-CallerIdTicket '+1 012 345 6789' -Location 'East Tower', 'Central Tower' -DisableSubjectCopy

      Copies the following ticket body to your clipboard:

      > phone number: +1 012 345 6789
      > 
      > * Caller ID says that it is at "East Tower" when it should say that it is at "Central Tower".

      Then writes "Copied..." plus the ticket body to the command-line interface.

    .EXAMPLE
      Copy-CallerIdTicket '+1 012 345 6789' -Name 'John Doe', 'Jane Doe'

      Copies the following ticket body to your clipboard:

      > phone number: +1 012 345 6789
      > 
      > * Caller ID says that it belongs to "John Doe" when it should say that it belongs to "Jane Doe".

      Then writes "Copied..." plus the ticket body to the command-line interface.

      Then waits for you to hit <enter> before copying the ticket subject to your clipboard and writing "Copied..." plus the ticket subject to the command-line interface.
  #>

  [Alias('callerid')]
  [CmdletBinding()]

  param (
    [Parameter(
      HelpMessage='Enter the phone number with the misconfigured caller ID.',
      Mandatory=$true,
      Position=0
    )]
    [string] $Phone,

    [string[]] $Location,

    [string[]] $Name,

    [Alias('nosubject')]
    [switch] $DisableSubjectCopy
  )

  $body = "phone number: $Phone`n"
  $subject = 'misconfigured caller ID'

  if ($Location) {
    if ($Location.Count -ne 2) {
      Write-Error 'The Location parameter must provide two--and only two--values: (1) the current, incorrect building listed in the caller ID; and (2) the correct building.'
      return
    }
    $body += "`n* Caller ID says that it is at ""$($Location[0])"" when it " `
        + "should say that it is at ""$($Location[1])""."
  }
  if ($Name) {
    if ($Name.Count -ne 2) {
      Write-Error 'The Name parameter must provide two--and only two--values: (1) the current, incorrect name listed in the caller ID; and (2) the correct name.'
      return
    }
    $body += "`n* Caller ID says that it belongs to ""$($Name[0])"" when it " `
        + "should say that it belongs to ""$($Name[1])""."
  }

  Set-Clipboard -Value $body
  Write-Host ''
  Write-Host 'Copied...' -ForegroundColor 'green'
  Write-Host $body
  Write-Host ''

  if (-not $DisableSubjectCopy) {
    Read-Host 'Press <enter> to copy the ticket subject'
    Set-Clipboard -Value $subject
    Write-Host 'Copied...' -ForegroundColor 'green'
    Write-Host $subject
    Write-Host ''
  }
}

function Copy-ComputerNameTicket {
  <#
    .SYNOPSIS
      Builds and copies the body of a ticket for a misconfigured computer name.

      ALIAS: pcname

    .PARAMETER CurrentName
      Represents the current, incorrect name set for the computer.

      ALIAS: name

    .PARAMETER AssetId
      Represents the ID on the asset tag of the computer.

      ALIAS: id

    .PARAMETER DisableSubjectCopy
      Signifies that the function should end after copying the ticket body, rather than also copying the ticket subject after waiting for you to press <enter>.

      ALIAS: nosubject

    .PARAMETER MissingTag
      Signifies that the customer was unable to locate the asset tag of the computer. Sets the AssetId parameter to "Customer was unable to locate an asset tag.".

      ALIAS: notag

    .EXAMPLE
      Copy-ComputerNameTicket INCORRECT_NAME ASSET_ID

      Copies the following ticket body to your clipboard:

      > current name
      > : INCORRECT_NAME
      > 
      > asset ID
      > : ASSET_ID

      Then writes "Copied..." plus the ticket body to the command-line interface.

      Then waits for you to hit <enter> before copying the ticket subject to your clipboard and writing "Copied..." plus the ticket subject to the command-line interface.

    .EXAMPLE
      Copy-ComputerNameTicket INCORRECT_NAME ASSET_ID -DisableSubjectCopy

      Copies the following ticket body to your clipboard:

      > current name
      > : INCORRECT_NAME
      > 
      > asset ID
      > : ASSET_ID

      Then writes "Copied..." plus the ticket body to the command-line interface.
  #>

  [Alias('pcname')]
  [CmdletBinding()]

  param (
    [Alias('name')]
    [Parameter(
      HelpMessage='Enter the current, incorrect name of the computer.',
      Mandatory=$true,
      Position=0
    )]
    [string] $CurrentName,

    [Alias('id')]
    [Parameter(Position=1)]
    [string] $AssetId,

    [Alias('nosubject')]
    [switch] $DisableSubjectCopy,

    [Alias('notag')]
    [switch] $MissingTag
  )

  if ($MissingTag) {
    if ($AssetId) {
      Write-Error ('The MissingTag switch cannot be used with the AssetId ' `
          + 'parameter.')
      return
    }
    $AssetId = 'Customer was unable to locate an asset tag.'
  } elseif (-not $AssetId) {
    Write-Error ('A string must be provided to the AssetId parameter unless ' `
        + 'the MissingTag switch is used.')
    return
  }

  $body = "current name`n: $CurrentName`n`nasset tag`n: $AssetId"
  $subject = 'misconfigured computer name'

  Set-Clipboard -Value $body
  Write-Host ''
  Write-Host 'Copied...' -ForegroundColor 'green'
  Write-Host $body
  Write-Host ''

  if (-not $DisableSubjectCopy) {
    Read-Host 'Press <enter> to copy the ticket subject'
    Set-Clipboard -Value $subject
    Write-Host 'Copied...' -ForegroundColor 'green'
    Write-Host $subject
    Write-Host ''
  }
}

function Copy-ConnectPrinterTicket {
  <#
    .SYNOPSIS
      Builds and copies the body of a ticket for connecting one or more computers to one or more printers.

      ALIAS: addprinter

    .PARAMETER Computer
      Represents the name(s) of the computer(s) being connected to the printer(s).

      ALIAS: pc

    .PARAMETER Printer
      Represents the path(s) of the printer(s) being connected to.

    .PARAMETER DisableSubjectCopy
      Signifies that the function should not offer to copy the ticket subject after waiting for you to press <enter>.

      ALIAS: nosubject

    .PARAMETER DisableFulfillmentCopy
      Signifies that the function should not offer to copy the ticket fulfillment comment after waiting for you to press <enter>.

      ALIAS: nocomment

    .EXAMPLE
      Copy-ConnectPrinterTicket COMPUTER_NAME PRINTER_NAME

      Copies the following ticket body to your clipboard:

      > Customer requested to have a computer ("COMPUTER_NAME") connected to a printer ("PRINTER_NAME").

      Then writes "Copied..." plus the ticket body to the command-line interface.

      Then waits for you to hit <enter> before copying the ticket subject to your clipboard and writing "Copied..." plus the ticket subject to the command-line interface.

      Then waits for you to hit <enter> before copying the ticket fulfillment comment to your clipboard and writing "Copied..." plus the ticket fulfillment comment to the command-line interface.

    .EXAMPLE
      Copy-ConnectPrinterTicket 'COMPUTER_0', 'COMPUTER_1' 'PRINTER_0', 'PRINTER_1'

      Copies the following ticket body to your clipboard:

      > Customer requested to have computers ("COMPUTER_0", "COMPUTER_1") connected to some printers ("PRINTER_0", "PRINTER_1").

      Then writes "Copied..." plus the ticket body to the command-line interface.

      Then waits for you to hit <enter> before copying the ticket subject to your clipboard and writing "Copied..." plus the ticket subject to the command-line interface.

      Then waits for you to hit <enter> before copying the ticket fulfillment comment to your clipboard and writing "Copied..." plus the ticket fulfillment comment to the command-line interface.

    .EXAMPLE
      Copy-ConnectPrinterTicket COMPUTER_NAME PRINTER_NAME -DisableSubjectCopy

      Copies the following ticket body to your clipboard:

      > Customer requested to have a computer ("COMPUTER_NAME") connected to a printer ("PRINTER_NAME").

      Then writes "Copied..." plus the ticket body to the command-line interface.

      Then waits for you to hit <enter> before copying the ticket fulfillment comment to your clipboard and writing "Copied..." plus the ticket fulfillment comment to the command-line interface.

    .EXAMPLE
      Copy-ConnectPrinterTicket COMPUTER_NAME PRINTER_NAME -DisableFulfillmentCopy

      Copies the following ticket body to your clipboard:

      > Customer requested to have a computer ("COMPUTER_NAME") connected to a printer ("PRINTER_NAME").

      Then writes "Copied..." plus the ticket body to the command-line interface.

      Then waits for you to hit <enter> before copying the ticket subject to your clipboard and writing "Copied..." plus the ticket subject to the command-line interface.

    .EXAMPLE
      Copy-ConnectPrinterTicket COMPUTER_NAME PRINTER_NAME -DisableSubjectCopy -DisableFulfillmentCopy

      Copies the following ticket body to your clipboard:

      > Customer requested to have a computer ("COMPUTER_NAME") connected to a printer ("PRINTER_NAME").

      Then writes "Copied..." plus the ticket body to the command-line interface.
  #>

  [Alias('addprinter')]
  [CmdletBinding()]

  param (
    [Alias('pc')]
    [Parameter(
      HelpMessage=('Enter the name of the computer that you are connecting'),
      Mandatory=$true,
      Position=0
    )]
    [string[]] $Computer,

    [Parameter(
      HelpMessage='Enter the path of the printer that you are connecting',
      Mandatory=$true,
      Position=1
    )]
    [string[]] $Printer,

    [Alias('nosubject')]
    [switch] $DisableSubjectCopy,

    [Alias('nocomment')]
    [switch] $DisableFulfillmentCopy
  )

  $computers = '"{0}"' -f ($Computer -join '", "')
  $printers = '"{0}"' -f ($Printer -join '", "')

  $body = 'Customer requested to have '
  $subject = 'connect computer(s) to printer(s)'
  $fulfillment = 'Connected '

  if ($Computer.Count -gt 1) {
    $body += "some computers ($computers) and "
    $fulfillment += "the computers ($computers) to "
  } else {
    $body += "a computer ($computers) and "
    $fulfillment += "the computer ($computers) to "
  }
  if ($Printer.Count -gt 1) {
    $body += "some printers ($printers) connected."
    $fulfillment += "the printers ($printers)."
  } else {
    $body += "a printer ($printers) connected."
    $fulfillment += "the printer ($printers)."
  }

  Set-Clipboard -Value $body
  Write-Host ''
  Write-Host 'Copied...' -ForegroundColor 'green'
  Write-Host $body
  Write-Host ''

  if (-not $DisableSubjectCopy) {
    Read-Host 'Press <enter> to copy the ticket subject'
    Set-Clipboard -Value $subject
    Write-Host 'Copied...' -ForegroundColor 'green'
    Write-Host $subject
    Write-Host ''
  }

  if (-not $DisableFulfillmentCopy) {
    Read-Host 'Press <enter> to copy the ticket fulfillment comment'
    Set-Clipboard -Value $fulfillment
    Write-Host 'Copied...' -ForegroundColor 'green'
    Write-Host $fulfillment
    Write-Host ''
  }
}

function Copy-MapDriveTicket {
  <#
    .SYNOPSIS
      Builds and copies the body and subject of a ticket for mapping network drive(s).

      ALIAS: mapdrive

    .PARAMETER Computer
      Represents the name(s) of the computer(s) having drives mapped.

      ALIAS: pc

    .PARAMETER Path
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
      Copy-MapDriveTicket -Computer COMPUTER_1 -Path PATH_1

      Copies the following ticket body to your clipboard:

      > Customer requested to have a drive ("PATH_1") mapped for a computer ("COMPUTER_1").

      Writes "Copied..." plus the ticket body to the command-line interface.

      Then waits for you to hit <enter> before copying the ticket subject ("map drive ("PATH_1")") to your clipboard and writing "Copied..." plus the ticket subject to the command-line interface.

    .EXAMPLE
      Copy-MapDriveTicket COMPUTER_1 PATH_1

      Copies the following ticket body to your clipboard:

      > Customer requested to have a drive ("PATH_1") mapped for a computer ("COMPUTER_1").

      Writes "Copied..." plus the ticket body to the command-line interface.

      Then waits for you to hit <enter> before copying the ticket subject ("map drive ("path_1")") to your clipboard and writing "Copied..." plus the ticket subject to the command-line interface.

    .EXAMPLE
      Copy-MapDriveTicket 'COMPUTER_1', 'COMPUTER_2' 'PATH_1', 'PATH_2'

      Copies the following ticket body to your clipboard:

      > Customer requested to have drives ("PATH_1", "PATH_2") mapped for some computers ("COMPUTER_1", "COMPUTER_2").

      Writes "Copied..." plus the ticket body to the command-line interface.

      Then waits for you to hit <enter> before copying the ticket subject ("map drives") to your clipboard and writing "Copied..." plus the ticket subject to the command-line interface.

    .EXAMPLE
      Copy-MapDriveTicket COMPUTER_1 PATH_1 -DisableSubjectCopy

      Copies the following ticket body to your clipboard:

      > Customer requested to have a drive ("PATH_1") mapped for a computer ("COMPUTER_1").

      Writes "Copied..." plus the ticket body to the command-line interface.

      Then waits for you to hit <enter> before copying the ticket fulfillment comment to your clipboard and writing "Copied..." plus the ticket fulfillment comment to the command-line interface.
  #>

  [Alias('mapdrive')]
  [CmdletBinding()]

  param (
    [Alias('pc')]
    [Parameter(
      HelpMessage=('Enter the computer name'),
      Mandatory=$true,
      Position=0
    )]
    [string[]] $Computer,

    [Parameter(
      HelpMessage=('Enter the network path'),
      Mandatory=$true,
      Position=1
    )]
    [string[]] $Path,

    [Alias('nosubject')]
    [switch] $DisableSubjectCopy,

    [Alias('nocomment')]
    [switch] $DisableFulfillmentCopy,

    [switch] $Remap
  )

  $computers = '"{0}"' -f ($Computer -join '", "')
  $paths = '"{0}"' -f ($Path -join '", "')

  $body = ''
  $subject = ''
  $fulfillment = ''

  if ($Remap) {
    $subject = 're-map drive(s)'
  } else {
    $subject = 'map drive(s)'
  }

  if ($Path.Count -gt 1) {
    $body = "Customer requested to have drives ($paths) "
    if ($Remap) {
      $fulfillment = "Re-mapped the paths ($paths) to some drives for the "
    } else {
      $fulfillment = "Mapped the paths ($paths) to some drives for the "
    }
  } else {
    $body = "Customer requested to have a drive ($paths) "
    if ($Remap) {
      $fulfillment = "Re-mapped the path ($paths) to a drive for the "
    } else {
      $fulfillment = "Mapped the path ($paths) to a drive for the "
    }
  }
  if ($Computer.Count -gt 1) {
    if ($Remap) {
      $body += "re-mapped for some computers ($computers)."
    } else {
      $body += "mapped for some computers ($computers)."
    }
    $fulfillment += "computers ($computers)."
  } else {
    if ($Remap) {
      $body += "re-mapped for a computer ($computers)."
    } else {
      $body += "mapped for a computer ($computers)."
    }
    $fulfillment += "computer ($computers)."
  }

  Set-Clipboard -Value $body
  Write-Host ''
  Write-Host 'Copied...' -ForegroundColor 'green'
  Write-Host $body
  Write-Host ''

  if (-not $DisableSubjectCopy) {
    Read-Host 'Press <enter> to copy the ticket subject'
    Set-Clipboard -Value $subject
    Write-Host 'Copied...' -ForegroundColor 'green'
    Write-Host $subject
    Write-Host ''
  }

  if (-not $DisableFulfillmentCopy) {
    Read-Host 'Press <enter> to copy the ticket fulfillment comment'
    Set-Clipboard -Value $fulfillment
    Write-Host 'Copied...' -ForegroundColor 'green'
    Write-Host $fulfillment
    Write-Host ''
  }
}

function Copy-OutlookProfileResetTicket {
  # [TODO] Write manual.

  [Alias('outlookreset')]
  [CmdletBinding()]

  $date = Get-Date -Format 'yyyy-MM-dd'
  $body = "Customer was unable to launch Outlook and would receive an " `
      + "error about their data file being corrupted.`n`nReset their " `
      + "Outlook profile by going to Control Panel > Mail (Microsoft " `
      + "Outlook) > Profiles > Show Profiles... > Add... > ""Profile " `
      + "Name:"":""$date"" > OK > ""OK""/""Apply"". Issue resolved."
  Set-Clipboard -Value $body
  Write-Host 'Copied...'
  Write-Host $body -ForegroundColor 'green'
}

function Copy-VoicemailPinResetTicket {
  <#
    .SYNOPSIS
      Builds and copies the body of a ticket for a request to reset the voicemail PIN for a phone number.

      ALIAS: voicemail

    .PARAMETER Phone
      Represents the phone number associated with the voicemail whose PIN you are requesting be reset.

    .PARAMETER DisableSubjectCopy
      Signifies that the function should not offer to copy the ticket subject after waiting for you to press <enter>.

      ALIAS: nosubject

    .EXAMPLE
      Copy-VoicemailPinResetTicket '+1 012 345 6789'

      Copies the following ticket body to your clipboard:

      > Customer requested a voicemail PIN reset.
      > 
      > phone number: +1 012 345 6789

      Writes "Copied..." plus the ticket body to the command-line interface.

      Then waits for you to hit <enter> before copying the ticket subject ("voicemail PIN reset") to your clipboard and writing "Copied..." plus the ticket subject to the command-line interface.
  #>

  [Alias('voicemail')]
  [CmdletBinding()]

  param (
    [Parameter(
      HelpMessage=('Enter the phone number associated with the voicemail ' `
          + 'whose PIN you are requesting to have reset.'),
      Mandatory=$true,
      Position=0,
      ValueFromPipeline=$true
    )]
    [string] $Phone,

    [Alias('nosubject')]
    [switch] $DisableSubjectCopy
  )

  $body = 'Customer requested a voicemail PIN reset.' `
      + "`n`nphone number: $Phone"
  $subject = 'voicemail PIN reset'

  Set-Clipboard -Value $body
  Write-Host ''
  Write-Host 'Copied...' -ForegroundColor 'green'
  Write-Host $body
  Write-Host ''

  if (-not $DisableSubjectCopy) {
    Read-Host 'Press <enter> to copy the ticket subject'
    Set-Clipboard -Value $subject
    Write-Host 'Copied...' -ForegroundColor 'green'
    Write-Host $subject
    Write-Host ''
  }
}
