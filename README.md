# introduction

I originally created these modules in my personal time just to make the work 
me and my co-workers did at an IT help desk more convenient and efficient. 
But, once the project reached a certain size, i felt that i may as well 
publicly release it as an open-source project (with a few, necessary 
modifications).

To view the manuals for each module, use PowerShell's built-in [help system](https://learn.microsoft.com/en-us/powershell/scripting/learn/ps101/02-help-system).

# contributing

If you have any suggestions or contributions that you would like to make 
then please feel free to contact me via [Signal](https://signal.me/#eu/wJ637b5VkqBblVxC12ticHfBFsbgnaVv1OCDIPX8pEZCZ650NP1Jm7pQbYQ+Dxi0).

# modules

A brief description of each module will be provided here, but if you would 
like to view the commands available with each module or the manuals for 
using said commands, please use the `Get-Command` and `Get-Help` commands 
from PowerShell's built-in [help system](https://learn.microsoft.com/en-us/powershell/scripting/learn/ps101/02-help-system).

## hdtk.ad

This module mainly focuses around commands that heavily browse and interact 
with Active Directory. For example, `Get-User` searches Active Directory for 
accounts via various available methods, prints certain account properties 
and allows you to perform certain actions relating to the account (such as 
resetting the password, printing group memberships, adding group 
memberships, et cetera).

## hdtk.tickets

The focus for this module is building and copying the body, subject/summary 
and fulfillment comment for common tickets. For example, 
`Copy-ConnectPrinterTicket` builds and copies the body, subject and 
fulfillment comment for a request to connect one or more computers to one 
or more printers.

# scripting guidelines

Due to the lack of formal PowerShell scripting conventions, i have included 
some general scripting conventions for this project that i will update as 
the project develops.

- "Why have scripting conventions?":
  <https://www.oracle.com/java/technologies/cc-java-programming-language.html>.
- 80-character column limit.
- 2-space indentation with 4-space indentation for continuation lines.
- Write in the [1TB style](https://en.wikipedia.org/wiki/Indentation_style#One_True_Brace).
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
    - The parameter type should immediately preceed the parameter name.
    - Switch parameters should be included at the end of the block.

# license

Please refer to the "LICENSE" file for details.
