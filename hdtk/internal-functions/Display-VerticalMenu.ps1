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
