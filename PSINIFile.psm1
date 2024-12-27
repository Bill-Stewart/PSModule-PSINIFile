#requires -version 3

# PSINIFile.psm1
# Written by Bill Stewart (bstewart AT iname.com)
#
# This is a PowerShell module that functions as a wrapper for the Windows API
# functions GetPrivateProfileString and WritePrivateProfileString.
#
# Inspired by https://gist.github.com/mklement0/006c2352ddae7bb05693be028240f5b6
# with some improvements:
# * Better code readability (IMO)
# * Improved error handling
# * Functionality split into separate function cmdlets
# * Prevent use of ']' character in sections and '=' character in keys
#
# Conditional syntax: (<f>,<t>)[<bool>]
# Outputs <f> if <bool> is $false or <t> otherwise
#
# API notes:
# * The GetPrivateProfileString API import uses '[Out] byte []' because
#   the System.Text.StringBuilder class doesn't support embedded nulls when
#   marshaling strings.
# * Apparently, assigning '[NullString]::Value' (which marshals to NULL/0)
#   directly to a typed parameter variable breaks things. To avoid problems,
#   the functions that call the Win32 APIs directly use the conditional syntax
#   (noted above) to pass '[NullString]::Value' to the API function directly
#   if needed.

Add-Type -TypeDefinition @'
namespace CE39BBC0482843A4A87B235B550FC993 {
  using System.Runtime.InteropServices;

  public static class Kernel32 {
    // [CE39BBC0482843A4A87B235B550FC993.Kernel32]::GetPrivateProfileString()
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern uint GetPrivateProfileString(string lpAppName,
      string lpKeyName, string lpDefault, [Out] byte[] lpBuffer, uint nSize,
      string lpFileName);

    // [CE39BBC0482843A4A87B235B550FC993.Kernel32]::WritePrivateProfileString()
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern bool WritePrivateProfileString(string lpAppName,
      string lpKeyName, string lpString, string lpFileName);
  }
}
'@

# Low-level wrapper for GetPrivateProfileString function
function GetIniValue {
  [CmdletBinding()]
  param(
    [parameter(Mandatory)]
    [String]
    $path,

    [String]
    $section,

    [String]
    $key,

    [String]
    $default
  )
  $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path)
  if ( -not (Test-Path -LiteralPath $fullPath) ) {
    throw (New-Object Management.Automation.ErrorRecord(
      ([IO.FileNotFoundException] "The file '$fullPath' was not found."),
      (Get-Variable Myinvocation -Scope 1).Value.MyCommand.Name,
      (([Management.Automation.ErrorCategory]::ObjectNotFound)),$fullPath))
  }
  $sectionSpecified = $PSBoundParameters.ContainsKey("section")
  $keySpecified = $PSBoundParameters.ContainsKey("key")
  $defaultSpecified = $PSBoundParameters.ContainsKey("default")
  # Per the GetPrivateProfileString API documentation:
  # * If neither lpAppName nor lpKeyName is NULL and the supplied destination
  #   buffer is too small to hold the requested string, the string is truncated
  #   and followed by a null character, and the return value is equal to nSize
  #   minus one.
  # * If either lpAppName or lpKeyName is NULL and the supplied destination
  #   buffer is too small to hold all the strings, the last string is truncated
  #   and followed by two null characters. In this case, the return value is
  #   equal to nSize minus two.
  # * In the event the initialization file specified by lpFileName is not
  #   found [a condition we've already eliminated above], or contains invalid
  #   values, this function will set errorno with a value of '0x2' (File Not
  #   Found). To retrieve extended error information, call GetLastError.
  # If GetLastError returns 2, we interpret this as simply "section or key name
  # not found" from the PowerShell perspective and return nothing.
  $charsDiff = (1,2)[(-not $sectionSpecified) -or (-not $keySpecified)]
  $charSize = [Text.Encoding]::Unicode.GetByteCount([Char] 0)
  $nSize = 0
  do {
    $nSize += 2KB
    $buffer = New-Object Byte[] ($nSize * $charSize)
    $charsCopied = [CE39BBC0482843A4A87B235B550FC993.Kernel32]::GetPrivateProfileString(
      ([NullString]::Value,$section)[$sectionSpecified],  # lpAppName
      ([NullString]::Value,$key)[$keySpecified],          # lpKeyName
      ([NullString]::Value,$default)[$defaultSpecified],  # lpDefault
      $buffer,                                            # lpbuffer
      $nSize,                                             # nSize
      $fullPath)                                          # lpFileName
    $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
    if ( ($lastError -ne 0) -and ($lastError -ne 2) ) {
      $message = "Unable to read from '$path' due to the following error: '{0}.'" -f
        ([ComponentModel.Win32Exception] $lastError).Message
      throw (New-Object ComponentModel.Win32Exception($lastError,$message))
    }
  }
  until ( ($charsCopied -eq 0) -or ($charsCopied -ne $nSize - $charsDiff) )
  # No characters copied to buffer; nothing to return
  if ( $charsCopied -eq 0 ) {
    return
  }
  # If -section or -key were omitted, split the string by embedded nulls;
  # otherwise, return the entire string (nothing to split)
  [Text.Encoding]::Unicode.GetString($buffer,0,
    (($charsCopied - --$charsDiff) * $charSize)) -split ([Char] 0)
}

# Wrapper for GetIniValue
function Get-IniValue {
  <#
  .SYNOPSIS
  Gets a value from an INI file.

  .DESCRIPTION
  Gets a value from an INI file. INI files are text files with section names surrounded by square braces ('[' and ']') and key=value pairs.

  .PARAMETER Path
  The filename of the INI file to read.

  .PARAMETER Section
  Specifies the section name. The section name cannot contain the ']' character.

  .PARAMETER Key
  Specifies the key name. The key name cannot contain the '=' character.

  .PARAMETER Default
  Specifies a default value to return if the key doesn't exist. The default is to return nothing.

  .INPUTS
  None.

  .OUTPUTS
  If value exists: System.String
  If value does not exist: Nothing, or value specified by -Default
  #>
  [CmdletBinding()]
  param(
    [parameter(Position = 0,Mandatory)]
    [String]
    $Path,

    [parameter(Position = 1,Mandatory)]
    [String]
    $Section,

    [parameter(Position = 2,Mandatory)]
    [String]
    $Key,

    [parameter(Position = 3)]
    [String]
    $Default
  )
  if ( $Section.IndexOf("]") -ne -1 ) {
    Write-Error "The section name cannot contain the ']' character." -Category InvalidArgument
    return
  }
  if ( $Key.IndexOf("=") -ne -1 ) {
    Write-Error "The key name cannot contain the '=' character." -Category InvalidArgument
    return
  }
  GetIniValue $Path $Section $Key $Default
}

# Wrapper for GetIniValue
function Get-IniSection {
  <#
  .SYNOPSIS
  Gets the sections from an INI file. INI files are text files with section names surrounded by square braces ('[' and ']') and key=value pairs.

  .DESCRIPTION
  Gets the sections from an INI file.

  .PARAMETER Path
  The filename of the INI file to read.

  .INPUTS
  None.

  .OUTPUTS
  If no sections exist: Nothing
  If 1 section exists: System.String
  If 2 or more sections exist: System.String[]
  #>
  [CmdletBinding()]
  param(
    [parameter(Position = 0,Mandatory)]
    [String]
    $Path
  )
  GetIniValue $Path
}

# Wrapper for GetIniValue
function Get-IniKey {
  <#
  .SYNOPSIS
  Gets the keys named in a section of an INI file.

  .DESCRIPTION
  Gets the keys named in a section of an INI file. INI files are text files with section names surrounded by square braces ('[' and ']') and key=value pairs.

  .PARAMETER Path
  The filename of the INI file to read.

  .PARAMETER Section
  Specifies the section name. The section name cannot contain the ']' character.

  .INPUTS
  None.

  .OUTPUTS
  If no keys exist: Nothing
  If 1 key exists: System.String
  If 2 or more keys exist: System.String[]
  #>
  [CmdletBinding()]
  param(
    [parameter(Position = 0,Mandatory)]
    [String]
    $Path,

    [parameter(Position = 1,Mandatory)]
    [String]
    $Section
  )
  if ( $Section.IndexOf("]") -ne -1 ) {
    Write-Error "The section name cannot contain the ']' character." -Category InvalidArgument
    return
  }
  GetIniValue $Path $Section
}

# Wrapper for Get-IniSection, Get-IniKey, and Get-IniValue
function Get-IniFile {
  <#
  .SYNOPSIS
  Gets the content of the specified INI file as a list of objects.

  .DESCRIPTION
  Gets the content of the specified INI file as a list of objects. INI files are text files with section names surrounded by square braces ('[' and ']') and key=value pairs.

  .PARAMETER Path
  The filename of the INI file to read.

  .INPUTS
  None.

  .OUTPUTS
  Objects with the following properties:
  Section - System.String
  Key - System.String
  Value - System.String
  #>
  param(
    [parameter(Mandatory)]
    [String]
    $Path
  )
  foreach ( $Section in (Get-IniSection $Path) ) {
    foreach ( $Key in (Get-IniKey $Path $Section) ) {
      [PSCustomObject] @{
        "Section" = $Section
        "Key"     = $Key
        "Value"   = Get-IniValue $Path $Section $Key
      }
    }
  }
}

# Low-level wrapper for WritePrivateProfileString function
function SetIniValue {
  [CmdletBinding()]
  param(
    [parameter(Mandatory)]
    [String]
    $path,

    [parameter(Mandatory)]
    [String]
    $section,

    [String]
    $key,

    [String]
    $value
  )
  $keySpecified = $PSBoundParameters.ContainsKey("key")
  $valueSpecified = $PSBoundParameters.ContainsKey("value")
  $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path)
  if ( (-not $keySpecified) -or (-not $valueSpecified) ) {
    if ( -not (Test-Path -LiteralPath $fullPath) ) {
      throw (New-Object Management.Automation.ErrorRecord(
        ([IO.FileNotFoundException] "The file '$fullPath' was not found."),
        (Get-Variable Myinvocation -Scope 1).Value.MyCommand.Name,
        (([Management.Automation.ErrorCategory]::ObjectNotFound)),$fullPath))
    }
  }
  $ok = [CE39BBC0482843A4A87B235B550FC993.Kernel32]::WritePrivateProfileString(
    $section,                                       # lpAppName
    ([NullString]::Value,$key)[$keySpecified],      # lpKeyName
    ([NullString]::Value,$value)[$valueSpecified],  # lpString
    $fullPath)                                      # lpFileName
  $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
  if ( (-not $ok) -or ($lastError -ne 0) ) {
    $message = "Unable to update '$fullPath' due to the following error: '{0}.'" -f
      ([ComponentModel.Win32Exception] $lastError).Message
    throw (New-Object ComponentModel.Win32Exception($lastError,$message))
  }
}

# Wrapper for SetIniValue
function Set-IniValue {
  <#
  .SYNOPSIS
  Sets a value in an INI file.

  .DESCRIPTION
  Sets a value in an INI file. INI files are text files with section names surrounded by square braces ('[' and ']') and key=value pairs.

  .PARAMETER Path
  The filename of the INI file to update.

  .PARAMETER Section
  Specifies the section name. The section name cannot contain the ']' character.

  .PARAMETER Key
  Specifies the key name. The key name cannot contain the '=' character.

  .PARAMETER Value
  Specifies the value for the key.

  .NOTES
  If the INI file does not exist, it will be written in ANSI character format. If you want to create a new Unicode INI file, you must manually create the file first using Unicode format.

  .INPUTS
  None.

  .OUTPUTS
  None.
  #>
  [CmdletBinding()]
  param(
    [parameter(Position = 0,Mandatory)]
    [String]
    $Path,

    [parameter(Position = 1,Mandatory)]
    [String]
    $Section,

    [parameter(Position = 2,Mandatory)]
    [String]
    $Key,

    [parameter(Position = 3,Mandatory)]
    [String]
    $Value
  )
  if ( $Section.IndexOf("]") -ne -1 ) {
    Write-Error "The section name cannot contain the ']' character." -Category InvalidArgument
    return
  }
  if ( $Key.IndexOf("=") -ne -1 ) {
    Write-Error "The key name cannot contain the '=' character." -Category InvalidArgument
    return
  }
  SetIniValue $Path $Section $Key $Value
}

# Wrapper for SetIniValue
function Remove-IniKey {
  <#
  .SYNOPSIS
  Removes a key and its value from an INI file.

  .DESCRIPTION
  Removes a key and its value from an INI file. INI files are text files with section names surrounded by square braces ('[' and ']') and key=value pairs.

  .PARAMETER Path
  The filename of the INI file to update.

  .PARAMETER Section
  Specifies the section name. The section name cannot contain the ']' character.

  .PARAMETER Key
  Specifies the key name. The key name cannot contain the '=' character.

  .INPUTS
  None.

  .OUTPUTS
  None.
  #>
  [CmdletBinding(SupportsShouldProcess,ConfirmImpact = "High")]
  param(
    [parameter(Mandatory)]
    [String]
    $Path,

    [parameter(Mandatory)]
    [String]
    $Section,

    [parameter(Mandatory)]
    [String]
    $Key
  )
  if ( $Section.IndexOf("]") -ne -1 ) {
    Write-Error "The section name cannot contain the ']' character." -Category InvalidArgument
    return
  }
  if ( $Key.IndexOf("=") -ne -1 ) {
    Write-Error "The key name cannot contain the '=' character." -Category InvalidArgument
    return
  }
  if ( $PSCmdlet.ShouldProcess($Path,"Remove key '$Key' from section '$Section'") ) {
    SetIniValue $Path $Section $Key
  }
}

# Wrapper for SetIniValue
function Remove-IniSection {
  <#
  .SYNOPSIS
  Removes an entire section from an INI file.

  .DESCRIPTION
  Removes an entire section from an INI file. INI files are text files with section names surrounded by square braces ('[' and ']') and key=value pairs.

  .PARAMETER Path
  The filename of the INI file to update.

  .PARAMETER Section
  Specifies the section name. The section name cannot contain the ']' character.

  .INPUTS
  None.

  .OUTPUTS
  None.
  #>
  [CmdletBinding(SupportsShouldProcess,ConfirmImpact = "High")]
  param(
    [parameter(Mandatory)]
    [String]
    $Path,

    [parameter(Mandatory)]
    [String]
    $Section
  )
  if ( $Section.IndexOf("]") -ne -1 ) {
    Write-Error "The section name cannot contain the or ']' character." -Category InvalidArgument
    return
  }
  if ( $PSCmdlet.ShouldProcess($Path,"Remove section '$section'") ) {
    SetIniValue $Path $Section
  }
}
