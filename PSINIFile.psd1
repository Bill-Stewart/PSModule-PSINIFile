@{
  RootModule        = 'PSIniFile.psm1'
  ModuleVersion     = '0.0.0.1'
  GUID              = '14936284-e597-43ca-81ff-a4ad7f61911f'
  Author            = 'Bill Stewart'
  CompanyName       = 'Bill Stewart'
  Copyright         = '(C) 2024 by Bill Stewart'
  Description       = 'Provides a convenient PowerShell cmdlet interface for reading from and writing to text-based .ini files.'
  PowerShellVersion = '3.0'
  FunctionsToExport = @(
    'Get-IniFile'
    'Get-IniKey'
    'Get-IniSection'
    'Get-IniValue'
    'Remove-IniKey'
    'Remove-IniSection'
    'Set-IniValue'
  )
}
