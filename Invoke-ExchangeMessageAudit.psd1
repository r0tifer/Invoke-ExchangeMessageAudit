@{
  RootModule = 'Invoke-ExchangeMessageAudit.psm1'
  ModuleVersion = '1.0.0'
  GUID = '30e883d6-a436-4ff0-af9a-d310f673d02d'
  Author = 'Michael Levesque'
  CompanyName = 'Community'
  Copyright = '(c) 2026 r0tifer. Licensed under the MIT License.'
  Description = 'Modular Exchange message tracing and audit workflow for Exchange 2013/2016/2019.'
  PowerShellVersion = '5.1'

  FunctionsToExport = @('Invoke-ExchangeMessageAudit')
  CmdletsToExport = @()
  VariablesToExport = @()
  AliasesToExport = @()

  PrivateData = @{
    PSData = @{
      Tags = @('Exchange', 'MessageTracking', 'Audit')
      ProjectUri = 'https://github.com/r0tifer/Invoke-ExchangeMessageAudit'
      LicenseUri = 'https://github.com/r0tifer/Invoke-ExchangeMessageAudit/blob/main/LICENSE'
      ReleaseNotes = 'Packaged project as a formal PowerShell module with a root manifest and module entry function.'
    }
  }
}
