@{
  RootModule = 'Invoke-ExchangeMessageAudit.psm1'
  ModuleVersion = '1.0.0'
  GUID = '30e883d6-a436-4ff0-af9a-d310f673d02d'
  Author = 'r0tifer'
  CompanyName = 'Community'
  Copyright = '(c) r0tifer. All rights reserved.'
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
      ReleaseNotes = 'Packaged project as a formal PowerShell module with a root manifest and module entry function.'
    }
  }
}
