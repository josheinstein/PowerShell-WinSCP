@{

    # MODULE
    Description        = "WinSCP Module for PowerShell"
    ModuleVersion      = '1.0'
    GUID               = '796D65D2-CF8A-4DC2-B49C-C7C7A5047749'

    # AUTHOR
    Author             = 'Josh Einstein'
    CompanyName        = 'Josh Einstein'
    Copyright          = 'Copyright 2014, Josh Einstein'

    # REQUIREMENTS
    PowerShellVersion  = '3.0'
    CLRVersion         = '4.0'
    RequiredModules    = @()
    RequiredAssemblies = @(
        'Bin\WinSCPnet.dll'
    )

    # CONTENTS
    #FormatsToProcess   = @('Types\Formats.ps1xml')
    #TypesToProcess     = @('Types\Types.ps1xml')
    ModuleToProcess    = 'WinSCP.psm1'

}