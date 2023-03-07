﻿#
# Module manifest for module 'PowerShellMailUtils'
#
# Generated by: Tommaso Toniolo
#
# Generated on: 01/01/2023
#

@{

# Script module or binary module file associated with this manifest.
# RootModule = ''

# Version number of this module.
ModuleVersion = '1.0.0'

# Supported PSEditions
# CompatiblePSEditions = @()

# ID used to uniquely identify this module
GUID = '610b6d39-b4bf-4f61-b4cb-afc29891941c'

# Author of this module
Author = 'Tommaso Toniolo'

# Company or vendor of this module
CompanyName = 'Tommaso Toniolo - https://www.linkedin.com/in/kavejo/en/'

# Copyright statement for this module
Copyright = '(c) Tommaso Toniolo. All rights reserved.'

# Description of the functionality provided by this module
Description = 'This modules provides the ability to interact with mail systems such as Microsoft Exchange, Exchange Online, GMail, etc.'

# Minimum version of the PowerShell engine required by this module
# PowerShellVersion = ''

# Name of the PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# DotNetFrameworkVersion = '4.7.2'

# Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# CLRVersion = '4.7.2'

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
# RequiredModules = @()

# Assemblies that must be loaded prior to importing this module
RequiredAssemblies = @('Azure.Core.dll', 'Azure.Identity.dll', 'BouncyCastle.Crypto.dll', 'BouncyCastle.Cryptography.dll', 'MailKit.dll', 'Microsoft.Bcl.AsyncInterfaces.dll', 'Microsoft.Exchange.WebServices.Auth.dll', 'Microsoft.Exchange.WebServices.dll', 'Microsoft.Graph.Core.dll', 'Microsoft.Graph.dll', 'Microsoft.Identity.Client.dll', 'Microsoft.Identity.Client.Extensions.Msal.dll', 'Microsoft.IdentityModel.Abstractions.dll', 'Microsoft.IdentityModel.JsonWebTokens.dll', 'Microsoft.IdentityModel.Logging.dll', 'Microsoft.IdentityModel.Protocols.dll', 'Microsoft.IdentityModel.Protocols.OpenIdConnect.dll', 'Microsoft.IdentityModel.Tokens.dll', 'MimeKit.dll', 'PowerShellMailUtils.dll', 'PowerShellMailUtils.dll.config', 'System.Buffers.dll', 'System.ComponentModel.Annotations.dll', 'System.Diagnostics.DiagnosticSource.dll', 'System.IdentityModel.Tokens.Jwt.dll', 'System.Management.Automation.dll', 'System.Memory.Data.dll', 'System.Memory.dll', 'System.Net.Http.WinHttpHandler.dll', 'System.Numerics.Vectors.dll', 'System.Runtime.CompilerServices.Unsafe.dll', 'System.Security.Cryptography.ProtectedData.dll', 'System.Security.Principal.Windows.dll', 'System.Text.Encodings.Web.dll', 'System.Text.Json.dll', 'System.Threading.Tasks.Extensions.dll', 'System.ValueTuple.dll')

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
# NestedModules = @()

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = @()

# Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
CmdletsToExport = @()

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
AliasesToExport = @()

# DSC resources to export from this module
# DscResourcesToExport = @()

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
# FileList = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        # Tags = @()

        # A URL to the license for this module.
        # LicenseUri = ''

        # A URL to the main website for this project.
        ProjectUri = 'https://github.com/kavejo/PowerShellMailUtils'

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        # ReleaseNotes = ''

        # Prerelease string of this module
        # Prerelease = ''

        # Flag to indicate whether the module requires explicit user acceptance for install/update/save
        RequireLicenseAcceptance = $false

        # External dependent modules of this module
        # ExternalModuleDependencies = @()

    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}