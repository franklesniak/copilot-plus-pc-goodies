# TODO: add function header

# Version 0.2.20250503.0

# This script is written to work with Microsoft.ML version 4.0.2.
[CmdletBinding()]
param()

$strLabeledDataPath = 'C:\Users\flesniak\Downloads\Service-Desk-Tickets\Project-Phoenix-Incidents-Labeled.csv'
$strColumnNameForTextData = 'Description'
$strColumnNameForLabel = 'ClusterLabel'
$strONNXFileOutputPath = 'C:\Users\flesniak\Downloads\Service-Desk-Tickets\Project-Phoenix-Incidents-Labeled.onnx'

function Get-PSVersion {
    # .SYNOPSIS
    # Returns the version of PowerShell that is running.
    #
    # .DESCRIPTION
    # The function outputs a [version] object representing the version of
    # PowerShell that is running.
    #
    # On versions of PowerShell greater than or equal to version 2.0, this
    # function returns the equivalent of $PSVersionTable.PSVersion
    #
    # PowerShell 1.0 does not have a $PSVersionTable variable, so this
    # function returns [version]('1.0') on PowerShell 1.0.
    #
    # .EXAMPLE
    # $versionPS = Get-PSVersion
    # # $versionPS now contains the version of PowerShell that is running.
    # # On versions of PowerShell greater than or equal to version 2.0,
    # # this function returns the equivalent of $PSVersionTable.PSVersion.
    #
    # .INPUTS
    # None. You can't pipe objects to Get-PSVersion.
    #
    # .OUTPUTS
    # System.Version. Get-PSVersion returns a [version] value indiciating
    # the version of PowerShell that is running.
    #
    # .NOTES
    # Version: 1.0.20250106.0

    #region License ####################################################
    # Copyright (c) 2025 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining
    # a copy of this software and associated documentation files (the
    # "Software"), to deal in the Software without restriction, including
    # without limitation the rights to use, copy, modify, merge, publish,
    # distribute, sublicense, and/or sell copies of the Software, and to
    # permit persons to whom the Software is furnished to do so, subject to
    # the following conditions:
    #
    # The above copyright notice and this permission notice shall be
    # included in all copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
    # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
    # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
    # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
    # BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
    # ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
    # CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    # SOFTWARE.
    #endregion License ####################################################

    if (Test-Path variable:\PSVersionTable) {
        return ($PSVersionTable.PSVersion)
    } else {
        return ([version]('1.0'))
    }
}

function Test-NuGetDotOrgRegisteredAsPackageSource {
    <#
    .SYNOPSIS
    Tests to see nuget.org is registered as a package source. If it is not, the
    function can optionally throw an error or warning

    .DESCRIPTION
    The Test-NuGetDotOrgRegisteredAsPackageSource function tests to see if nuget.org is
    registered as a package source. If it is not, the function can optionally throw an
    error or warning that gives the user instructions to register nuget.org as a
    package source.

    .PARAMETER ThrowErrorIfNuGetDotOrgNotRegistered
    Is a switch parameter. If this parameter is specified, an error is thrown to tell
    the user that nuget.org is not registered as a package source, and the user is
    given instructions on how to register it. If this parameter is not specified, no
    error is thrown.

    .PARAMETER ThrowWarningIfNuGetDotOrgNotRegistered
    Is a switch parameter. If this parameter is specified, a warning is thrown to tell
    the user that nuget.org is not registered as a package source, and the user is
    given instructions on how to register it. If this parameter is not specified, or if
    the ThrowErrorIfNuGetDotOrgNotRegistered parameter was specified, no warning is
    thrown.

    .EXAMPLE
    $boolResult = Test-NuGetDotOrgRegisteredAsPackageSource -ThrowErrorIfNuGetDotOrgNotRegistered

    This example checks to see if nuget.org is registered as a package source. If it is
    not, an error is thrown to tell the user that nuget.org is not registered as a
    package source, and the user is given instructions on how to register it.

    .OUTPUTS
    [boolean] - Returns $true if nuget.org is registered as a package source; otherwise, returns $false.

    .NOTES
    Requires PowerShell v5.0 or newer
    #>

    #region License ################################################################
    # Copyright (c) 2025 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy of
    # this software and associated documentation files (the "Software"), to deal in the
    # Software without restriction, including without limitation the rights to use,
    # copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
    # Software, and to permit persons to whom the Software is furnished to do so,
    # subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in all
    # copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
    # FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
    # COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
    # AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
    # WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    #endregion License ################################################################

    #region DownloadLocationNotice  ################################################
    # The most up-to-date version of this script can be found on the author's GitHub
    # repository at https://github.com/franklesniak/PowerShell_Resources
    #endregion DownloadLocationNotice  ################################################

    # Version 1.1.20250106.0

    [CmdletBinding()]
    [OutputType([Boolean])]
    param (
        [Parameter(Mandatory = $false)][switch]$ThrowErrorIfNuGetDotOrgNotRegistered,
        [Parameter(Mandatory = $false)][switch]$ThrowWarningIfNuGetDotOrgNotRegistered
    )

    function Get-PSVersion {
        # .SYNOPSIS
        # Returns the version of PowerShell that is running.
        #
        # .DESCRIPTION
        # The function outputs a [version] object representing the version of
        # PowerShell that is running.
        #
        # On versions of PowerShell greater than or equal to version 2.0, this
        # function returns the equivalent of $PSVersionTable.PSVersion
        #
        # PowerShell 1.0 does not have a $PSVersionTable variable, so this
        # function returns [version]('1.0') on PowerShell 1.0.
        #
        # .EXAMPLE
        # $versionPS = Get-PSVersion
        # # $versionPS now contains the version of PowerShell that is running.
        # # On versions of PowerShell greater than or equal to version 2.0,
        # # this function returns the equivalent of $PSVersionTable.PSVersion.
        #
        # .INPUTS
        # None. You can't pipe objects to Get-PSVersion.
        #
        # .OUTPUTS
        # System.Version. Get-PSVersion returns a [version] value indiciating
        # the version of PowerShell that is running.
        #
        # .NOTES
        # Version: 1.0.20250106.0

        #region License ####################################################
        # Copyright (c) 2025 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining
        # a copy of this software and associated documentation files (the
        # "Software"), to deal in the Software without restriction, including
        # without limitation the rights to use, copy, modify, merge, publish,
        # distribute, sublicense, and/or sell copies of the Software, and to
        # permit persons to whom the Software is furnished to do so, subject to
        # the following conditions:
        #
        # The above copyright notice and this permission notice shall be
        # included in all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
        # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
        # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
        # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
        # BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
        # ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
        # CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ####################################################

        if (Test-Path variable:\PSVersionTable) {
            return ($PSVersionTable.PSVersion)
        } else {
            return ([version]('1.0'))
        }
    }

    $versionPS = Get-PSVersion
    if ($versionPS -lt ([version]'5.0')) {
        Write-Warning 'Test-NuGetDotOrgRegisteredAsPackageSource requires PowerShell version 5.0 or newer.'
        return
    }

    $WarningPreferenceAtStartOfFunction = $WarningPreference
    $VerbosePreferenceAtStartOfFunction = $VerbosePreference
    $DebugPreferenceAtStartOfFunction = $DebugPreference

    $boolThrowErrorForMissingPackageSource = $false
    $boolThrowWarningForMissingPackageSource = $false

    if ($ThrowErrorIfNuGetDotOrgNotRegistered.IsPresent -eq $true) {
        $boolThrowErrorForMissingPackageSource = $true
    } elseif ($ThrowWarningIfNuGetDotOrgNotRegistered.IsPresent -eq $true) {
        $boolThrowWarningForMissingPackageSource = $true
    }

    $boolPackageSourceFound = $true
    Write-Debug ('Checking for registered package sources (Get-PackageSource)...')
    $WarningPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    $VerbosePreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    $DebugPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    $arrPackageSources = @(Get-PackageSource)
    $WarningPreference = $WarningPreferenceAtStartOfFunction
    $VerbosePreference = $VerbosePreferenceAtStartOfFunction
    $DebugPreference = $DebugPreferenceAtStartOfFunction
    if (@($arrPackageSources | Where-Object { $_.Location -eq 'https://api.nuget.org/v3/index.json' }).Count -eq 0) {
        $boolPackageSourceFound = $false
    }

    if ($boolPackageSourceFound -eq $false) {
        $strMessage = 'The nuget.org package source is not registered. Please register it and then try again.' + [System.Environment]::NewLine + 'You can register it by running the following command: ' + [System.Environment]::NewLine + '[void](Register-PackageSource -Name NuGetOrg -Location https://api.nuget.org/v3/index.json -ProviderName NuGet);'

        if ($boolThrowErrorForMissingPackageSource -eq $true) {
            Write-Error $strMessage
        } elseif ($boolThrowWarningForMissingPackageSource -eq $true) {
            Write-Warning $strMessage
        }
    }

    return $boolPackageSourceFound
}

function Get-PackagesUsingHashtable {
    <#
    .SYNOPSIS
    Gets a list of installed "software packages" (typically NuGet packages) for each
    entry in a hashtable.

    .DESCRIPTION
    The Get-PackagesUsingHashtable function steps through each entry in the supplied
    hashtable. If a corresponding package is installed, then the information about
    the newest version of that package is stored in the value of the hashtable entry
    corresponding to the software package.

    .PARAMETER ReferenceToHashtable
    Is a reference to a hashtable. The value of the reference should be a hashtable
    with keys that are the names software packages and values that are initialized
    to be $null.

    .EXAMPLE
    $hashtablePackageNameToInstalledPackageMetadata = @{}
    $hashtablePackageNameToInstalledPackageMetadata.Add('Accord', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Accord.Math', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Accord.Statistics', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Accord.MachineLearning', $null)
    $refHashtablePackageNameToInstalledPackages = [ref]$hashtablePackageNameToInstalledPackageMetadata
    Get-PackagesUsingHashtable -ReferenceToHashtable $refHashtablePackageNameToInstalledPackages

    This example checks each of the four software packages specified. For each software
    package specified, if the software package is installed, the value of the hashtable
    entry will be set to the newest-installed version of the package. If the software
    package is not installed, the value of the hashtable entry remains $null.

    .OUTPUTS
    None

    .NOTES
    Requires PowerShell v5.0 or newer
    #>

    #region License ################################################################
    # Copyright (c) 2025 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy of
    # this software and associated documentation files (the "Software"), to deal in the
    # Software without restriction, including without limitation the rights to use,
    # copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
    # Software, and to permit persons to whom the Software is furnished to do so,
    # subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in all
    # copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
    # FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
    # COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
    # AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
    # WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    #endregion License ################################################################

    #region DownloadLocationNotice  ################################################
    # The most up-to-date version of this script can be found on the author's GitHub
    # repository at https://github.com/franklesniak/PowerShell_Resources
    #endregion DownloadLocationNotice  ################################################

    # Version 1.0.20250106.0

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][ref]$ReferenceToHashtable
    )

    function Get-PSVersion {
        # .SYNOPSIS
        # Returns the version of PowerShell that is running.
        #
        # .DESCRIPTION
        # The function outputs a [version] object representing the version of
        # PowerShell that is running.
        #
        # On versions of PowerShell greater than or equal to version 2.0, this
        # function returns the equivalent of $PSVersionTable.PSVersion
        #
        # PowerShell 1.0 does not have a $PSVersionTable variable, so this
        # function returns [version]('1.0') on PowerShell 1.0.
        #
        # .EXAMPLE
        # $versionPS = Get-PSVersion
        # # $versionPS now contains the version of PowerShell that is running.
        # # On versions of PowerShell greater than or equal to version 2.0,
        # # this function returns the equivalent of $PSVersionTable.PSVersion.
        #
        # .INPUTS
        # None. You can't pipe objects to Get-PSVersion.
        #
        # .OUTPUTS
        # System.Version. Get-PSVersion returns a [version] value indiciating
        # the version of PowerShell that is running.
        #
        # .NOTES
        # Version: 1.0.20250106.0

        #region License ####################################################
        # Copyright (c) 2025 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining
        # a copy of this software and associated documentation files (the
        # "Software"), to deal in the Software without restriction, including
        # without limitation the rights to use, copy, modify, merge, publish,
        # distribute, sublicense, and/or sell copies of the Software, and to
        # permit persons to whom the Software is furnished to do so, subject to
        # the following conditions:
        #
        # The above copyright notice and this permission notice shall be
        # included in all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
        # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
        # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
        # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
        # BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
        # ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
        # CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ####################################################

        if (Test-Path variable:\PSVersionTable) {
            return ($PSVersionTable.PSVersion)
        } else {
            return ([version]('1.0'))
        }
    }

    $versionPS = Get-PSVersion
    if ($versionPS -lt ([version]'5.0')) {
        Write-Warning 'Get-PackagesUsingHashtable requires PowerShell version 5.0 or newer.'
        return
    }

    $WarningPreferenceAtStartOfFunction = $WarningPreference
    $VerbosePreferenceAtStartOfFunction = $VerbosePreference
    $DebugPreferenceAtStartOfFunction = $DebugPreference

    $arrPackagesToGet = @(($ReferenceToHashtable.Value).Keys)

    $WarningPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    $VerbosePreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    $DebugPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
    $arrPackagesInstalled = @(Get-Package)
    $WarningPreference = $WarningPreferenceAtStartOfFunction
    $VerbosePreference = $VerbosePreferenceAtStartOfFunction
    $DebugPreference = $DebugPreferenceAtStartOfFunction

    for ($intCounter = 0; $intCounter -lt $arrPackagesToGet.Count; $intCounter++) {
        Write-Debug ('Checking for ' + $arrPackagesToGet[$intCounter] + ' software package...')
        $arrMatchingPackages = @($arrPackagesInstalled | Where-Object { $_.Name -eq $arrPackagesToGet[$intCounter] })
        if ($arrMatchingPackages.Count -eq 0) {
            ($ReferenceToHashtable.Value).Item($arrPackagesToGet[$intCounter]) = $null
        } else {
            ($ReferenceToHashtable.Value).Item($arrPackagesToGet[$intCounter]) = $arrMatchingPackages[0]
        }
    }
}

function Test-PackageInstalledUsingHashtable {
    <#
    .SYNOPSIS
    Tests to see if a software package (typically a NuGet package) is installed based
    on entries in a hashtable. If the software package is not installed, an error or
    warning message may optionally be displayed.

    .DESCRIPTION
    The Test-PackageInstalledUsingHashtable function steps through each entry in the
    supplied hashtable and, if there are any software packages not installed, it
    optionally throws an error or warning for each software package that is not
    installed. If all software packages are installed, the function returns $true;
    otherwise, if any software package is not installed, the function returns $false.

    .PARAMETER ReferenceToHashtableOfInstalledPackages
    Is a reference to a hashtable. The hashtable must have keys that are the names of
    software packages with each key's value populated with
    Microsoft.PackageManagement.Packaging.SoftwareIdentity objects (the result of
    Get-Package). If a software package is not installed, the value of the hashtable
    entry should be $null.

    .PARAMETER ReferenceToHashtableOfSkippingDependencies
    Is a reference to a hashtable. The hashtable must have keys that are the names of
    software packages with each key's value populated with a boolean value. The boolean
    indicates whether the software package should be installed without its
    dependencies. Generally, dependencies should not be skipped, so the default value
    for each key should be $false. However, sometimes the Install-Package command
    throws an erroneous dependency loop error, but in investigating its dependencies in
    the package's .nuspec file, you may find that the version of .NET that you will use
    has no dependencies. In this case, it's safe to use -SkipDependencies.

    This can also be verified here:
    https://www.nuget.org/packages/<PackageName>/#dependencies-body-tab

    If this parameter is not supplied, or if a key-value pair is not supplied in the
    hashtable for a given software package, the script will default to not skipping the
    software package's dependencies.

    .PARAMETER ThrowErrorIfPackageNotInstalled
    Is a switch parameter. If this parameter is specified, an error is thrown for each
    software package that is not installed. If this parameter is not specified, no
    error is thrown.

    .PARAMETER ThrowWarningIfPackageNotInstalled
    Is a switch parameter. If this parameter is specified, a warning is thrown for each
    software package that is not installed. If this parameter is not specified, or if
    the ThrowErrorIfPackageNotInstalled parameter was specified, no warning is thrown.

    .PARAMETER ReferenceToHashtableOfCustomNotInstalledMessages
    Is a reference to a hashtable. The hashtable must have keys that are custom error
    or warning messages (string) to be displayed if one or more software packages are
    not installed. The value for each key must be an array of software package names
    (strings) relevant to that error or warning message.

    If this parameter is not supplied, or if a custom error or warning message is not
    supplied in the hashtable for a given software package, the script will default to
    using the following message:

    <PACKAGENAME> software package not found. Please install it and then try again.
    You can install the <PACKAGENAME> software package by running the following
    command:
    Install-Package -ProviderName NuGet -Name <PACKAGENAME> -Force -Scope CurrentUser;

    .PARAMETER ReferenceToArrayOfMissingPackages
    Is a reference to an array. The array must be initialized to be empty. If any
    software packages are not installed, the names of those software packages are added
    to the array.

    .EXAMPLE
    $hashtablePackageNameToInstalledPackageMetadata = @{}
    $hashtablePackageNameToInstalledPackageMetadata.Add('Azure.Core', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Microsoft.Identity.Client', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Azure.Identity', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Accord', $null)
    $refHashtablePackageNameToInstalledPackages = [ref]$hashtablePackageNameToInstalledPackageMetadata
    Get-PackagesUsingHashtable -ReferenceToHashtable $refHashtablePackageNameToInstalledPackages
    $hashtableCustomNotInstalledMessageToPackageNames = @{}
    $strAzureIdentityNotInstalledMessage = 'Azure.Core, Microsoft.Identity.Client, and/or Azure.Identity packages were not found. Please install the Azure.Identity package and its dependencies and then try again.' + [System.Environment]::NewLine + 'You can install the Azure.Identity package and its dependencies by running the following command:' + [System.Environment]::NewLine + 'Install-Package -ProviderName NuGet -Name ''Azure.Identity'' -Force -Scope CurrentUser;' + [System.Environment]::NewLine + [System.Environment]::NewLine
    $hashtableCustomNotInstalledMessageToPackageNames.Add($strAzureIdentityNotInstalledMessage, @('Azure.Core', 'Microsoft.Identity.Client', 'Azure.Identity'))
    $refhashtableCustomNotInstalledMessageToPackageNames = [ref]$hashtableCustomNotInstalledMessageToPackageNames
    $boolResult = Test-PackageInstalledUsingHashtable -ReferenceToHashtableOfInstalledPackages $refHashtablePackageNameToInstalledPackages -ThrowErrorIfPackageNotInstalled -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToPackageNames

    This example checks to see if the Azure.Core, Microsoft.Identity.Client,
    Azure.Identity, and Accord packages are installed. If any of these packages are not
    installed, an error is thrown and $boolResult is set to $false. Because a custom
    error message was specified for the Azure.Core, Microsoft.Identity.Client, and
    Azure.Identity packages, if any one of those is missing, the custom error message
    is thrown just once. However, if Accord is missing, a separate error message would
    be thrown. If all packages are installed, $boolResult is set to $true.

    .OUTPUTS
    [boolean] - Returns $true if all Packages are installed; otherwise, returns $false.
    #>

    #region License ################################################################
    # Copyright (c) 2024 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy of
    # this software and associated documentation files (the "Software"), to deal in the
    # Software without restriction, including without limitation the rights to use,
    # copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
    # Software, and to permit persons to whom the Software is furnished to do so,
    # subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in all
    # copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
    # FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
    # COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
    # AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
    # WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    #endregion License ################################################################

    #region DownloadLocationNotice  ################################################
    # The most up-to-date version of this script can be found on the author's GitHub
    # repository at https://github.com/franklesniak/PowerShell_Resources
    #endregion DownloadLocationNotice  ################################################

    # Version 1.0.20240401.0

    [CmdletBinding()]
    [OutputType([Boolean])]
    param (
        [Parameter(Mandatory = $true)][ref]$ReferenceToHashtableOfInstalledPackages,
        [Parameter(Mandatory = $false)][ref]$ReferenceToHashtableOfSkippingDependencies,
        [Parameter(Mandatory = $false)][switch]$ThrowErrorIfPackageNotInstalled,
        [Parameter(Mandatory = $false)][switch]$ThrowWarningIfPackageNotInstalled,
        [Parameter(Mandatory = $false)][ref]$ReferenceToHashtableOfCustomNotInstalledMessages,
        [Parameter(Mandatory = $false)][ref]$ReferenceToArrayOfMissingPackages
    )

    $boolThrowErrorForMissingPackage = $false
    $boolThrowWarningForMissingPackage = $false

    if ($ThrowErrorIfPackageNotInstalled.IsPresent -eq $true) {
        $boolThrowErrorForMissingPackage = $true
    } elseif ($ThrowWarningIfPackageNotInstalled.IsPresent -eq $true) {
        $boolThrowWarningForMissingPackage = $true
    }

    $boolResult = $true

    $hashtableMessagesToThrowForMissingPackage = @{}
    $hashtablePackageNameToCustomMessageToThrowForMissingPackage = @{}
    if ($null -ne $ReferenceToHashtableOfCustomNotInstalledMessages) {
        $arrMessages = @(($ReferenceToHashtableOfCustomNotInstalledMessages.Value).Keys)
        foreach ($strMessage in $arrMessages) {
            $hashtableMessagesToThrowForMissingPackage.Add($strMessage, $false)

            ($ReferenceToHashtableOfCustomNotInstalledMessages.Value).Item($strMessage) | ForEach-Object {
                $hashtablePackageNameToCustomMessageToThrowForMissingPackage.Add($_, $strMessage)
            }
        }
    }

    $arrPackageNames = @(($ReferenceToHashtableOfInstalledPackages.Value).Keys)
    foreach ($strPackageName in $arrPackageNames) {
        if ($null -eq ($ReferenceToHashtableOfInstalledPackages.Value).Item($strPackageName)) {
            # Package not installed
            $boolResult = $false

            if ($hashtablePackageNameToCustomMessageToThrowForMissingPackage.ContainsKey($strPackageName) -eq $true) {
                $strMessage = $hashtablePackageNameToCustomMessageToThrowForMissingPackage.Item($strPackageName)
                $hashtableMessagesToThrowForMissingPackage.Item($strMessage) = $true
            } else {
                if ($null -ne $ReferenceToHashtableOfSkippingDependencies) {
                    if (($ReferenceToHashtableOfSkippingDependencies.Value).ContainsKey($strPackageName) -eq $true) {
                        $boolSkipDependencies = ($ReferenceToHashtableOfSkippingDependencies.Value).Item($strPackageName)
                    } else {
                        $boolSkipDependencies = $false
                    }
                } else {
                    $boolSkipDependencies = $false
                }

                if ($boolSkipDependencies -eq $true) {
                    $strMessage = $strPackageName + ' software package not found. Please install it and then try again.' + [System.Environment]::NewLine + 'You can install the ' + $strPackageName + ' package by running the following command:' + [System.Environment]::NewLine + 'Install-Package -ProviderName NuGet -Name ''' + $strPackageName + ''' -Force -Scope CurrentUser -SkipDependencies;' + [System.Environment]::NewLine + [System.Environment]::NewLine
                } else {
                    $strMessage = $strPackageName + ' software package not found. Please install it and then try again.' + [System.Environment]::NewLine + 'You can install the ' + $strPackageName + ' package by running the following command:' + [System.Environment]::NewLine + 'Install-Package -ProviderName NuGet -Name ''' + $strPackageName + ''' -Force -Scope CurrentUser;' + [System.Environment]::NewLine + [System.Environment]::NewLine
                }

                $hashtableMessagesToThrowForMissingPackage.Add($strMessage, $true)
            }

            if ($null -ne $ReferenceToArrayOfMissingPackages) {
                ($ReferenceToArrayOfMissingPackages.Value) += $strPackageName
            }
        }
    }

    if ($boolThrowErrorForMissingPackage -eq $true) {
        $arrMessages = @($hashtableMessagesToThrowForMissingPackage.Keys)
        foreach ($strMessage in $arrMessages) {
            if ($hashtableMessagesToThrowForMissingPackage.Item($strMessage) -eq $true) {
                Write-Error $strMessage
            }
        }
    } elseif ($boolThrowWarningForMissingPackage -eq $true) {
        $arrMessages = @($hashtableMessagesToThrowForMissingPackage.Keys)
        foreach ($strMessage in $arrMessages) {
            if ($hashtableMessagesToThrowForMissingPackage.Item($strMessage) -eq $true) {
                Write-Warning $strMessage
            }
        }
    }

    return $boolResult
}

function Get-DLLPathsForPackagesUsingHashtable {
    <#
    .SYNOPSIS
    Using a hashtable of installed software package metadata, gets the path to the
    .dll file(s) within each software package that is most appropriate to use

    .DESCRIPTION
    Software packages contain .dll files for different .NET Framework versions. This
    function steps through each entry in the supplied hashtable. If a corresponding
    package is installed, then the path to the .dll file(s) within the package that is
    most appropriate to use is stored in the value of the hashtable entry corresponding
    to the software package.

    .PARAMETER ReferenceToHashtableOfInstalledPackages
    Is a reference to a hashtable. The hashtable must have keys that are the names of
    software packages with each key's value populated with
    Microsoft.PackageManagement.Packaging.SoftwareIdentity objects (the result of
    Get-Package). If a software package is not installed, the value of the hashtable
    entry should be $null.

    .PARAMETER ReferenceToHashtableOfSpecifiedDotNETVersions
    Is an optional parameter. If supplied, it must be a reference to a hashtable. The
    hashtable must have keys that are the names of software packages with each key's
    value populated with a string that is the version of .NET Framework that the
    software package is to be used with. If a key-value pair is not supplied in the
    hashtable for a given software package, the function will default to doing its best
    to select the most appropriate version of the software package given the current
    operating environment and PowerShell version.

    .PARAMETER ReferenceToHashtableOfEffectiveDotNETVersions
    Is initially a reference to an empty hashtable. When execution completes, the
    hashtable will be populated with keys that are the names of the software packages
    specified in the hashtable referenced by the
    ReferenceToHashtableOfInstalledPackages parameter. The value of each entry will be
    a string that is the folder corresponding to the version of .NET that makes the
    most sense given the current platform and .NET Framework version. If no suitable
    folder is found, the value of the hashtable entry remains an empty string.

    For example, reference the following folder name taxonomy at nuget.org:
    https://www.nuget.org/packages/System.Text.Json#supportedframeworks-body-tab

    .PARAMETER ReferenceToHashtableOfDLLPaths
    Is initially a reference to an empty hashtable. When execution completes, the
    hashtable will be populated with keys that are the names of the software packages
    specified in the hashtable referenced by the
    ReferenceToHashtableOfInstalledPackages parameter. The value of each entry will be
    an array populated with the path to the .dll file(s) within the package that are
    most appropriate to use, given the current platform and .NET Framework version.
    If no suitable DLL file is found, the array will be empty.

    .EXAMPLE
    $hashtablePackageNameToInstalledPackageMetadata = @{}
    $hashtablePackageNameToInstalledPackageMetadata.Add('Azure.Core', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Microsoft.Identity.Client', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Azure.Identity', $null)
    $hashtablePackageNameToInstalledPackageMetadata.Add('Accord', $null)
    $refHashtablePackageNameToInstalledPackages = [ref]$hashtablePackageNameToInstalledPackageMetadata
    Get-PackagesUsingHashtable -ReferenceToHashtable $refHashtablePackageNameToInstalledPackages
    $hashtableCustomNotInstalledMessageToPackageNames = @{}
    $strAzureIdentityNotInstalledMessage = 'Azure.Core, Microsoft.Identity.Client, and/or Azure.Identity packages were not found. Please install the Azure.Identity package and its dependencies and then try again.' + [System.Environment]::NewLine + 'You can install the Azure.Identity package and its dependencies by running the following command:' + [System.Environment]::NewLine + 'Install-Package -ProviderName NuGet -Name 'Azure.Identity' -Force -Scope CurrentUser;' + [System.Environment]::NewLine + [System.Environment]::NewLine
    $hashtableCustomNotInstalledMessageToPackageNames.Add($strAzureIdentityNotInstalledMessage, @('Azure.Core', 'Microsoft.Identity.Client', 'Azure.Identity'))
    $refhashtableCustomNotInstalledMessageToPackageNames = [ref]$hashtableCustomNotInstalledMessageToPackageNames
    $boolResult = Test-PackageInstalledUsingHashtable -ReferenceToHashtableOfInstalledPackages $refHashtablePackageNameToInstalledPackages -ThrowErrorIfPackageNotInstalled -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToPackageNames
    if ($boolResult -eq $false) { return }
    $hashtablePackageNameToEffectiveDotNETVersions = @{}
    $refHashtablePackageNameToEffectiveDotNETVersions = [ref]$hashtablePackageNameToEffectiveDotNETVersions
    $hashtablePackageNameToDLLPaths = @{}
    $refHashtablePackageNameToDLLPaths = [ref]$hashtablePackageNameToDLLPaths
    Get-DLLPathsForPackagesUsingHashtable -ReferenceToHashtableOfInstalledPackages $refHashtablePackageNameToInstalledPackages -ReferenceToHashtableOfEffectiveDotNETVersions $refHashtablePackageNameToEffectiveDotNETVersions -ReferenceToHashtableOfDLLPaths $refHashtablePackageNameToDLLPaths

    This example checks each of the four software packages specified. For each software
    package specified, if the software package is installed, the value of the hashtable
    entry will be set to the path to the .dll file(s) within the package that are most
    appropriate to use, given the current platform and .NET Framework version. If no
    suitable DLL file is found, the value of the hashtable entry remains an empty array
    (@()).

    .OUTPUTS
    None

    .NOTES
    Requires PowerShell v5.0 or newer
    #>

    #region License ################################################################
    # Copyright (c) 2025 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy of
    # this software and associated documentation files (the "Software"), to deal in the
    # Software without restriction, including without limitation the rights to use,
    # copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
    # Software, and to permit persons to whom the Software is furnished to do so,
    # subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in all
    # copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
    # FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
    # COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
    # AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
    # WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    #endregion License ################################################################

    #region DownloadLocationNotice  ################################################
    # The most up-to-date version of this script can be found on the author's GitHub
    # repository at https://github.com/franklesniak/PowerShell_Resources
    #endregion DownloadLocationNotice  ################################################

    # Version 1.1.20250106.0

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][ref]$ReferenceToHashtableOfInstalledPackages,
        [Parameter(Mandatory = $false)][ref]$ReferenceToHashtableOfSpecifiedDotNETVersions,
        [Parameter(Mandatory = $true)][ref]$ReferenceToHashtableOfEffectiveDotNETVersions,
        [Parameter(Mandatory = $true)][ref]$ReferenceToHashtableOfDLLPaths
    )

    function Get-PSVersion {
        # .SYNOPSIS
        # Returns the version of PowerShell that is running.
        #
        # .DESCRIPTION
        # The function outputs a [version] object representing the version of
        # PowerShell that is running.
        #
        # On versions of PowerShell greater than or equal to version 2.0, this
        # function returns the equivalent of $PSVersionTable.PSVersion
        #
        # PowerShell 1.0 does not have a $PSVersionTable variable, so this
        # function returns [version]('1.0') on PowerShell 1.0.
        #
        # .EXAMPLE
        # $versionPS = Get-PSVersion
        # # $versionPS now contains the version of PowerShell that is running.
        # # On versions of PowerShell greater than or equal to version 2.0,
        # # this function returns the equivalent of $PSVersionTable.PSVersion.
        #
        # .INPUTS
        # None. You can't pipe objects to Get-PSVersion.
        #
        # .OUTPUTS
        # System.Version. Get-PSVersion returns a [version] value indiciating
        # the version of PowerShell that is running.
        #
        # .NOTES
        # Version: 1.0.20250106.0

        #region License ####################################################
        # Copyright (c) 2025 Frank Lesniak
        #
        # Permission is hereby granted, free of charge, to any person obtaining
        # a copy of this software and associated documentation files (the
        # "Software"), to deal in the Software without restriction, including
        # without limitation the rights to use, copy, modify, merge, publish,
        # distribute, sublicense, and/or sell copies of the Software, and to
        # permit persons to whom the Software is furnished to do so, subject to
        # the following conditions:
        #
        # The above copyright notice and this permission notice shall be
        # included in all copies or substantial portions of the Software.
        #
        # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
        # EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
        # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
        # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
        # BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
        # ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
        # CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        # SOFTWARE.
        #endregion License ####################################################

        if (Test-Path variable:\PSVersionTable) {
            return ($PSVersionTable.PSVersion)
        } else {
            return ([version]('1.0'))
        }
    }

    $versionPS = Get-PSVersion
    if ($versionPS -lt ([version]'5.0')) {
        Write-Warning 'Get-DLLPathsForPackagesUsingHashtable requires PowerShell version 5.0 or newer.'
        return
    }

    $arrPackageNames = @(($ReferenceToHashtableOfInstalledPackages.Value).Keys)
    foreach ($strPackageName in $arrPackageNames) {
        ($ReferenceToHashtableOfEffectiveDotNETVersions.Value).Add($strPackageName, '')
        ($ReferenceToHashtableOfDLLPaths.Value).Add($strPackageName, @())
    }

    # Get the base folder path for each package
    $hashtablePackageNameToBaseFolderPath = @{}
    foreach ($strPackageName in $arrPackageNames) {
        if ($null -ne ($ReferenceToHashtableOfInstalledPackages.Value).Item($strPackageName)) {
            $strPackageFilePath = ($ReferenceToHashtableOfInstalledPackages.Value).Item($strPackageName).Source
            $strPackageFilePath = $strPackageFilePath.Replace('file:///', '')
            $strPackageFileParentFolderPath = [System.IO.Path]::GetDirectoryName($strPackageFilePath)

            $hashtablePackageNameToBaseFolderPath.Add($strPackageName, $strPackageFileParentFolderPath)
        }
    }

    # Determine the current platform
    $boolIsLinux = $false
    if (Test-Path variable:\IsLinux) {
        if ($IsLinux -eq $true) {
            $boolIsLinux = $true
        }
    }

    $boolIsMacOS = $false
    if (Test-Path variable:\IsMacOS) {
        if ($IsMacOS -eq $true) {
            $boolIsMacOS = $true
        }
    }

    if ($boolIsLinux -eq $true) {
        if (($versionPS -ge [version]'7.5') -and ($versionPS -lt [version]'7.6')) {
            # .NET 9.0
            $arrDotNETVersionPreferenceOrder = @('net9.0', 'net8.0', 'net7.0', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.4') -and ($versionPS -lt [version]'7.5')) {
            # .NET 8.0
            $arrDotNETVersionPreferenceOrder = @('net8.0', 'net7.0', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.3') -and ($versionPS -lt [version]'7.4')) {
            # .NET 7.0
            $arrDotNETVersionPreferenceOrder = @('net7.0', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.2') -and ($versionPS -lt [version]'7.3')) {
            # .NET 6.0
            $arrDotNETVersionPreferenceOrder = @('net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.1') -and ($versionPS -lt [version]'7.2')) {
            # .NET 5.0
            $arrDotNETVersionPreferenceOrder = @('net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.0') -and ($versionPS -lt [version]'7.1')) {
            # .NET Core 3.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.2') -and ($versionPS -lt [version]'7.0')) {
            # .NET Core 2.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.1') -and ($versionPS -lt [version]'6.2')) {
            # .NET Core 2.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.0') -and ($versionPS -lt [version]'6.1')) {
            # .NET Core 2.0
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } else {
            # A future, undefined version of PowerShell
            $arrDotNETVersionPreferenceOrder = @('net15.0', 'net14.0', 'net13.0', 'net12.0', 'net11.0', 'net10.0', 'net9.0', 'net8.0', 'net7.0', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        }
    } elseif ($boolIsMacOS -eq $true) {
        if (($versionPS -ge [version]'7.5') -and ($versionPS -lt [version]'7.6')) {
            # .NET 9.0
            $arrDotNETVersionPreferenceOrder = @('net9.0-macos', 'net9.0', 'net8.0-macos', 'net8.0', 'net7.0-macos', 'net7.0', 'net6.0-macos', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.4') -and ($versionPS -lt [version]'7.5')) {
            # .NET 8.0
            $arrDotNETVersionPreferenceOrder = @('net8.0-macos', 'net8.0', 'net7.0-macos', 'net7.0', 'net6.0-macos', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.3') -and ($versionPS -lt [version]'7.4')) {
            # .NET 7.0
            $arrDotNETVersionPreferenceOrder = @('net7.0-macos', 'net7.0', 'net6.0-macos', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.2') -and ($versionPS -lt [version]'7.3')) {
            # .NET 6.0
            $arrDotNETVersionPreferenceOrder = @('net6.0-macos', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.1') -and ($versionPS -lt [version]'7.2')) {
            # .NET 5.0
            $arrDotNETVersionPreferenceOrder = @('net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.0') -and ($versionPS -lt [version]'7.1')) {
            # .NET Core 3.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.2') -and ($versionPS -lt [version]'7.0')) {
            # .NET Core 2.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.1') -and ($versionPS -lt [version]'6.2')) {
            # .NET Core 2.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.0') -and ($versionPS -lt [version]'6.1')) {
            # .NET Core 2.0
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } else {
            # A future, undefined version of PowerShell
            $arrDotNETVersionPreferenceOrder = @('net15.0-macos', 'net15.0', 'net14.0-macos', 'net14.0', 'net13.0-macos', 'net13.0', 'net12.0-macos', 'net12.0', 'net11.0-macos', 'net11.0', 'net10.0-macos', 'net10.0', 'net9.0-macos', 'net9.0', 'net8.0-macos', 'net8.0', 'net7.0-macos', 'net7.0', 'net6.0-macos', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        }
    } else {
        # Windows
        if (($versionPS -ge [version]'7.5') -and ($versionPS -lt [version]'7.6')) {
            # .NET 9.0
            $arrDotNETVersionPreferenceOrder = @('net9.0-windows', 'net9.0', 'net8.0-windows', 'net8.0', 'net7.0-windows', 'net7.0', 'net6.0-windows', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.4') -and ($versionPS -lt [version]'7.5')) {
            # .NET 8.0
            $arrDotNETVersionPreferenceOrder = @('net8.0-windows', 'net8.0', 'net7.0-windows', 'net7.0', 'net6.0-windows', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.3') -and ($versionPS -lt [version]'7.4')) {
            # .NET 7.0
            $arrDotNETVersionPreferenceOrder = @('net7.0-windows', 'net7.0', 'net6.0-windows', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.2') -and ($versionPS -lt [version]'7.3')) {
            # .NET 6.0
            $arrDotNETVersionPreferenceOrder = @('net6.0-windows', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.1') -and ($versionPS -lt [version]'7.2')) {
            # .NET 5.0
            $arrDotNETVersionPreferenceOrder = @('net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'7.0') -and ($versionPS -lt [version]'7.1')) {
            # .NET Core 3.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.2') -and ($versionPS -lt [version]'7.0')) {
            # .NET Core 2.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.1') -and ($versionPS -lt [version]'6.2')) {
            # .NET Core 2.1
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'6.0') -and ($versionPS -lt [version]'6.1')) {
            # .NET Core 2.0
            $arrDotNETVersionPreferenceOrder = @('netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        } elseif (($versionPS -ge [version]'5.0') -and ($versionPS -lt [version]'6.0')) {
            if ((Test-Path 'HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full') -eq $true) {
                $intDotNETFrameworkRelease = (Get-ItemPropertyValue -LiteralPath 'HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release)
                # if ($intDotNETFrameworkRelease -ge 533320) {
                #     # .NET Framework 4.8.1
                #     $arrDotNETVersionPreferenceOrder = @('net481', 'net48', 'net472', 'net471', 'netstandard2.0', 'net47', 'net463', 'net462', 'net461', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4', 'net46', 'net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 528040) {
                #     # .NET Framework 4.8
                #     $arrDotNETVersionPreferenceOrder = @('net48', 'net472', 'net471', 'netstandard2.0', 'net47', 'net463', 'net462', 'net461', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4', 'net46', 'net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 461808) {
                #     # .NET Framework 4.7.2
                #     $arrDotNETVersionPreferenceOrder = @('net472', 'net471', 'netstandard2.0', 'net47', 'net463', 'net462', 'net461', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4', 'net46', 'net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 461308) {
                #     # .NET Framework 4.7.1
                #     $arrDotNETVersionPreferenceOrder = @('net471', 'netstandard2.0', 'net47', 'net463', 'net462', 'net461', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4', 'net46', 'net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 460798) {
                #     # .NET Framework 4.7
                #     $arrDotNETVersionPreferenceOrder = @('net47', 'net463', 'net462', 'net461', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4', 'net46', 'net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 394802) {
                #     # .NET Framework 4.6.2
                #     $arrDotNETVersionPreferenceOrder = @('net462', 'net461', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4', 'net46', 'net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 394254) {
                #     # .NET Framework 4.6.1
                #     $arrDotNETVersionPreferenceOrder = @('net461', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4', 'net46', 'net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 393295) {
                #     # .NET Framework 4.6
                #     $arrDotNETVersionPreferenceOrder = @('net46', 'net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 379893) {
                #     # .NET Framework 4.5.2
                #     $arrDotNETVersionPreferenceOrder = @('net452', 'net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 378675) {
                #     # .NET Framework 4.5.1
                #     $arrDotNETVersionPreferenceOrder = @('net451', 'net45', 'net40')
                # } elseif ($intDotNETFrameworkRelease -ge 378389) {
                #     # .NET Framework 4.5
                #     $arrDotNETVersionPreferenceOrder = @('net45', 'net40')
                # } else {
                #     # .NET Framework 4.5 or newer not found?
                #     # This should not be possible since this function requires
                #     # PowerShell 5.0 or newer, PowerShell 5.0 requires WMF 5.0, and
                #     # WMF 5.0 requires .NET Framework 4.5 or newer.
                #     Write-Warning 'The .NET Framework 4.5 or newer was not found. This should not be possible since this function requires PowerShell 5.0 or newer, PowerShell 5.0 requires WMF 5.0, and WMF 5.0 requires .NET Framework 4.5 or newer.'
                #     return
                # }
                if ($intDotNETFrameworkRelease -ge 533320) {
                    # .NET Framework 4.8.1
                    $arrDotNETVersionPreferenceOrder = @('net481', 'net48', 'net472', 'net471', 'net47', 'net463', 'net462', 'net461', 'net46', 'net452', 'net451', 'net45', 'net40', 'netstandard2.0', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4')
                } elseif ($intDotNETFrameworkRelease -ge 528040) {
                    # .NET Framework 4.8
                    $arrDotNETVersionPreferenceOrder = @('net48', 'net472', 'net471', 'net47', 'net463', 'net462', 'net461', 'net46', 'net452', 'net451', 'net45', 'net40', 'netstandard2.0', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4')
                } elseif ($intDotNETFrameworkRelease -ge 461808) {
                    # .NET Framework 4.7.2
                    $arrDotNETVersionPreferenceOrder = @('net472', 'net471', 'net47', 'net463', 'net462', 'net461', 'net46', 'net452', 'net451', 'net45', 'net40', 'netstandard2.0', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4')
                } elseif ($intDotNETFrameworkRelease -ge 461308) {
                    # .NET Framework 4.7.1
                    $arrDotNETVersionPreferenceOrder = @('net471', 'net47', 'net463', 'net462', 'net461', 'net46', 'net452', 'net451', 'net45', 'net40', 'netstandard2.0', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4')
                } elseif ($intDotNETFrameworkRelease -ge 460798) {
                    # .NET Framework 4.7
                    $arrDotNETVersionPreferenceOrder = @('net47', 'net463', 'net462', 'net461', 'net46', 'net452', 'net451', 'net45', 'net40', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4')
                } elseif ($intDotNETFrameworkRelease -ge 394802) {
                    # .NET Framework 4.6.2
                    $arrDotNETVersionPreferenceOrder = @('net462', 'net461', 'net46', 'net452', 'net451', 'net45', 'net40', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4')
                } elseif ($intDotNETFrameworkRelease -ge 394254) {
                    # .NET Framework 4.6.1
                    $arrDotNETVersionPreferenceOrder = @('net461', 'net46', 'net452', 'net451', 'net45', 'net40', 'netstandard1.6', 'netstandard1.5', 'netstandard1.4')
                } elseif ($intDotNETFrameworkRelease -ge 393295) {
                    # .NET Framework 4.6
                    $arrDotNETVersionPreferenceOrder = @('net46', 'net452', 'net451', 'net45', 'net40')
                } elseif ($intDotNETFrameworkRelease -ge 379893) {
                    # .NET Framework 4.5.2
                    $arrDotNETVersionPreferenceOrder = @('net452', 'net451', 'net45', 'net40')
                } elseif ($intDotNETFrameworkRelease -ge 378675) {
                    # .NET Framework 4.5.1
                    $arrDotNETVersionPreferenceOrder = @('net451', 'net45', 'net40')
                } elseif ($intDotNETFrameworkRelease -ge 378389) {
                    # .NET Framework 4.5
                    $arrDotNETVersionPreferenceOrder = @('net45', 'net40')
                } else {
                    # .NET Framework 4.5 or newer not found?
                    # This should not be possible since this function requires
                    # PowerShell 5.0 or newer, PowerShell 5.0 requires WMF 5.0, and
                    # WMF 5.0 requires .NET Framework 4.5 or newer.
                    Write-Warning 'The .NET Framework 4.5 or newer was not found. This should not be possible since this function requires PowerShell 5.0 or newer, PowerShell 5.0 requires WMF 5.0, and WMF 5.0 requires .NET Framework 4.5 or newer.'
                    return
                }
            }
        } else {
            # A future, undefined version of PowerShell
            $arrDotNETVersionPreferenceOrder = @('net15.0-windows', 'net15.0', 'net14.0-windows', 'net14.0', 'net13.0-windows', 'net13.0', 'net12.0-windows', 'net12.0', 'net11.0-windows', 'net11.0', 'net10.0-windows', 'net10.0', 'net9.0-windows', 'net9.0', 'net8.0-windows', 'net8.0', 'net7.0-windows', 'net7.0', 'net6.0-windows', 'net6.0', 'net5.0', 'netcoreapp3.1', 'netcoreapp3.0', 'netstandard2.1', 'netcoreapp2.2', 'netcoreapp2.1', 'netcoreapp2.0', 'netstandard2.0', 'netcoreapp1.1', 'netcoreapp1.0', 'netstandard1.6')
        }
    }

    foreach ($strPackageName in $arrPackageNames) {
        if ($null -ne $hashtablePackageNameToBaseFolderPath.Item($strPackageName)) {
            $strBaseFolderPath = ($hashtablePackageNameToBaseFolderPath.Item($strPackageName))

            $strDLLFolderPath = ''
            if ($null -ne $ReferenceToHashtableOfSpecifiedDotNETVersions) {
                if ($null -ne ($ReferenceToHashtableOfSpecifiedDotNETVersions.Value).Item($strPackageName)) {
                    $strDotNETVersion = ($ReferenceToHashtableOfSpecifiedDotNETVersions.Value).Item($strPackageName)

                    if ([string]::IsNullOrEmpty($strDotNETVersion) -eq $false) {
                        $strDLLFolderPath = Join-Path -Path $strBaseFolderPath -ChildPath 'lib'
                        $strDLLFolderPath = Join-Path -Path $strDLLFolderPath -ChildPath $strDotNETVersion

                        # Search this folder for .dll files and add them to the array
                        if (Test-Path -Path $strDLLFolderPath -PathType Container) {
                            $arrDLLFiles = @(Get-ChildItem -Path $strDLLFolderPath -Filter '*.dll' -File -Recurse)
                            if ($arrDLLFiles.Count -gt 0) {
                                # One or more DLL files found
                                ($ReferenceToHashtableOfEffectiveDotNETVersions.Value).Item($strPackageName) = $strDotNETVersion
                                ($ReferenceToHashtableOfDLLPaths.Value).Item($strPackageName) = @($arrDLLFiles | ForEach-Object {
                                        $_.FullName
                                    })
                            }
                        } else {
                            # The specified .NET version folder does not exist
                            # Set the DLL folder path to an empty string to then do a
                            # search for a usable folder
                            $strDLLFolderPath = ''
                        }
                    }
                }
            }

            if ([string]::IsNullOrEmpty($strDLLFolderPath)) {
                # Do a search for a usable folder

                foreach ($strDotNETVersion in $arrDotNETVersionPreferenceOrder) {
                    $strDLLFolderPath = Join-Path -Path $strBaseFolderPath -ChildPath 'lib'
                    $strDLLFolderPath = Join-Path -Path $strDLLFolderPath -ChildPath $strDotNETVersion

                    if (Test-Path -Path $strDLLFolderPath -PathType Container) {
                        $arrDLLFiles = @(Get-ChildItem -Path $strDLLFolderPath -Filter '*.dll' -File -Recurse)
                        if ($arrDLLFiles.Count -gt 0) {
                            # One or more DLL files found
                            ($ReferenceToHashtableOfEffectiveDotNETVersions.Value).Item($strPackageName) = $strDotNETVersion
                            ($ReferenceToHashtableOfDLLPaths.Value).Item($strPackageName) = @($arrDLLFiles | ForEach-Object {
                                    $_.FullName
                                })
                            break
                        }
                    }
                }
            }
        }
    }
}

function Test-PipelineStepONNXExportable {
    param (
        [Microsoft.ML.MLContext]$MLContext,
        [Microsoft.ML.IDataView]$SampleData,
        [System.Collections.Generic.List[object]]$PipelineSteps,
        [int]$OpSet = 12
    )

    $pipeline = $null
    $exportable = @()
    $failingStep = $null

    foreach ($step in $PipelineSteps) {
        if (-not $pipeline) {
            $pipeline = $step
        } else {
            $pipeline = [Microsoft.ML.LearningPipelineExtensions]::Append($pipeline, $step, [Microsoft.ML.Data.TransformerScope]::Everything)
        }

        try {
            $model = $pipeline.Fit($SampleData)

            $stream = New-Object System.IO.MemoryStream
            [Microsoft.ML.OnnxExportExtensions]::ConvertToOnnx($MLContext.Model, $model, $SampleData, $OpSet, $stream)
            $stream.Dispose()

            $exportable += $step.ToString()
        } catch {
            $failingStep = $step.ToString()
            break
        }
    }

    [PSCustomObject]@{
        ExportableSteps = $exportable
        FailingStep     = $failingStep
    }
}

function ConvertTo-NewlinelessText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)][string]$Text
    )
    process {
        $CRLF = "`r`n"
        $LF = [char]"`n"
        $CR = [char]"`r"
        $CRLFLessText = $Text -replace $CRLF, ' '
        $sb = New-Object -TypeName 'Text.StringBuilder'
        foreach ($ch in $CRLFLessText.ToCharArray()) {
            if ($ch -ne $LF -and $ch -ne $CR) {
                [void]$sb.Append($ch)
            } else {
                [void]$sb.Append(' ')
            }
        }
        $sb.ToString()
    }
}

function ConvertTo-UnaccentedText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)][string]$Text
    )
    process {
        # 1. Canonical‑decompose: é → e + ́
        $decomposed = $Text.Normalize([Text.NormalizationForm]::FormD)

        # 2. Keep everything that is *not* a combining mark
        $sb = New-Object -TypeName 'Text.StringBuilder'
        foreach ($ch in $decomposed.ToCharArray()) {
            if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne
                [System.Globalization.UnicodeCategory]::NonSpacingMark) {
                [void]$sb.Append($ch)
            }
        }
        $sb.ToString()
    }
}

function ConvertTo-UnpunctuatedText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)][string]$Text
    )
    process {
        $sb = New-Object -TypeName 'Text.StringBuilder'
        foreach ($ch in $Text.ToCharArray()) {
            if (-not [char]::IsPunctuation($ch)) { [void]$sb.Append($ch) }
        }
        $sb.ToString()
    }
}

function ConvertTo-NumberlessText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)][string]$Text
    )
    process {
        $sb = New-Object -TypeName 'Text.StringBuilder'
        foreach ($ch in $Text.ToCharArray()) {
            if (-not [char]::IsDigit($ch)) { [void]$sb.Append($ch) }
        }
        $sb.ToString()
    }
}

#region Validate PowerShell version ################################################
$versionPS = Get-PSVersion
if ($versionPS.Major -lt 6) {
    Write-Warning 'This script requires PowerShell Core 6.0 or newer.'
    return
}
#endregion Validate PowerShell version ################################################

#region Make sure that nuget.org is registered as a package source #################
# (if not, throw a warning and quit)
$boolNuGetDotOrgRegisteredAsPackageSource = Test-NuGetDotOrgRegisteredAsPackageSource -ThrowWarningIfNuGetDotOrgNotRegistered
if ($boolNuGetDotOrgRegisteredAsPackageSource -eq $false) {
    return # Quit script
}
#endregion Make sure that nuget.org is registered as a package source #################

#region Make sure that required NuGet packages are installed #######################
$hashtablePackageNameToInstalledPackageMetadata = @{}

$hashtablePackageNameToInstalledPackageMetadata.Add('Google.Protobuf', $null)
$hashtablePackageNameToInstalledPackageMetadata.Add('Microsoft.ML', $null)
$hashtablePackageNameToInstalledPackageMetadata.Add('Microsoft.ML.CpuMath', $null)
$hashtablePackageNameToInstalledPackageMetadata.Add('Microsoft.ML.DataView', $null)
$hashtablePackageNameToInstalledPackageMetadata.Add('Microsoft.ML.OnnxConverter', $null)
$hashtablePackageNameToInstalledPackageMetadata.Add('System.Numerics.Tensors', $null)

$refHashtablePackageNameToInstalledPackages = [ref]$hashtablePackageNameToInstalledPackageMetadata
Get-PackagesUsingHashtable -ReferenceToHashtable $refHashtablePackageNameToInstalledPackages

$hashtablePackageNameToSkippingDependencies = @{}
$refHashtablePackageNameToSkippingDependencies = [ref]$hashtablePackageNameToSkippingDependencies

$hashtableCustomNotInstalledMessageToPackageNames = @{}
$strAccordMachineLearningNotInstalledMessage = 'Google.Protobuf, Microsoft.ML, Microsoft.ML.CpuMath, Microsoft.ML.DataView, Microsoft.ML.OnnxConverter, and/or System.Numerics.Tensors packages were not found. Please install the Microsoft.ML.OnnxConverter package and its dependencies and then try again.' + [System.Environment]::NewLine + 'You can install the Microsoft.ML.OnnxConverter package and its dependencies by running the following command:' + [System.Environment]::NewLine + 'Install-Package -ProviderName NuGet -Name ''Microsoft.ML.OnnxConverter'' -Force -Scope CurrentUser;' + [System.Environment]::NewLine + [System.Environment]::NewLine
$hashtableCustomNotInstalledMessageToPackageNames.Add($strAccordMachineLearningNotInstalledMessage, @('Google.Protobuf', 'Microsoft.ML', 'Microsoft.ML.CpuMath', 'Microsoft.ML.DataView', 'Microsoft.ML.OnnxConverter', 'System.Numerics.Tensors'))
$refhashtableCustomNotInstalledMessageToPackageNames = [ref]$hashtableCustomNotInstalledMessageToPackageNames

$boolResult = Test-PackageInstalledUsingHashtable -ReferenceToHashtableOfInstalledPackages $refHashtablePackageNameToInstalledPackages -ThrowWarningIfPackageNotInstalled -ReferenceToHashtableOfSkippingDependencies $refHashtablePackageNameToSkippingDependencies -ReferenceToHashtableOfCustomNotInstalledMessages $refhashtableCustomNotInstalledMessageToPackageNames
if ($boolResult -eq $false) {
    return # Quit script
}
#endregion Make sure that required NuGet packages are installed #######################

#region Load the Microsoft.ML NuGet Package DLLs #####################################
$hashtablePackageNameToEffectiveDotNETVersions = @{}
$refHashtablePackageNameToEffectiveDotNETVersions = [ref]$hashtablePackageNameToEffectiveDotNETVersions
$hashtablePackageNameToDLLPaths = @{}
$refHashtablePackageNameToDLLPaths = [ref]$hashtablePackageNameToDLLPaths
Get-DLLPathsForPackagesUsingHashtable -ReferenceToHashtableOfInstalledPackages $refHashtablePackageNameToInstalledPackages -ReferenceToHashtableOfEffectiveDotNETVersions $refHashtablePackageNameToEffectiveDotNETVersions -ReferenceToHashtableOfDLLPaths $refHashtablePackageNameToDLLPaths

$arrNuGetPackages = @(
    'Google.Protobuf',
    'Microsoft.ML',
    'Microsoft.ML.CpuMath',
    'Microsoft.ML.DataView',
    'Microsoft.ML.OnnxConverter',
    'System.Numerics.Tensors'
)

$arrDLLPaths = @()

foreach ($strPackageName in $arrNuGetPackages) {
    if ($hashtablePackageNameToDLLPaths.Item($strPackageName).Count -gt 0) {
        foreach ($strDLLPath in $hashtablePackageNameToDLLPaths.Item($strPackageName)) {
            $arrDLLPaths += $strDLLPath

            # Load the .dll
            Write-Information ('Loading .dll: "' + $strDLLPath + '"')
            try {
                Add-Type -Path $strDLLPath
            } catch {
                $strMessage = 'Error loading .dll: "' + $strDLLPath + '"; the LoaderException(s) are: '
                $_.Exception.LoaderExceptions | ForEach-Object { $strMessage += $_.Message + '; ' }
                Write-Warning $strMessage
                return
            }
        }
    }
}
#endregion Load the Microsoft.ML NuGet Package DLLs #####################################

#region Validate input parameters ##################################################
# Validate CSV file
if (-not (Test-Path $strLabeledDataPath)) {
    Write-Warning "CSV file not found: $strLabeledDataPath"
    return
}

# Read CSV to inspect structure
$arrCSV = Import-Csv -Path $strLabeledDataPath
if ($arrCSV.Count -eq 0) {
    Write-Warning "The CSV file is empty: $strLabeledDataPath"
    return
}

# Get column names and validate
$arrColumns = @($arrCSV[0].PSObject.Properties.Name)
$intTextDataColumnIndex = $arrColumns.IndexOf($strColumnNameForTextData)
$intLabelColumnIndex = $arrColumns.IndexOf($strColumnNameForLabel)
if ($intTextDataColumnIndex -eq -1) {
    Write-Warning "Column '$strColumnNameForTextData' not found in CSV. Available columns: $($arrColumns -join ', ')"
    return
}
if ($intLabelColumnIndex -eq -1) {
    Write-Warning "Column '$strColumnNameForLabel' not found in CSV. Available columns: $($arrColumns -join ', ')"
    return
}

# Check for non-empty data
$validRows = $arrCSV | Where-Object { $_.$strColumnNameForTextData -and $_.$strColumnNameForLabel }
if ($validRows.Count -eq 0) {
    Write-Warning "No rows with non-empty '$strColumnNameForTextData' and '$strColumnNameForLabel' found in CSV."
    return
}
Write-Verbose -Message "Found $($validRows.Count) valid rows in CSV."
#endregion Validate input parameters ##################################################

$MLContext = [Microsoft.ML.MLContext]::New()

#region Load the CSV data into a Microsoft.ML.Data.DataView ########################
$strPathToDotNETStandardDLL = Join-Path ([System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) 'netstandard.dll'

$arrAssemblyReferences = @($hashtablePackageNameToDLLPaths['Microsoft.ML']) + $strPathToDotNETStandardDLL

$strCSharpCode = @"
namespace TicketClassification {
    public class Ticket {
        [Microsoft.ML.Data.LoadColumn($intTextDataColumnIndex)] public string TextData { get; set; }
        [Microsoft.ML.Data.LoadColumn($intLabelColumnIndex)] public string Label { get; set; }
    }
    public class TicketPrediction {
        [Microsoft.ML.Data.ColumnName("PredictedLabel")] public string PredictedLabel { get; set; }
    }
}
"@
try {
    Add-Type -TypeDefinition $strCSharpCode -Language CSharp -ReferencedAssemblies $arrAssemblyReferences
} catch {
    Write-Warning "Error compiling C# class: $($_.Exception.Message)"
    return
}

$records = [System.Collections.Generic.List[TicketClassification.Ticket]]::new()

$arrCSV | ForEach-Object {
    $ticket = [TicketClassification.Ticket]::new()
    # Do some text normalization since we had to disable these text normalization
    # features in the ML pipeline in order to allow for a clean export to ONNX
    if ([string]::IsNullOrEmpty($_.$strColumnNameForTextData)) {
        Write-Warning "Skipping row with empty text data: $($_)"
    } elseif ([string]::IsNullOrEmpty($_.$strColumnNameForLabel)) {
        Write-Warning "Skipping row with empty label: $($_)"
    } else {
        $ticket.TextData = $_.$strColumnNameForTextData | ConvertTo-NewlinelessText |
            ConvertTo-UnaccentedText | ConvertTo-UnpunctuatedText |
            ConvertTo-NumberlessText
        $ticket.Label = $_.$strColumnNameForLabel
        $records.Add($ticket) | Out-Null
    }
}

# Find the generic 2‑parameter overload (IEnumerable<T>, SchemaDefinition):
$methodInfo = $MLContext.Data.GetType().GetMethods() |
    Where-Object {
        $_.Name -eq 'LoadFromEnumerable' -and $_.IsGenericMethodDefinition -and
        $_.GetParameters().Count -eq 2
    } |
    Select-Object -First 1

# Bind the generic parameter:
$generic = $methodInfo.MakeGenericMethod([TicketClassification.Ticket])

# Invoke – 2nd argument being $null lets ML.NET infer the schema
$data = $generic.Invoke($MLContext.Data, @($records, $null))
#endregion Load the CSV data into a Microsoft.ML.Data.DataView ########################

# Code if you want to review a sample of the imported data:
# $preview = [Microsoft.ML.DebuggerExtensions]::Preview($data, 5)
# $preview.RowView | Format-Table

# Split into train / test:
$split = $MLContext.Data.TrainTestSplit($data, 0.2)

#region Build the training pipeline ################################################

#region Pipeline tester ############################################################
# I ran into issues building a pipeline and exporting it to ONNX. The following code
# helps with testing a pipeline and figuring out where it is failing.

# $steps = [System.Collections.Generic.List[object]]::new()
# $steps.Add([Microsoft.ML.TextCatalog]::NormalizeText($MLContext.Transforms.Text, 'TextNorm', 'TextData', [Microsoft.ML.Transforms.Text.TextNormalizingEstimator+CaseMode]::Lower, $true, $true, $true))
# $steps.Add([Microsoft.ML.TextCatalog]::TokenizeIntoWords($MLContext.Transforms.Text, 'Tokens', 'TextNorm'))
# $steps.Add([Microsoft.ML.TextCatalog]::RemoveDefaultStopWords($MLContext.Transforms.Text, 'Tokens', 'Tokens'))
# #####################################################################################
# # These steps are disabled because they are not supported when exporting to ONNX:
# #   $steps.Add([Microsoft.ML.ConversionsExtensionsCatalog]::MapValueToKey($MLContext.Transforms.Conversion, 'TokensKey', 'Tokens'))
# #   $steps.Add([Microsoft.ML.TextCatalog]::ProduceHashedNGrams($MLContext.Transforms.Text, 'Features', 'TokensKey', 16, 2, 0, $true, 314489979, $false, 0, $false))
# #####################################################################################
# $steps.Add([Microsoft.ML.TextCatalog]::ProduceWordBags(
#         $MLContext.Transforms.Text,
#         'Features',                # output vector
#         'Tokens',                  # input
#         2,                         # ngramLength = 2 (bi‑grams)
#         0,                         # skipLength
#         $true,                     # useAllLengths (uni + bi)
#         30000,                     # cap vocabulary
#         [Microsoft.ML.Transforms.Text.NgramExtractingEstimator+WeightingCriteria]::TfIdf))
# $steps.Add([Microsoft.ML.ConversionsExtensionsCatalog]::MapValueToKey($MLContext.Transforms.Conversion, 'LabelKey', 'Label'))
# $steps.Add([Microsoft.ML.StandardTrainersCatalog]::LbfgsMaximumEntropy($MLContext.MulticlassClassification.Trainers, 'LabelKey', 'Features'))
# $steps.Add([Microsoft.ML.ConversionsExtensionsCatalog]::MapKeyToValue($MLContext.Transforms.Conversion, 'PredictedLabel', 'PredictedLabel'))

# $result = Test-PipelineStepONNXExportable -MLContext $MLContext -SampleData $data -PipelineSteps $steps -OpSet 12
# $result | Format-List
#endregion Pipeline tester ############################################################

# Normalize text:
# Parameter order:
#   - catalog
#   - outputColumnName
#   - inputColumnName
#   - caseMode (we're converting to lower case)
#   - keepDiacritics (true = keep diacritics)
#   - keepPunctuations (true = keep punctuation)
#   - keepNumbers (true = keep numbers)
# Note: we have to use $true for keepDiacritics, keepPunctuations, and keepNumbers to
#       allow for a clean export to ONNX. So, instead, we pre-processed the text data
#       above as it was read from CSV using the ConvertTo-UnaccentedText, ConvertTo-UnpunctuatedText,
#       and ConvertTo-NumberlessText functions
$norm = [Microsoft.ML.TextCatalog]::NormalizeText($MLContext.Transforms.Text, 'TextNorm', 'TextData', [Microsoft.ML.Transforms.Text.TextNormalizingEstimator+CaseMode]::Lower, $true, $true, $true)

$tokens = [Microsoft.ML.TextCatalog]::TokenizeIntoWords($MLContext.Transforms.Text, 'Tokens', 'TextNorm')
# Possible future consideration:
# Optional consideration from LLM critique: Add a fixed truncation length (helps a
# little at conversion time (QNN parses the entire graph), and it guarantees
# predictable latency for very long tickets, but the classifier itself sees only the
# fixed‑size hashed vector):
#   - The hashed bag‑of‑words representation itself is shape‑agnostic, but QNN’s parser
#     still walks the upstream nodes. Clipping sentences to, say, 256 tokens keeps the
#     longest path bounded:
# NOTE: This code does not work - it would need to be fixed if we were implementing it:
# $tokens = [Microsoft.ML.TextCatalog]::TokenizeIntoWords(
#              $MLContext.Transforms.Text,
#              'Tokens','TextNorm',[Microsoft.ML.TextOptions+WordTokenizerOptions]::None,
#              256)   # maxTokens

# Remove stop words (this improves the model’s accuracy because stop words - e.g., a,
# an, and, the - are not useful for classification):
$stop = [Microsoft.ML.TextCatalog]::RemoveDefaultStopWords($MLContext.Transforms.Text, 'Tokens', 'Tokens')

#region Disabled steps required to allow for ONNX export ###########################
# The ProduceHashedNGrams step is not supported when exporting to ONNX. And since that
# step is disabled, there is no need to use the MapValueToKey step for the tokens. So,
# we will not be using the MapValueToKey step for the tokens.
#
#######################################################################################
# Convert words to categorical IDs:
# $tokensKey = [Microsoft.ML.ConversionsExtensionsCatalog]::MapValueToKey($MLContext.Transforms.Conversion, 'TokensKey', 'Tokens')

# # Tune the hashed‑n‑gram stage for an edge NPU:
# #   - 65 536 dimensions keeps the dense weight matrix under 256 k × 65 536 ≈ 16 MB in
# #     FP32, well inside the X Elite NPU’s SRAM window; bigger vectors may spill to
# #     DRAM and lose speed.
# #   - If you need very small models, drop to 14 bits (16 384 dims) and quantize to
# #     INT8 after export.
# # Parameter order:
# #   - catalog
# #   - outputColumnName
# #   - inputColumnName
# #   - numberOfBits (2^16 = 65,536 dims)
# #   - ngramLength (2 = unigram + bigram))
# #   - skipLength
# #   - useAllLengths ($true = include unigrams)
# #   - seed ([Microsoft.ML.Transforms.HashingEstimator]::DefaultSeed)
# #   - useOrderedHashing
# #   - maximumNumberOfInverts
# #   - rehashUnigrams
# $seed = [Microsoft.ML.Transforms.HashingEstimator]::DefaultSeed
# $ngrams = [Microsoft.ML.TextCatalog]::ProduceHashedNGrams($MLContext.Transforms.Text, 'Features', 'TokensKey', 16, 2, 0, $true, $seed, $false, 0, $false)
#endregion Disabled steps required to allow for ONNX export ###########################

# The ProduceHashedNGrams step is not supported when exporting to ONNX, but the
# ProduceWordBags step is supported. So, we are using ProduceWordBags instead.
# Parameter order:
#   - catalog
#   - outputColumnName
#   - inputColumnName
#   - ngramLength (2 = unigram + bigram))
#   - skipLength (Maximum number of tokens to skip when constructing an n-gram.)
#   - useAllLengths ($true = include ngrams < ngramLength, i.e., include unigrams)
#   - maximumNgramsCount (30000 = cap vocabulary to 30,000 ngrams)
#   - weighting
#     ([Microsoft.ML.Transforms.Text.NgramExtractingEstimator+WeightingCriteria]::TfIdf
#     = The product of the term frequency and the inverse document frequency.)
$wordBags = [Microsoft.ML.TextCatalog]::ProduceWordBags($MLContext.Transforms.Text, 'Features', 'Tokens', 2, 0, $true, 30000, [Microsoft.ML.Transforms.Text.NgramExtractingEstimator+WeightingCriteria]::TfIdf)

$labelToKey = [Microsoft.ML.ConversionsExtensionsCatalog]::MapValueToKey($MLContext.Transforms.Conversion, 'LabelKey', 'Label')
$trainer = [Microsoft.ML.StandardTrainersCatalog]::LbfgsMaximumEntropy($MLContext.MulticlassClassification.Trainers, 'LabelKey', 'Features')
# The following trainer is not supported by QNN:
# $trainer = [Microsoft.ML.StandardTrainersCatalog]::SdcaMaximumEntropy($MLContext.MulticlassClassification.Trainers, 'LabelKey', 'Features')
$keyToValue = [Microsoft.ML.ConversionsExtensionsCatalog]::MapKeyToValue($MLContext.Transforms.Conversion, 'PredictedLabel', 'PredictedLabel')

# Build the pipeline
# Start with the normalizer:
$pipelineCore = $norm
# Loop through and add the other steps:
# foreach ($step in @($tokens, $stop, $tokensKey, $ngrams, $labelToKey, $trainer, $keyToValue)) {
foreach ($step in @($tokens, $stop, $wordBags, $labelToKey, $trainer, $keyToValue)) {
    $pipelineCore = [Microsoft.ML.LearningPipelineExtensions]::Append($pipelineCore, $step, [Microsoft.ML.Data.TransformerScope]::Everything)
}
#endregion Build the training pipeline ################################################

#region Train the ML model #########################################################
# Train the model:
$modelCore = $pipelineCore.Fit($split.TrainSet)

$metrics = $MLContext.MulticlassClassification.Evaluate($modelCore.Transform($split.TestSet), 'LabelKey')
Write-Verbose -Message ('ML Model Macro Accuracy: ' + $metrics.MacroAccuracy)
Write-Verbose -Message ('ML Model Micro Accuracy: ' + $metrics.MicroAccuracy)
Write-Verbose -Message ('ML Model Log Loss: ' + $metrics.LogLoss)
Write-Verbose -Message ('ML Model Log Loss Reduction: ' + $metrics.LogLossReduction)
#endregion Train the ML model #########################################################

#region Save the ML model to an ONNX file ##########################################
# Set the OpSet (version) number:
#   - Using OpSet 12 because it is the maximum supported by Microsoft.ML.OnnxConverter
#     0.22.2
#   - OpSet 12 is fully sufficient for linear pipelines like that which we are using
#     here
#   - OpSet 12 is supported by QNNExecutionProvider
#   - There is no performance or accuracy gain expected from exporting this model at
#     OpSet 17, 19, or 21.
$opSet = 12
$stream = [System.IO.File]::Create($strONNXFileOutputPath)
try {
    # Export the ML model to ONNX format:
    # Parameters:
    #   - catalog
    #   - transform
    #   - inputData
    #   - opSetVersion
    #   - stream
    [Microsoft.ML.OnnxExportExtensions]::ConvertToOnnx($MLContext.Model, $modelCore, $data, $opSet, $stream)
} finally {
    $stream.Dispose()
}
#endregion Save the ML model to an ONNX file ##########################################
