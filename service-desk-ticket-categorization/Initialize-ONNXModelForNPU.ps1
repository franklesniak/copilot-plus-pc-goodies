# TODO: add function header

# Run this script on an x86-x64 (Intel or AMD) dev box
# to prepare the ONNX model for NPU quantization.
# This script is designed to be run in a PowerShell environment.
# It installs the necessary Python packages, checks for the ONNX model file,
# and attempts to fix the input shape of the model if necessary.
# It also handles the output file and prepares the model for the next step.
# Ensure you have Python and pip installed on your system.

#region IMPORTANT -- Prereqs #######################################################
# In the registry key:
# HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows Defender\Policy Manager
# Set the value of ASRRules to the following string:
# 56a863a9-875e-4185-98a7-b882c64b5ce5=1|7674ba52-37eb-4a4f-a9a1-f0f9a1619a2c=1|d4f940ab-401b-4efc-aadc-ad5f3c50688a=1|9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2=1|be9ba2d9-53ea-4cdc-84e5-9b1eeee46550=1|01443614-cd74-433a-b99e-2ecdc07bfc25=2|5beb7efe-fd9a-4556-801d-275e5ffc04cc=1|d3e037e1-3eb8-44c8-a917-57927947596d=1|3b576869-a4ec-4529-8536-b80a7769e899=1|75668c1f-73b5-4cf0-bb93-3ecf5cb7cc84=1|26190899-1602-49e8-8b27-eb1d0a1ce869=1|e6db77e5-3df2-4cf1-b95a-636979351e5b=1|d1e49aac-8f56-4280-b9ba-993a6d77406c=1|33ddedf1-c6e0-47cb-833e-de6133960387=1|b2b3f03d-6a65-4f7b-a9c7-1c7ef74a9ba4=1|c0033c00-d16d-4114-a5a0-dc9b3a7d2ceb=2|a8f5898e-1dc8-49a9-9878-85004b8a61e6=1|92e97fa1-2edf-4476-bdd6-9dd0b4dddc7b=1|c1db55ab-c21a-4637-bb3f-a12568109d35=1
# When you're done, you can set it back to the following:
# 56a863a9-875e-4185-98a7-b882c64b5ce5=1|7674ba52-37eb-4a4f-a9a1-f0f9a1619a2c=1|d4f940ab-401b-4efc-aadc-ad5f3c50688a=1|9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2=1|be9ba2d9-53ea-4cdc-84e5-9b1eeee46550=1|01443614-cd74-433a-b99e-2ecdc07bfc25=1|5beb7efe-fd9a-4556-801d-275e5ffc04cc=1|d3e037e1-3eb8-44c8-a917-57927947596d=1|3b576869-a4ec-4529-8536-b80a7769e899=1|75668c1f-73b5-4cf0-bb93-3ecf5cb7cc84=1|26190899-1602-49e8-8b27-eb1d0a1ce869=1|e6db77e5-3df2-4cf1-b95a-636979351e5b=1|d1e49aac-8f56-4280-b9ba-993a6d77406c=1|33ddedf1-c6e0-47cb-833e-de6133960387=1|b2b3f03d-6a65-4f7b-a9c7-1c7ef74a9ba4=1|c0033c00-d16d-4114-a5a0-dc9b3a7d2ceb=2|a8f5898e-1dc8-49a9-9878-85004b8a61e6=1|92e97fa1-2edf-4476-bdd6-9dd0b4dddc7b=1|c1db55ab-c21a-4637-bb3f-a12568109d35=1
#
# You'll also need to set the value of "AllowCloudProtection" to 0 in the same registry
# key. When you're done, you can set it back to 1.
#
# And "SubmitSamplesConsent" to 2. When you're done, you can set it back to 3.
#
# At the time of writing, the onnx Python module, on which the olive-ai module depends,
# does not work with Python 3.13 or newer; use Python 3.12 or older. Review the release notes
# here to see if a newer Python version is supported:
# https://github.com/onnx/onnx/releases
# Also review the QNN SDK setup instructions for Windows to see what version of Python
# is supported. At the time of writing, the QNN SDK requires Python 3.10. The
# instructions for the QNN SDK are here:
# https://docs.qualcomm.com/bundle/publicresource/topics/80-63442-50/windows_setup.html
# Python 3.10 can be installed side by side with other versions of Python.
# You can download Python 3.10 from: https://github.com/adang1345/PythonWindows
#
# You need to have the Qualcomm Neural Processing SDK for AI downloaded and extracted
# See: https://www.qualcomm.com/developer/software/neural-processing-sdk-for-ai
# Once you have it downloaded and unzipped, you need to configure it. See:
# https://docs.qualcomm.com/bundle/publicresource/topics/80-63442-50/windows_setup.html
#
# Once you have the SDK, set the environment variable QNN_SDK_ROOT to the path of the
# SDK folder. The SDK folder should contain the following subfolders:
#   - bin
#   - include
#   - lib
#   (etc.)
# to set the environment variable, you can use the following command in PowerShell:
# $env:QNN_SDK_ROOT = "C:\Users\flesniak\Downloads\v2.34.0.250424\qairt\2.34.0.250424"
#
# NOTE: During this script, we will be downloading and installing the Qualcomm Hexagon
# SDK.
#endregion IMPORTANT -- Prereqs #######################################################

[CmdletBinding()]
param()

$PathToPython310 = 'C:\Users\flesniak\AppData\Local\Programs\Python\Python310\python.exe'
$strLabeledDataPath = 'C:\Users\flesniak\Downloads\Service-Desk-Tickets\Project-Phoenix-Incidents-Labeled.csv'
$strCalibrationOutputFilePath = 'C:\Users\flesniak\Downloads\Service-Desk-Tickets\Project-Phoenix-Incidents-Calibration-Inputs.txt'
$strColumnNameForTextData = 'Description'
$inputFile = 'C:\Users\flesniak\Downloads\Service-Desk-Tickets\Project-Phoenix-Incidents-Labeled.onnx'
$outputFile = 'C:\Users\flesniak\Downloads\Service-Desk-Tickets\Project-Phoenix-Incidents-Labeled-Fixed.onnx'
$inputNodeName = 'TextData'
$PathToOliveJSON = 'C:\Users\flesniak\Downloads\Service-Desk-Tickets\olive-qnn.json'
$targetPercentage = 0.10 # Target 10% of the valid data for calibration
$maxSamples = 1000       # Maximum number of samples to use
$minSamples = 100        # Minimum number of samples to aim for (if dataset is very small)
$QualcommUsername = 'franklesniak@domain.com'
$QualcommPassword = 'StoreThisInAKeyVaultAndUseSecretsManagementModule'

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

if ([string]::IsNullOrEmpty($env:QNN_SDK_ROOT)) {
    Write-Error -Message "Environment variable QNN_SDK_ROOT is not set. Please set it to the path of the QNN SDK folder."
    return
}
if (Test-Path (Join-Path $env:QNN_SDK_ROOT 'bin')) {
    Write-Verbose -Message "QNN SDK bin folder found: $($env:QNN_SDK_ROOT)\bin"
} else {
    Write-Error -Message "QNN SDK bin folder not found at $($env:QNN_SDK_ROOT)\bin. Please check that you have the environment variable QNN_SDK_ROOT set correctly."
    return
}
$env:PATH = "$env:QNN_SDK_ROOT\bin;$(Split-Path $PathToPython310);$(Join-Path (Split-Path $PathToPython310) 'Scripts');$env:PATH"

#region Create an isolated env that the QNN SDK and olive-ai can share #############
$strUserProfile = (Get-Item '~').FullName
$strPythonEnvsFolder = Join-Path $strUserProfile 'envs'
if (-not (Test-Path $strPythonEnvsFolder)) {
    New-Item -Path $strPythonEnvsFolder -ItemType Directory -Force | Out-Null
}
$strOliveQNNEnvFolder = Join-Path $strPythonEnvsFolder 'olive-qnn'
if (-not (Test-Path $strOliveQNNEnvFolder)) {
    New-Item -Path $strOliveQNNEnvFolder -ItemType Directory -Force | Out-Null
}
$strCommand = 'py -3.10 -m venv "' + $strOliveQNNEnvFolder + '"'
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand

$strPythonEnvsScriptFolder = Join-Path $strOliveQNNEnvFolder 'Scripts'
$strActivationScriptPath = Join-Path $strPythonEnvsScriptFolder 'Activate.ps1'
if (-not (Test-Path $strActivationScriptPath)) {
    Write-Error -Message "Activation script not found at $strActivationScriptPath. Please check the path."
    return
}
$strCommand = '& ' + $strActivationScriptPath
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand

$strCommand = (Join-Path $strPythonEnvsScriptFolder 'python.exe') + ' -m pip install --upgrade pip'
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand
#endregion Create an isolated env that the QNN SDK and olive-ai can share #############

#region Set up Qualcomm SDK ########################################################
$strCommand = 'python "' + (Join-Path (Join-Path $env:QNN_SDK_ROOT 'bin') 'check-python-dependency') + '"'
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand

$strCommand = 'pip install torch==1.13.1'
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand

$strCommand = 'pip install onnx==1.12.0'
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand

$strCommand = 'pip install onnxruntime==1.17.1'
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand

$strCommand = 'pip install onnxsim==0.4.36'
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand

$strCommand = 'qik LOGIN ' + $QualcommUsername + ' ' + $QualcommPassword
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand

$strCommand = 'qik LICENSE ACTIVATE hexagonsdk6.x'
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand

$strCommand = 'qik LICENSE ACTIVATE compute1.x'
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand

$strCommand = 'qik install hexagonsdk6.x'
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand

$strCommand = 'qik install compute1.x'
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand

$strPathToHexagonSDK = Get-ChildItem 'C:\Qualcomm\Hexagon_SDK' | Sort-Object -Descending | Select-Object -First 1 | ForEach-Object { $_.FullName }
$strPathToHexagonSetupScript = Join-Path $strPathToHexagonSDK 'setup_sdk_env.cmd'
$strCommand = $strPathToHexagonSetupScript
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand
#endregion Set up Qualcomm SDK ########################################################

$strCommand = 'pip install --upgrade olive-ai[qnn]'
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand

$actionPreferencePreviousProgress = $ProgressPreference
$ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
$strCommand = 'olive configure-qualcomm-sdk --py_version 3.8 --sdk qnn'
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand
$ProgressPreference = $actionPreferencePreviousProgress

#region Validate input parameters ##################################################
# Validate CSV file
if (-not (Test-Path $strLabeledDataPath)) {
    Write-Warning "CSV file not found: $strLabeledDataPath"
    return
}

Write-Verbose -Message "Starting calibration data preparation..."
Write-Verbose -Message "Reading CSV file: $strLabeledDataPath"

# Read CSV to inspect structure
$allRecords = Import-Csv -Path $strLabeledDataPath
if ($allRecords.Count -eq 0) {
    Write-Warning "The CSV file is empty: $strLabeledDataPath"
    return
}

# # Get column names and validate
# $arrColumns = @($allRecords[0].PSObject.Properties.Name)
# $intTextDataColumnIndex = $arrColumns.IndexOf($strColumnNameForTextData)
# if ($intTextDataColumnIndex -eq -1) {
#     Write-Warning "Column '$strColumnNameForTextData' not found in CSV. Available columns: $($arrColumns -join ', ')"
#     return
# }

# Check for non-empty data
Write-Verbose -Message "Filtering records with non-empty text in column '$strColumnNameForTextData'..."
$validRecords = $allRecords | Where-Object { -not [string]::IsNullOrWhiteSpace($_.$strColumnNameForTextData) }
$totalValidRecords = $validRecords.Count

if ($totalValidRecords -eq 0) {
    Write-Error -Message "No records with valid text data found in column '$strColumnNameForTextData'. Cannot create calibration file."
    return
}
Write-Verbose -Message "Found $($validRows.Count) valid rows in CSV."
#endregion Validate input parameters ##################################################

# Calculate the number of samples to take
$desiredCount = [int]($totalValidRecords * $targetPercentage)
$numSamplesToTake = [System.Math]::Min($maxSamples, [System.Math]::Max($minSamples, $desiredCount))

# Ensure we don't request more samples than available
$numSamplesToTake = [System.Math]::Min($numSamplesToTake, $totalValidRecords)

Write-Verbose -Message "Targeting $($targetPercentage * 100)% of data, min $minSamples, max $maxSamples."
Write-Verbose -Message "Selecting $numSamplesToTake random samples for calibration."

# Select random samples
# Use a temporary variable to avoid potential pipeline issues with large datasets in memory
$selectedSamples = $validRecords | Get-Random -Count $numSamplesToTake

# Prepare the output file (clear if it exists)
if (Test-Path $strCalibrationOutputFilePath) {
    Write-Verbose -Message "Clearing existing calibration file: $strCalibrationOutputFilePath"
    Clear-Content -Path $strCalibrationOutputFilePath
}

Write-Verbose -Message "Processing selected samples and writing to calibration file..."
$count = 0
foreach ($record in $selectedSamples) {
    $textData = $record.$strColumnNameForTextData

    # Pre-process: Apply ML.NET pre-processing steps used outside the ONNX graph
    $processedText = $textData | ConvertTo-NewlinelessText |
        ConvertTo-UnaccentedText | ConvertTo-UnpunctuatedText |
        ConvertTo-NumberlessText

    # Write the single line of processed text to the file
    Add-Content -Path $strCalibrationOutputFilePath -Value $processedText
    $count++
}

Write-Verbose "Successfully created calibration file '$strCalibrationOutputFilePath' with $count samples."

$strCommand = 'olive run --config ' + $PathToOliveJSON
$scriptblockCommand = [scriptblock]::Create($strCommand)
& $scriptblockCommand















# This command assumes the input is a single string (batch size 1).
# The shape [1] represents a single element (the input string).
$fixShapeCommand = "python -m onnxruntime.tools.convert_dynamic_shape --input_model $inputFile --output_model $outputFile --input_dims $inputNodeName 1"
$fixShapeCommand = "python -m onnxruntime.tools.make_dynamic_shape_fixed --dim_param $inputNodeName --dim_value 1 $inputFile $outputFile"

# Execute the command
Write-Host "Running command: $fixShapeCommand"
cmd /k $fixShapeCommand
Write-Host "Attempted to fix input shape. Output potentially saved to: $outputFile"
if (-not (Test-Path $outputFile)) {
     Write-Warning "Output file $outputFile was not created. The model might already have fixed shapes, or an error occurred."
     # If it didn't create a new file, use the original for the next step.
     $modelToQuantize = $inputFile
} else {
     $modelToQuantize = $outputFile
}
Write-Host "Using model for next step: $modelToQuantize"