# .SYNOPSIS
# Demonstrates how to query the NPU (Neural Processing Unit) engine metadata and usage.
#
# .DESCRIPTION
# This script queries the NPU engine metadata and usage using the DXCore API.
#
# For at least Qualcomm and AMD NPUs, it gathers metadata such as the adapter
# description and and memory budget. For Qualcomm, AMD, and Intel NPUs, this script
# also retrieves the NPU engine usage in percentage.
#
# .EXAMPLE
# & .\Get-NPUMetadataAndUsage.ps1
# # This will gather NPU metadata, at least for Qualcomm/AMD NPUs. Then, it will show
# # the NPU usage.
#
# .EXAMPLE
# . .\Get-NPUMetadataAndUsage.ps1
# while ($true) { Get-NPUUtilization }
# # This will continuously query the NPU utilization and print it to the console.
#
# .INPUTS
# None. You can't pipe objects to Get-NPUMetadataAndUsage.
#
# .OUTPUTS
# None. The function writes information to the host (console), but does not return any
# objects.
#
# .NOTES
# Intel's NPU driver does not currently conform to the DXCore API, so it is not
# possible to query the NPU engine metadata. The script uses a "hack" workaround to
# retrieve the NPU engine metadata for Intel NPUs. This workaround may not be
# reliable in the future, as it relies on the specific implementation of the Intel
# NPU driver.
#
# Version: 0.1.20250505.0

#region License ############################################################
# Copyright (c) 2025 Frank Lesniak
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
#endregion License ############################################################

[CmdletBinding()]
param()

#region Define helper functions ####################################################
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

function Test-Windows {
    # .SYNOPSIS
    # Returns $true if PowerShell is running on Windows; otherwise, returns
    # $false.
    #
    # .DESCRIPTION
    # Returns a boolean ($true or $false) indicating whether the current
    # PowerShell session is running on Windows. This function is useful for
    # writing scripts that need to behave differently on Windows and non-
    # Windows platforms (Linux, macOS, etc.). Additionally, this function is
    # useful because it works on Windows PowerShell 1.0 through 5.1, which do
    # not have the $IsWindows global variable.
    #
    # .PARAMETER PSVersion
    # This parameter is optional; if supplied, it must be the version number of
    # the running version of PowerShell. If the version of PowerShell is
    # already known, it can be passed in to this function to avoid the overhead
    # of unnecessarily determining the version of PowerShell. If this parameter
    # is not supplied, the function will determine the version of PowerShell
    # that is running as part of its processing.
    #
    # .EXAMPLE
    # $boolIsWindows = Test-Windows
    #
    # .EXAMPLE
    # # The version of PowerShell is known to be 2.0 or above:
    # $boolIsWindows = Test-Windows -PSVersion $PSVersionTable.PSVersion
    #
    # .INPUTS
    # None. You can't pipe objects to Test-Windows.
    #
    # .OUTPUTS
    # System.Boolean. Test-Windows returns a boolean value indiciating whether
    # PowerShell is running on Windows. $true means that PowerShell is running
    # on Windows; $false means that PowerShell is not running on Windows.
    #
    # .NOTES
    # This function also supports the use of a positional parameter instead of
    # a named parameter. If a positional parameter is used instead of a named
    # parameter, then one positional parameter is required: it must be the
    # version number of the running version of PowerShell. If the version of
    # PowerShell is already known, it can be passed in to this function to
    # avoid the overhead of unnecessarily determining the version of
    # PowerShell. If this parameter is not supplied, the function will
    # determine the version of PowerShell that is running as part of its
    # processing.
    #
    # Version: 1.1.20250106.1

    #region License ########################################################
    # Copyright (c) 2025 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a
    # copy of this software and associated documentation files (the
    # "Software"), to deal in the Software without restriction, including
    # without limitation the rights to use, copy, modify, merge, publish,
    # distribute, sublicense, and/or sell copies of the Software, and to permit
    # persons to whom the Software is furnished to do so, subject to the
    # following conditions:
    #
    # The above copyright notice and this permission notice shall be included
    # in all copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
    # OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
    # MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN
    # NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
    # DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
    # OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE
    # USE OR OTHER DEALINGS IN THE SOFTWARE.
    #endregion License ########################################################

    param (
        [version]$PSVersion = ([version]'0.0')
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

    if ($PSVersion -ne ([version]'0.0')) {
        if ($PSVersion.Major -ge 6) {
            return $IsWindows
        } else {
            return $true
        }
    } else {
        $versionPS = Get-PSVersion
        if ($versionPS.Major -ge 6) {
            return $IsWindows
        } else {
            return $true
        }
    }
}

function Get-GPUEnginePath {
    # full path list, Utilization counters only
    (Get-Counter -ListSet 'GPU Engine').PathsWithInstances |
        Where-Object { $_ -like '*\Utilization Percentage' }
}

function Get-AdapterGroup ($paths) {
    # group by the *full* LUID (two hex numbers)
    $rxLuid = 'luid_0x[0-9A-F]+_0x[0-9A-F]+'
    $paths | Group-Object {
        if ($_ -match $rxLuid) { $Matches[0] } else { 'noluid' }
    }
}

function Get-NPUPath {
    $all = Get-GPUEnginePath
    $byLuid = Get-AdapterGroup $all

    foreach ($g in $byLuid) {
        $engTypes = $g.Group |
            Where-Object { $_ -match 'engtype_([^)\\]+)' } |
            ForEach-Object { $Matches[1] } |
            Sort-Object -Unique

        # modern Intel / future-proof: NPU = any AI engine
        if ($engTypes -contains 'AI') {
            return $g.Group
        }

        # Qualcomm/AMD: NPU = standalone compute engine
        if (($engTypes | Where-Object { $_ -eq 'Compute' }).Count -eq $engTypes.Count) {
            return $g.Group
        }

        # early Intel preview: NPU tile exported *only* 3D engines
        if (($engTypes | Where-Object { $_ -eq '3D' }).Count -eq $engTypes.Count) {
            if (@($g.Group)[0] -match 'luid_0x([0-9A-F]+)_0x([0-9A-F]+)') {
                $uintLowLuid = [Convert]::ToUInt32($Matches[2], 16)
                $intHighLuid = [Convert]::ToInt32($Matches[1], 16)
                $strDescription = [NpuQuery]::DescribeAdapterByLuid($uintLowLuid, $intHighLuid)
                if ($strDescription -eq 'Microsoft Basic Render Driver') {
                    # Skip the Microsoft Basic Render Driver!
                } else {
                    return $g.Group
                }
            }
        }
    }
    return @()   # nothing matched
}

function Get-NPUUtilization {
    $npu = Get-NPUPath
    if (-not $npu) { Write-Warning 'No NPU engine counters found.'; return [double]::NaN }

    (Get-Counter $npu -MaxSamples 1).CounterSamples |
        Measure-Object CookedValue -Sum |
        Select-Object -ExpandProperty Sum
}
#endregion Define helper functions ####################################################

$versionPS = Get-PSVersion
if ($versionPS.Major -lt 7) {
    Write-Warning 'This script requires PowerShell 7 or later.'
    return
}
$boolIsWindows = Test-Windows -PSVersion $versionPS
if (-not $boolIsWindows) {
    Write-Warning 'This script is only supported on Windows.'
    return
}

#region Create a DXCore C# wrapper so that we can access NPU usage from PowerShell #

#region C# Code Definition #########################################################
$CSharpSource = @"
using System;
using System.Runtime.InteropServices;
using System.Text; // Required for GetProperty with string results like Description
using System.Diagnostics; // For Debug.WriteLine

[ComImport]
[Guid("78EE5945-C36E-4B13-A669-005DD11C0F06")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IDXCoreAdapterFactory
{
    // v-tbl[3]
    [PreserveSig] int CreateAdapterList(
        uint  numAttributes,
        [In, MarshalAs(UnmanagedType.LPArray,
                       ArraySubType = UnmanagedType.LPStruct)]
        Guid[] filterAttributes,          // may be null when count == 0
        [In] ref Guid riid,
        [Out] out IntPtr ppvAdapterList);

    // v-tbl[4]
    [PreserveSig] int GetAdapterByLuid(
        [In] ref LUID luid,
        [In] ref Guid riid,
        [Out] out IntPtr ppvAdapter);

    // v-tbl[5]
    [return: MarshalAs(UnmanagedType.Bool)]
    [PreserveSig] bool IsAdapterAttributeSupported([In] ref Guid attribute);

    // v-tbl[6]
    [PreserveSig] int RegisterAdaptersChangedEvent(
        IntPtr hEvent,           // HANDLE
        out uint cookie);        // DXCoreNotificationCookie

    // v-tbl[7]
    [PreserveSig] void UnregisterAdaptersChangedEvent(uint cookie);
}

[ComImport]
[Guid("526C7776-40E9-459B-B711-F32AD76DFC28")]   //  IDXCoreAdapterList
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IDXCoreAdapterList
{
    [PreserveSig] int  GetAdapter(uint index, [In] ref Guid riid, [Out] out IntPtr ppvAdapter);
    [PreserveSig] uint GetAdapterCount();

    [return: MarshalAs(UnmanagedType.Bool)]
    [PreserveSig] bool IsStale();

    [PreserveSig] int  GetFactory([In] ref Guid riid, [Out] out IntPtr ppvFactory);
    [PreserveSig] int  Sort(uint numPreferences,
        [In, MarshalAs(UnmanagedType.LPArray)] DXCoreAdapterPreference[] preferences);

    [return: MarshalAs(UnmanagedType.Bool)]
    [PreserveSig] bool IsAdapterPreferenceSupported(DXCoreAdapterPreference preference);
}

[ComImport]
[Guid("F0DB4C7F-FE5A-42A2-BD62-F2A6CF6FC83E")]   //  IDXCoreAdapter
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IDXCoreAdapter
{
    [return: MarshalAs(UnmanagedType.Bool)]
    [PreserveSig] bool IsValid();

    [return: MarshalAs(UnmanagedType.Bool)]
    [PreserveSig] bool IsAttributeSupported([In] ref Guid attributeGuid);

    [return: MarshalAs(UnmanagedType.Bool)]
    [PreserveSig] bool IsPropertySupported(DXCoreAdapterProperty property);

    [PreserveSig] int  GetProperty(DXCoreAdapterProperty property,
                                   nuint                 bufferSize,
                                   IntPtr                propertyData);
    [PreserveSig] int  GetPropertySize(DXCoreAdapterProperty property,
                                   out nuint               bufferSize);

    [return: MarshalAs(UnmanagedType.Bool)]
    [PreserveSig] bool IsQueryStateSupported(DXCoreAdapterState state);

    [PreserveSig] int  QueryState(DXCoreAdapterState state,
                                  nuint                inputStateDetailsSize,
                                  IntPtr               inputStateDetails, // Should be null if size is 0
                                  nuint                outputBufferSize,
                                  IntPtr               outputBuffer);

    [return: MarshalAs(UnmanagedType.Bool)]
    [PreserveSig] bool IsSetStateSupported(DXCoreAdapterState state);

    [PreserveSig] int  SetState(DXCoreAdapterState state,
                                nuint inputStateDetailsSize,
                                IntPtr inputStateDetails, // Should be null if size is 0
                                nuint inputDataSize,
                                IntPtr inputData); // Should be null if size is 0

    [PreserveSig] int  GetFactory([In] ref Guid riid, [Out] out IntPtr ppvFactory);
}

// --- Enums and Constants ---
public static class DXCoreGUIDs
{
    public static readonly Guid IID_IDXCoreAdapterList = new Guid("526C7776-40E9-459B-B711-F32AD76DFC28");
    public static readonly Guid IID_IDXCoreAdapter = new Guid("F0DB4C7F-FE5A-42A2-BD62-F2A6CF6FC83E");
    public static readonly Guid IID_IDXCoreAdapterFactory = new Guid("78EE5945-C36E-4B13-A669-005DD11C0F06");
    // ───────── Hardware-type attributes ─────────
    //  Source: https://github.com/MicrosoftDocs/win32/docs/desktop-src/dxcore/dxcore-adapter-attribute-guids.md
    public static readonly Guid DXCORE_HARDWARE_TYPE_ATTRIBUTE_NPU =
        new Guid("B69EB219-3DED-4464-979F-A00BD4687006");

    public static readonly Guid DXCORE_HARDWARE_TYPE_ATTRIBUTE_GPU =
        new Guid("E0B195DA-58EF-4A22-90F1-1F28169CAB8D");

    public static readonly Guid DXCORE_HARDWARE_TYPE_ATTRIBUTE_COMPUTE_ACCELERATOR =
        new Guid("D46140C4-ADD7-451B-9E56-06FE8C3B58ED");

    public static readonly Guid DXCORE_HARDWARE_TYPE_ATTRIBUTE_MEDIA_ACCELERATOR =
        new Guid("188F40E5-61D3-45AE-964F-6AE92A00EAFD");
    // ───────── Capability attributes (Direct3D-12 features) ─────────
    public static readonly Guid DXCORE_ADAPTER_ATTRIBUTE_D3D12_GENERIC_ML =
        new Guid("B71B0D41-1088-422F-A27C-0250B7D3A988");   // ML meta-commands
    public static readonly Guid DXCORE_ADAPTER_ATTRIBUTE_D3D12_GRAPHICS =
        new Guid("0C9ECE4D-2F6E-4F01-8C96-E89E331B47B1");   // “can do graphics”
    public static readonly Guid DXCORE_ADAPTER_ATTRIBUTE_D3D12_CORE_COMPUTE =
        new Guid("248E2800-A793-4724-ABAA-23A6DE1BE090");
}

// Keep Enums as they were - DXCoreAdapterProperty now includes IsNpu
public enum DXCoreAdapterProperty : uint { InstanceLuid = 0, DriverVersion = 1, DriverDescription = 2, HardwareID = 3, KmdModelVersion = 4, ComputePreemptionGranularity = 5, GraphicsPreemptionGranularity = 6, DedicatedAdapterMemory = 7, DedicatedSystemMemory = 8, SharedSystemMemory = 9, AcgCompatible = 10, IsHardware = 11, IsIntegrated = 12, IsDetachable = 13, HardwareIDParts = 14, PhysicalAdapterCount = 15, AdapterEngineCount = 16, AdapterSoftwareType = 17, IsNpu = 18 }
public enum DXCoreAdapterState : uint
{
    IsDriverUpdateInProgress                 = 0,
    AdapterMemoryBudget                      = 1,
    AdapterMemoryUsageBytes                  = 2,
    AdapterMemoryUsageByProcessBytes         = 3,
    AdapterEngineRunningTimeMicroseconds     = 4,
    AdapterEngineRunningTimeByProcessUS      = 5
    // … keep the rest if you later need them
}
public enum DXCoreNotificationType : uint { AdapterListStale = 0, AdapterNoLongerValid = 1, AdapterBudgetChange = 2, AdapterHardwareContentProtectionTeardown = 3 }
public enum DXCoreAdapterPreference : uint {
    Hardware = 0,
    MinimumPower = 1,
    HighPerformance = 2
}
public enum DXCoreSegmentGroup : uint
{
    Local = 0,
    NonLocal = 1
}

// --- Structs ---
[StructLayout(LayoutKind.Sequential)] internal struct LUID { internal uint LowPart; internal int HighPart; }
[StructLayout(LayoutKind.Sequential)] public struct DXCoreAdapterMemoryBudget { public ulong budget; public ulong currentUsage; public ulong availableForOS; public ulong availableForProcess; }
[StructLayout(LayoutKind.Sequential)] internal struct DXCoreAdapterMemoryBudgetNodeSegmentGroup
{
    public uint nodeIndex;              // 0 for single-adapter
    public DXCoreSegmentGroup segmentGroup;
}
[StructLayout(LayoutKind.Sequential)] internal struct DXCoreAdapterEngineIndex
{
    public uint physicalAdapterIndex;      // usually 0
    public uint engineIndex;               // NPU engine 0, 1 …
}
[StructLayout(LayoutKind.Sequential)]internal struct DXCoreEngineQueryOutput
{
    public ulong runningTime;              // µs since boot
    [MarshalAs(UnmanagedType.I1)] public bool processQuerySucceeded;
}

// --- P/Invoke for DXCoreCreateAdapterFactory ---
internal static class NativeMethods
{
    [DllImport("dxcore.dll", CharSet = CharSet.Unicode, SetLastError = true, CallingConvention = CallingConvention.StdCall)]
    internal static extern int DXCoreCreateAdapterFactory([In] ref Guid iid, [Out] out IntPtr ppvFactory);
}

// --- Main Query Class ---
public static class NpuQuery
{
    // Helper methods (GetAdapterProperty<T>, GetAdapterStringProperty, QueryAdapterState<T>)
    // These remain largely the same, but note the bool handling in GetAdapterProperty
    private static T? GetAdapterProperty<T>(IDXCoreAdapter adapter, DXCoreAdapterProperty property) where T : struct
    {
        if (!adapter.IsPropertySupported(property)) return null; // Check support first

        nuint propertySize;
        int hr = adapter.GetPropertySize(property, out propertySize);
        if (hr < 0 || propertySize == 0)
        {
            Debug.WriteLine("GetPropertySize failed (HRESULT: 0x" + hr.ToString("X8") + ") or returned 0 size for property " + property + ".");
            return null;
        }

        IntPtr buffer = Marshal.AllocHGlobal((int)propertySize);
        try
        {
            hr = adapter.GetProperty(property, propertySize, buffer);
            if (hr < 0)
            {
                Debug.WriteLine("GetProperty failed (HRESULT: 0x" + hr.ToString("X8") + ") for property " + property + ".");
                return null;
            }

            if (typeof(T) == typeof(bool))
            {
                 // For bool (marshaled as BOOL), check if the first byte is non-zero
                 if (propertySize >= 1) // Ensure there's at least one byte
                 {
                      byte value = Marshal.ReadByte(buffer);
                      return (T)(object)(value != 0);
                 }
                 else
                 {
                     Debug.WriteLine("GetProperty for bool property " + property + " returned unexpected size " + propertySize + ".");
                     return null; // Unexpected size
                 }
            }
            else
            {
                return Marshal.PtrToStructure<T>(buffer);
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine("Exception during GetAdapterProperty marshaling for " + property + ": " + ex.Message);
            return null;
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    private static string GetAdapterStringProperty(IDXCoreAdapter adp,
                                              DXCoreAdapterProperty prop)
    {
        if (!adp.IsPropertySupported(prop)) return "N/A (Unsupported)";

        nuint len;
        int hr = adp.GetPropertySize(prop, out len);
        if (hr < 0 || len == 0) return ("SizeErr 0x" + hr.ToString("X8"));

        IntPtr buf = Marshal.AllocHGlobal((int)len);
        try
        {
            hr = adp.GetProperty(prop, len, buf);
            if (hr < 0) return ("Err 0x" + hr.ToString("X8"));

            // Quick heuristic: does it *look* like UTF-16?
            bool looksUtf16 = true;

            // Copy unmanaged buffer into a managed array
            int length = checked((int)len);
            byte[] bytes = new byte[length];
            Marshal.Copy(buf, bytes, 0, length);

            // Work with managed data
            for (int i = 1; i < bytes.Length; i += 2)
            {
                if (bytes[i] != 0) { looksUtf16 = false; break; }
            }

            string s = looksUtf16
                    ? Marshal.PtrToStringUni(buf)!                     // UTF-16-LE
#if NET5_0_OR_GREATER
                    : Marshal.PtrToStringUTF8(buf, (int)len)!;         // UTF-8
#else
                    : Encoding.UTF8.GetString(bytes);  // UTF-8 (fallback for older .NET versions)
#endif
            return s.TrimEnd('\0');
        }
        finally { Marshal.FreeHGlobal(buf); }
    }

    private static T? QueryAdapterState<T>(IDXCoreAdapter adapter, DXCoreAdapterState state) where T : struct
    {
        if (!adapter.IsQueryStateSupported(state))
        {
            Debug.WriteLine("QueryState is not supported for state " + state + ".");
            return null;
        }

        nuint bufferSize = (nuint)Marshal.SizeOf<T>();
        IntPtr buffer = Marshal.AllocHGlobal((int)bufferSize);
        try
        {
            // Pass IntPtr.Zero for inputStateDetails as we are not providing any for AdapterMemoryBudget
            int hr = adapter.QueryState(state, (nuint)0, IntPtr.Zero, bufferSize, buffer);
            if (hr < 0)
            {
                Debug.WriteLine("QueryState failed for state " + state + " with HRESULT: 0x" + hr.ToString("X8"));
                return null;
            }
            return Marshal.PtrToStructure<T>(buffer);
        }
        catch (Exception ex)
        {
            Debug.WriteLine("Exception during QueryState marshaling for " + state + ": " + ex.Message);
            return null;
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    public class NpuQueryResult
    {
        public string AdapterDescription { get; set; }
        public bool   IsNpu              { get; set; } // Will be set true if detected
        public ulong  BudgetBytes        { get; set; }
        public ulong  CurrentUsageBytes  { get; set; }
        public ulong  AvailableForOSBytes{ get; set; }
        public ulong  AvailableForProcessBytes { get; set; }
        public bool   FoundNpu           { get; set; }
        public string ErrorMessage       { get; set; }
        public string DetectionMethod    { get; set; } // "Attribute" or "Property"
        public double UtilizationPercent { get; set; }
        public uint   LuidLow            { get; set; }
        public int    LuidHigh           { get; set; }

        // parameterless ctor to set your defaults
        public NpuQueryResult()
        {
            AdapterDescription       = "N/A";
            IsNpu                   = false;
            BudgetBytes             = 0;
            CurrentUsageBytes       = 0;
            AvailableForOSBytes     = 0;
            AvailableForProcessBytes= 0;
            FoundNpu                = false;
            ErrorMessage            = null;
            DetectionMethod         = "None";
            UtilizationPercent      = -1; // -1 = not-available
            LuidLow                 = 0;
            LuidHigh                = 0;
        }
    }

    /// <summary>Measures engine-busy fraction over a short window.</summary>
    private static double? SampleNpuUtilPercent(IDXCoreAdapter adp,
                                                int sampleMs   = 200,
                                                uint engineIdx = 0,   // most NPUs expose one engine
                                                uint physIdx   = 0)
    {
        if (!adp.IsQueryStateSupported(DXCoreAdapterState.AdapterEngineRunningTimeMicroseconds))
            return null;

        nuint inSize   = (nuint)Marshal.SizeOf<DXCoreAdapterEngineIndex>();
        nuint outSize  = (nuint)Marshal.SizeOf<DXCoreEngineQueryOutput>();
        IntPtr inPtr   = Marshal.AllocHGlobal((int)inSize);
        IntPtr outPtr  = Marshal.AllocHGlobal((int)outSize);

        try
        {
            var idx = new DXCoreAdapterEngineIndex { physicalAdapterIndex = physIdx, engineIndex = engineIdx };
            Marshal.StructureToPtr(idx, inPtr, false);

            // first sample
            int hr = adp.QueryState(DXCoreAdapterState.AdapterEngineRunningTimeMicroseconds,
                                    inSize, inPtr, outSize, outPtr);
            if (hr < 0) return null;
            var first = Marshal.PtrToStructure<DXCoreEngineQueryOutput>(outPtr).runningTime;

            System.Threading.Thread.Sleep(sampleMs);

            // second sample
            hr = adp.QueryState(DXCoreAdapterState.AdapterEngineRunningTimeMicroseconds,
                                inSize, inPtr, outSize, outPtr);
            if (hr < 0) return null;
            var second = Marshal.PtrToStructure<DXCoreEngineQueryOutput>(outPtr).runningTime;

            ulong deltaBusyUs = second - first;
            double elapsedUs  = sampleMs * 1_000.0;

            return Math.Min(100.0, (deltaBusyUs / elapsedUs) * 100.0); // cap at 100 %
        }
        finally
        {
            Marshal.FreeHGlobal(inPtr);
            Marshal.FreeHGlobal(outPtr);
        }
    }

    private static int CreateDxCoreList(IDXCoreAdapterFactory f,
                                    Guid[] attribs,        // may be length-0
                                    out IntPtr listPtr)
    {
        Guid iidList = DXCoreGUIDs.IID_IDXCoreAdapterList;

        // First try the obvious ("new") contract: pass the array you got.
        int hr = f.CreateAdapterList((uint)attribs.Length, attribs,
                                    ref iidList, out listPtr);
        if (hr != unchecked((int)0x80070057))        // E_INVALIDARG ?
            return hr;                               // ← success or other error

        // Retry with the "legacy" null-pointer contract.
        return f.CreateAdapterList(0, null, ref iidList, out listPtr);
    }

    public static NpuQueryResult GetNpuUsage()
    {
        NpuQueryResult result = new NpuQueryResult();
        IntPtr factoryPtr = IntPtr.Zero;
        IntPtr adapterListPtr = IntPtr.Zero;
        IntPtr npuAdapterPtr = IntPtr.Zero; // Pointer to the found NPU adapter
        IDXCoreAdapterFactory factory = null;
        IDXCoreAdapterList adapterList = null;
        IDXCoreAdapter npuAdapter = null; // Managed object for the found NPU adapter

        try
        {
            // 1. Create Factory
            Guid factoryIID = DXCoreGUIDs.IID_IDXCoreAdapterFactory;
            int hr = NativeMethods.DXCoreCreateAdapterFactory(ref factoryIID, out factoryPtr);
            if (hr < 0 || factoryPtr == IntPtr.Zero) { result.ErrorMessage = "Failed to create DXCore Adapter Factory (HRESULT: 0x" + hr.ToString("X8") + ")"; return result; }
            factory = Marshal.GetObjectForIUnknown(factoryPtr) as IDXCoreAdapterFactory;
            if (factory == null) { result.ErrorMessage = "Failed to get Factory object from pointer."; return result; } // Pointer released in finally

            // 2. Create Adapter List using a COMMON FILTER (e.g., Compute)
            //    This system seems to require *a* filter, so use one likely supported by GPUs/NPUs.
            //    The loop logic below will perform the actual NPU-specific check.
            // 1st choice - adapters that advertise ML meta-commands
            Guid[] mlFilter = { DXCoreGUIDs.DXCORE_ADAPTER_ATTRIBUTE_D3D12_GENERIC_ML };
            hr = CreateDxCoreList(factory, mlFilter, out adapterListPtr);

            // If there are none, drop back to the old CORE_COMPUTE list
            if (hr >= 0 && adapterListPtr != IntPtr.Zero &&
                ((IDXCoreAdapterList)Marshal.GetObjectForIUnknown(adapterListPtr))
                    .GetAdapterCount() == 0)
            {
                Marshal.Release(adapterListPtr);
                Guid[] computeFilter = { DXCoreGUIDs.DXCORE_ADAPTER_ATTRIBUTE_D3D12_CORE_COMPUTE };
                hr = CreateDxCoreList(factory, computeFilter, out adapterListPtr);
            }
            if (hr < 0 || adapterListPtr == IntPtr.Zero)
            {
                // Adjust error message slightly
                result.ErrorMessage = "Failed to create Compute-filtered DXCore Adapter List (HRESULT: 0x" + hr.ToString("X8") + "). Cannot proceed.";
                return result;
            }
            adapterList = Marshal.GetObjectForIUnknown(adapterListPtr) as IDXCoreAdapterList;
            if (adapterList == null) { result.ErrorMessage = "Failed to get Adapter List object from pointer."; return result; } // Pointer released in finally

            // 3. Iterate through adapters IN THE FILTERED LIST to find the NPU
            //    The list might contain the GPU, but the checks inside the loop are specific to NPU.
            uint adapterCount = adapterList.GetAdapterCount();
            if (adapterCount == 0) {
                 // If compute filter yields nothing, it's unlikely an NPU is present or detectable this way
                 result.ErrorMessage = "No adapters found matching the D3D12_CORE_COMPUTE filter.";
                 return result;
            }

            Guid adapterIID = DXCoreGUIDs.IID_IDXCoreAdapter;
            Guid npuAttributeGuid = DXCoreGUIDs.DXCORE_HARDWARE_TYPE_ATTRIBUTE_NPU; // Cache the GUID
            Guid gpuAttributeGuid = DXCoreGUIDs.DXCORE_HARDWARE_TYPE_ATTRIBUTE_GPU; // Cache the GUID
            Guid graphicsAttr = DXCoreGUIDs.DXCORE_ADAPTER_ATTRIBUTE_D3D12_GRAPHICS; // Cache the GUID

            for (uint i = 0; i < adapterCount; i++)
            {
                IntPtr currentAdapterPtr = IntPtr.Zero;
                IDXCoreAdapter currentAdapter = null;
                bool isNpuCandidate = false;
                string currentDetectionMethod = "None";

                try
                {
                    hr = adapterList.GetAdapter(i, ref adapterIID, out currentAdapterPtr);
                    if (hr < 0 || currentAdapterPtr == IntPtr.Zero)
                    {
                        Debug.WriteLine("Failed to get adapter at index " + i + " from compute list (HRESULT: 0x" + hr.ToString("X8") + ". Skipping.");
                        continue; // Try next adapter
                    }
                    currentAdapter = Marshal.GetObjectForIUnknown(currentAdapterPtr) as IDXCoreAdapter;
                    if (currentAdapter == null || !currentAdapter.IsValid())
                    {
                        Debug.WriteLine("Adapter at index " + i + " from compute list is null or invalid. Skipping.");
                        continue; // Try next adapter
                    }

                    // --- Debug Output ---
                    string desc = GetAdapterStringProperty(currentAdapter, DXCoreAdapterProperty.DriverDescription);
                    bool? isHardware = GetAdapterProperty<bool>(currentAdapter, DXCoreAdapterProperty.IsHardware);
                    Debug.WriteLine(
                        string.Format(
                            "Checking Adapter {0} (from compute list): Desc='{1}', IsHardware={2}",
                            i,
                            desc,
                            isHardware.HasValue
                                ? isHardware.Value.ToString()
                                : "N/A"
                        )
                    );
                    // --- End Debug ---

                    bool graphics = currentAdapter.IsAttributeSupported(ref graphicsAttr);
                    bool gpuAttr  = currentAdapter.IsAttributeSupported(ref gpuAttributeGuid);
                    bool npuAttr = currentAdapter.IsAttributeSupported(ref npuAttributeGuid);

                    bool? isNpuProp = null;
                    if (currentAdapter.IsPropertySupported(DXCoreAdapterProperty.IsNpu))
                        isNpuProp = GetAdapterProperty<bool>(currentAdapter, DXCoreAdapterProperty.IsNpu);

                    // keep the adapter ONLY if
                    //  - it advertises the NPU hardware-type attribute  (rule A)
                    //    - and does NOT advertise GPU                    (Arc fails here)
                    //  - OR the IsNpu property is explicitly true        (rule B)
                    if ((npuAttr && !gpuAttr) || (isNpuProp == true))
                    {
                        isNpuCandidate = true;
                        currentDetectionMethod = npuAttr ? "Attribute" : "Property";
                        Debug.WriteLine("Adapter " + i + " '" + desc + "' identified as NPU via Attribute or Property.");
                    }

                    // If identified SPECIFICALLY AS NPU by either method:
                    if (isNpuCandidate)
                    {
                        // *** Sanity check: IsHardware should ideally be true ***
                        if (isHardware == false) // Check the cached value
                        {
                            Debug.WriteLine("Adapter " + i + " '" + desc + "' identified as NPU (" + currentDetectionMethod + ") but IsHardware reported false. Skipping (Likely Software).");
                            continue; // Skip this one, might be software adapter
                        }

                        // Found the NPU! Assign and break the loop.
                        result.FoundNpu = true;
                        result.DetectionMethod = currentDetectionMethod;
                        npuAdapter = currentAdapter;       // Transfer ownership to outer scope
                        npuAdapterPtr = currentAdapterPtr; // Transfer ownership to outer scope
                        currentAdapter = null;       // Prevent release in inner finally
                        currentAdapterPtr = IntPtr.Zero; // Prevent release in inner finally
                        Debug.WriteLine("*** Selected NPU: " + desc + " (Index " + i + ") ***");
                        break; // Exit the loop - WE FOUND THE NPU
                    }
                    else {
                        Debug.WriteLine("Adapter " + i + " '" + desc + "' did NOT match NPU criteria. Checking next.");
                    }
                }
                finally
                {
                    // Release the *current* adapter COM object *if* it wasn't selected
                    // Let the RCW handle the underlying pointer release.
                    if (currentAdapter != null) Marshal.ReleaseComObject(currentAdapter);
                    // Optional: Zero out local pointer for clarity, though it goes out of scope anyway
                    // currentAdapterPtr = IntPtr.Zero;
                }
            } // End of adapter loop

            // 4. Query NPU State (only if an NPU was found)
            if (result.FoundNpu && npuAdapter != null)
            {
                result.AdapterDescription = GetAdapterStringProperty(npuAdapter, DXCoreAdapterProperty.DriverDescription);
                var luid = GetAdapterProperty<LUID>(npuAdapter, DXCoreAdapterProperty.InstanceLuid);
                if (luid.HasValue)
                {
                    result.LuidLow  = luid.Value.LowPart;
                    result.LuidHigh = luid.Value.HighPart;
                }
                result.IsNpu = true; // Set explicitly based on detection

                // *** Query Memory Budget (Corrected logic based on expert feedback) ***
                if (npuAdapter.IsQueryStateSupported(DXCoreAdapterState.AdapterMemoryBudget))
                {
                    var segGroups = new[] { DXCoreSegmentGroup.Local, DXCoreSegmentGroup.NonLocal };
                    uint node = 0;                                      // Single-adapter case

                    result.BudgetBytes = 0;         // reset accumulators
                    result.CurrentUsageBytes = 0;
                    result.AvailableForOSBytes = 0;
                    result.AvailableForProcessBytes = 0;

                    foreach (var grp in segGroups)
                    {
                        var seg = new DXCoreAdapterMemoryBudgetNodeSegmentGroup
                                { nodeIndex = node, segmentGroup = grp };

                        nuint inSize  = (nuint)Marshal.SizeOf(seg);
                        nuint outSize = (nuint)Marshal.SizeOf<DXCoreAdapterMemoryBudget>();

                        IntPtr inPtr  = Marshal.AllocHGlobal((int)inSize);
                        IntPtr outPtr = Marshal.AllocHGlobal((int)outSize);
                        try
                        {
                            Marshal.StructureToPtr(seg, inPtr, false);
                            hr = npuAdapter.QueryState(DXCoreAdapterState.AdapterMemoryBudget,
                                                        inSize,  inPtr,
                                                        outSize, outPtr);
                            if (hr == 0)                            // S_OK
                            {
                                var b = Marshal.PtrToStructure<DXCoreAdapterMemoryBudget>(outPtr);

                                // Option 1 - aggregate the numbers:
                                result.BudgetBytes          += b.budget;
                                result.CurrentUsageBytes    += b.currentUsage;
                                result.AvailableForOSBytes  += b.availableForOS;
                                result.AvailableForProcessBytes += b.availableForProcess;

                                // Option 2 - keep per-segment results in a list if you need detail.
                            }
                            else if (hr == unchecked((int)0x887A0002) /* DXGI_ERROR_NOT_FOUND */ ||
                                    hr == unchecked((int)0x887A0001) /* DXGI_ERROR_INVALID_CALL */)
                            {
                                // This node/segment simply isn't present - ignore.
                            }
                            else
                            {
                                Debug.WriteLine("Budget query failed for " + grp + ": 0x" + hr.ToString("X8"));
                            }
                        }
                        finally
                        {
                            Marshal.FreeHGlobal(inPtr);
                            Marshal.FreeHGlobal(outPtr);
                        }
                    }
                }
                else
                {
                     // State is not supported
                     result.ErrorMessage = (result.ErrorMessage ?? "") + " NPU adapter does not support querying memory budget state (DXCoreAdapterState.AdapterMemoryBudget).";
                     Debug.WriteLine("NPU adapter reported IsQueryStateSupported(AdapterMemoryBudget) as false.");
                }

                // --- Live utilisation ----------------------------------------------------
                if (npuAdapter.IsQueryStateSupported(
                        DXCoreAdapterState.AdapterEngineRunningTimeMicroseconds))
                {
                    double? util = SampleNpuUtilPercent(npuAdapter);   // 200 ms window
                    result.UtilizationPercent = util ?? -1;            // -1 = no data
                }
                else
                {
                    result.UtilizationPercent = -1;
                    result.ErrorMessage +=
                        " AdapterEngineRunningTimeMicroseconds not supported by this " +
                        "driver/OS - cannot sample live utilisation.";
                }
            }
            else if (!result.FoundNpu && string.IsNullOrEmpty(result.ErrorMessage))
            {
                // If loop finished and no NPU found, set a message
                result.ErrorMessage = "No NPU detected using Attribute or IsNpu Property methods.";
            }
        }
        catch (Exception ex)
        {
            result.ErrorMessage = "An exception occurred: " + ex.GetType().Name + " - " + ex.Message;
            // Include stack trace in debug output for more detail
            Debug.WriteLine("Exception in GetNpuUsage: " + ex.ToString());
        }
        finally
        {
            // Release COM objects obtained. Let the RCWs handle the release.
            // Order can matter if objects depend on each other, but here releasing adapter->list->factory is safe.

            // Release the NPU adapter's RCW (if found and assigned)
            if (npuAdapter != null) Marshal.ReleaseComObject(npuAdapter);

            // Release the list's RCW
            if (adapterList != null) Marshal.ReleaseComObject(adapterList);

            // Release the factory's RCW
            if (factory != null) Marshal.ReleaseComObject(factory);

            // It's good practice (though not strictly required for the fix) to ensure
            // the IntPtrs are zeroed after the corresponding RCW is released,
            // especially if they were transferred like npuAdapterPtr.
            npuAdapterPtr = IntPtr.Zero;
            adapterListPtr = IntPtr.Zero;
            factoryPtr = IntPtr.Zero;
        }
        return result;
    }

    // --------------------------------------------------------------------
    //  LUID → human-readable description via DXCore
    // --------------------------------------------------------------------
    /// Translate a DXCore adapter-LUID into the adapter's DriverDescription.
    /// Works on any Windows 10 19041 + (no extra DLLs).
    public static string DescribeAdapterByLuid(uint low, int high)
    {
        // 1. factory ---------------------------------------------------------
        Guid iidFactory = DXCoreGUIDs.IID_IDXCoreAdapterFactory;
        if (NativeMethods.DXCoreCreateAdapterFactory(ref iidFactory,
                                                    out IntPtr facPtr) < 0 ||
            facPtr == IntPtr.Zero)
            return "DXCore factory failed";

        var factory = (IDXCoreAdapterFactory)Marshal.GetObjectForIUnknown(facPtr);
        try
        {
            //-----------------------------------------------------------------
            // 2. adapter list - *must* contain at least one attribute on
            //    builds that reject an unfiltered list with E_INVALIDARG.
            //-----------------------------------------------------------------
            Guid[] compute = { DXCoreGUIDs.DXCORE_ADAPTER_ATTRIBUTE_D3D12_CORE_COMPUTE };

            IntPtr listPtr;
            int hr = CreateDxCoreList(factory, compute, out listPtr);
            if (hr < 0 || listPtr == IntPtr.Zero)
                return ("CreateAdapterList failed: 0x" + hr.ToString("X8"));

            var list = (IDXCoreAdapterList)Marshal.GetObjectForIUnknown(listPtr);
            try
            {
                uint count  = list.GetAdapterCount();
                Guid iidAdp = DXCoreGUIDs.IID_IDXCoreAdapter;

                for (uint i = 0; i < count; ++i)
                {
                    list.GetAdapter(i, ref iidAdp, out IntPtr adpPtr);
                    if (adpPtr == IntPtr.Zero) continue;

                    var adp = (IDXCoreAdapter)Marshal.GetObjectForIUnknown(adpPtr);
                    try
                    {
                        var luid = GetAdapterProperty<LUID>(
                                    adp, DXCoreAdapterProperty.InstanceLuid);

                        if (luid.HasValue &&
                            luid.Value.LowPart == low &&
                            luid.Value.HighPart == high)
                            return GetAdapterStringProperty(
                                    adp, DXCoreAdapterProperty.DriverDescription);
                    }
                    finally { Marshal.ReleaseComObject(adp); }
                }
            }
            finally { Marshal.ReleaseComObject(list); }
        }
        finally { Marshal.ReleaseComObject(factory); }

        return "adapter not found";
    }
}

//---------------------------------------------------------------------
//  Minimal DXGI helper - never changes across Windows builds
//---------------------------------------------------------------------
[ComImport, Guid("25F8848E-6DA9-4C04-8F6D-72B6CF5ECA5B"),
 InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IDXGIFactory4
{
    void _VtblGap1_15();                              // skip 15 methods
    int EnumAdapterByLuid(LUID luid, in Guid iid, out IntPtr ppv);
}

[ComImport, Guid("1BC6EA02-EF36-464F-BF0C-21CA39E5168A"),
 InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IDXGIAdapter4
{
    void _VtblGap1_3();                               // skip to GetDesc3
    DXGI_ADAPTER_DESC3 GetDesc3();
}

// -----------------------------------------------------------------
//  Minimal IDXGIFactory2 - only what we actually call
// -----------------------------------------------------------------
[ComImport,
 Guid("50C83A1C-E072-4C48-87B0-3630FA36A6D0"),
 InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IDXGIFactory2
{
    /* 0-15  -- same as IDXGIFactory1 - we don't use them here */
    void _VtblGap1_15();

    /* 16 */              // **THIS is the one we need**
    [PreserveSig] int EnumAdapters1(
        uint   adapter,
        out IntPtr ppvAdapter);        // returns IDXGIAdapter1

    /* 17-?? further methods we don't need */
    void _VtblGap2_1();
}


[ComImport, Guid("770aae78-f26f-4dba-a829-253c83d1b387"),
 InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IDXGIAdapter1
{
    void _VtblGap1_8();                 // skip to GetDesc1
    DXGI_ADAPTER_DESC1 GetDesc1();
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
struct DXGI_ADAPTER_DESC3
{
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
    public string Description;
    public uint VendorId, DeviceId, SubSysId, Revision;
    public IntPtr DedicatedVideoMemory, DedicatedSystemMemory, SharedSystemMemory;
    public LUID AdapterLuid;
    public uint Flags;                                // DXGI_ADAPTER_FLAG3_
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
struct DXGI_ADAPTER_DESC1
{
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
    public string Description;
    public uint VendorId, DeviceId, SubSysId, Revision;
    public IntPtr DedicatedVideoMemory, SharedSystemMemory;
    public IntPtr AdapterLuid;          // actually a LUID struct
    public uint Flags;
}

public static class DxgiMini
{
    [DllImport("dxgi.dll")]
    static extern int CreateDXGIFactory2(uint flags, in Guid iid, out IntPtr ppv);

    public static string GetDesc(uint low, int high)
    {
        Guid iidFac4 = typeof(IDXGIFactory4).GUID;
        int  hr      = CreateDXGIFactory2(0, in iidFac4, out var facPtr);

        bool haveFac4 = hr >= 0;          // first attempt succeeded?

        if (!haveFac4)
        {
            // ── fallback to the older factory ─────────────────────────────
            Guid iidFac2 = typeof(IDXGIFactory2).GUID;
            hr = CreateDXGIFactory2(0, in iidFac2, out facPtr);

            if (hr < 0 || facPtr == IntPtr.Zero)
                return ("CreateDXGIFactory2 failed: 0x" + hr.ToString("X8"));

            var fac2 = (IDXGIFactory2)Marshal.GetObjectForIUnknown(facPtr);
            for (uint i = 0; fac2.EnumAdapters1(i, out var adpPtr) == 0; ++i)
            {
                var adp1 = (IDXGIAdapter1)Marshal.GetObjectForIUnknown(adpPtr);
                var d    = adp1.GetDesc1();
                var luid = Marshal.PtrToStructure<LUID>(d.AdapterLuid);
                if (luid.LowPart == low && luid.HighPart == high)
                    return (d.Description.Trim() + "  (Flags=0x" + d.Flags.ToString("X") + ")");

                Marshal.ReleaseComObject(adp1);
            }
            return "adapter not found";
        }

        // we really have an IDXGIFactory4 here
        var fac = (IDXGIFactory4)Marshal.GetObjectForIUnknown(facPtr);
        try
        {
            Guid iidAdp4 = typeof(IDXGIAdapter4).GUID;
            var  luid    = new LUID { LowPart = low, HighPart = high };

            if (fac.EnumAdapterByLuid(luid, in iidAdp4, out var adpPtr) < 0)
                return "adapter not found";

            var adp = (IDXGIAdapter4)Marshal.GetObjectForIUnknown(adpPtr);
            try
            {
                var d = adp.GetDesc3();
                return (d.Description.Trim() + "  (Flags=0x" + d.Flags.ToString("X") + ")");
            }
            finally { Marshal.ReleaseComObject(adp); }
        }
        finally { Marshal.ReleaseComObject(fac); }
    }
}
"@
#endregion C# Code Definition #########################################################

Write-Verbose "Attempting to compile DXCore C# helper..."
try {
    # Add Debug flag for richer C# out`t if needed during testing
    # Add-Type -TypeDefinition $CSharpSource -ReferencedAssemblies System.Diagnostics.Process -ErrorAction Stop
    Add-Type -TypeDefinition $CSharpSource -ErrorAction Stop
    Write-Verbose "C# helper compiled successfully."
} catch {
    Write-Error "Failed to compile C# code. Ensure .NET Framework/Core compatible with PowerShell is installed and DXCore dependencies are met."
    Write-Error "Compilation Error: $($_.Exception.Message)"
    # Display loader exceptions if available (useful for P/Invoke/dependency issues)
    if ($_.Exception -is [System.Reflection.ReflectionTypeLoadException]) {
        Write-Error "Loader Exceptions:"
        $_.Exception.LoaderExceptions | ForEach-Object { Write-Error "- $($_.Message)" }
    }
    return 1
}
#endregion Create a DXCore C# wrapper so that we can access NPU usage from PowerShell #

Write-Verbose "Querying NPU information via DXCore (using Attribute/Property detection)..."
try {
    $npuResult = [NpuQuery]::GetNpuUsage()

    if ($null -eq $npuResult) {
        Write-Error "The NPU query method returned null unexpectedly."
    } elseif ($npuResult.FoundNpu) {
        Write-Host "NPU Adapter Found:" -ForegroundColor Cyan
        Write-Host "  Detection Method: $($npuResult.DetectionMethod)"
        Write-Host "  Description: $($npuResult.AdapterDescription)"

        # Display Memory Budget info if available
        $budgetMB = if ($npuResult.BudgetBytes -gt 0) { [Math]::Round($npuResult.BudgetBytes / 1MB, 2) } else { 0 }
        $usageMB = if ($npuResult.CurrentUsageBytes -gt 0) { [Math]::Round($npuResult.CurrentUsageBytes / 1MB, 2) } else { 0 }

        # Check if budget info was successfully queried (even if values are 0)
        $budgetQueryAttempted = $true # Assume attempted unless specific error below
        if (-not [string]::IsNullOrEmpty($npuResult.ErrorMessage) -and $npuResult.ErrorMessage.Contains("memory budget state")) {
            Write-Warning "  Could not query memory budget state."
            #Write-Warning "  Details: $($npuResult.ErrorMessage)" # Message can be long, maybe hide details by default
            $budgetQueryAttempted = $false
        }

        if ($budgetQueryAttempted) {
            if ($npuResult.BudgetBytes -gt 0 -or $npuResult.CurrentUsageBytes -gt 0) {
                Write-Host "  Memory Budget: $budgetMB MB ($($npuResult.BudgetBytes) Bytes)"
                # Write-Host "  Current Usage: $usageMB MB ($($npuResult.CurrentUsageBytes) Bytes)" # This is not working on Snapdragon X Elite, but the query is done correctly because it works on GPUs... must be a limitaiton of DxCore on the NPU side.
                # Optional: Display Available OS/Process if needed
                # $availOS_MB = if ($npuResult.AvailableForOSBytes -gt 0) { [Math]::Round($npuResult.AvailableForOSBytes / 1MB, 2) } else { 0 }
                # $availProc_MB = if ($npuResult.AvailableForProcessBytes -gt 0) { [Math]::Round($npuResult.AvailableForProcessBytes / 1MB, 2) } else { 0 }
                # Write-Host "  Available for OS: $availOS_MB MB"
                # Write-Host "  Available for Process: $availProc_MB MB"
            } else {
                Write-Host "  Memory budget information reported as zero or query returned zero values."
            }
        }

        if ($npuResult.UtilizationPercent -ge 0) {
            Write-Host ("  NPU Utilization : {0:N1}%" -f $npuResult.UtilizationPercent)
        } else {
            Write-Verbose "  Could not read live utilisation (driver/OS doesn't expose AdapterEngineRunningTime)."

            $doubleNPUUtil = Get-NPUUtilization
            Write-Host ("  NPU Utilization : {0:N1}%" -f $doubleNPUUtil)
        }

        # Display other non-budget error messages if they exist
        if (-not [string]::IsNullOrEmpty($npuResult.ErrorMessage) -and -not $npuResult.ErrorMessage.Contains("memory budget state") ) {
            Write-Warning "Query completed with additional message(s): $($npuResult.ErrorMessage)"
        }
    } else {
        $doubleNPUUtil = Get-NPUUtilization
        if ($doubleNPUUtil -ne [double]::NaN) {
            Write-Warning "This appears to be an Intel system with an NPU; unfortunately Intel's driver does not expose the NPU to the operating system correctly, so this script is limited to reporting % utilization."
            if (-not [string]::IsNullOrEmpty($npuResult.ErrorMessage)) {
                Write-Warning "Details: $($npuResult.ErrorMessage)"
            }
            Write-Host ("  NPU Utilization : {0:N1}%" -f $doubleNPUUtil)
        } else {
            Write-Warning "No NPU detected or query failed."
            if (-not [string]::IsNullOrEmpty($npuResult.ErrorMessage)) {
                Write-Warning "Details: $($npuResult.ErrorMessage)"
            }
        }
    }
} catch {
    Write-Error "An error occurred while executing the NPU query:"
    Write-Error $_.Exception.ToString() # Full exception details for debugging
}
