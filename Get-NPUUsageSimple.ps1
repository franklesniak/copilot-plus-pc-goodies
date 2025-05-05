# NOTE: This script does not work on Intel NPUs as of 2025-05-05.
# Its usage is predicated on the NPU device having a name that includes "compute";
# therefore, this may not be the most reliable script over time.

$doubleNPUUtil = [double]0
Get-Counter '\GPU Engine(*compute*)\Utilization Percentage' |
    ForEach-Object { $_.CounterSamples } |
    ForEach-Object { $doubleNPUUtil += $_.CookedValue }
return $doubleNPUUtil
