$arrNPUs = @(Get-PnpDevice -Class "ComputeAccelerator" -Status OK -ErrorAction SilentlyContinue | Select-Object Status, Class, FriendlyName, InstanceID)
if ($arrNPUs.Count -eq 0) {
    # No NPU found - don't trigger the proactive remediation
    return 0
} else {
    # look for some other condition, like a setting that you want to check
    # or a specific NPU driver version that you want to check for
    $condition = $true # this is just a placeholder for your condition check
    if ($condition) {
        # Device has an NPU and the condition was met, trigger the proactive
        # remediation
        return 1
    } else {
        # Device has an NPU but the Condition was not met, so do not trigger the
        # proactive remediation
        return 0
    }
}
