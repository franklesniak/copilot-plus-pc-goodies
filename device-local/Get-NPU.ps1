Get-PnpDevice -Class "ComputeAccelerator" -Status OK -ErrorAction SilentlyContinue | Select-Object Status, Class, FriendlyName, InstanceID
