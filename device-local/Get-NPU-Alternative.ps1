Get-CimInstance Win32_PnPEntity | Where-Object { $_.ClassGuid -eq '{F01A9D53-3FF6-48D2-9F97-C8A7004BE10C}' } | Select-Object Name, Status, ClassGuid, PnPDeviceID
