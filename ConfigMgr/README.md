# Configuration Manager - NPU WQL Query

This query looks for devices that have a PNP device driver that equals the NPU class GUID. The NPU class GUID is `{f01a9d53-3ff6-48d2-9f97-c8a7004be10c}`.

```wql
select SMS_R_System.Name, SMS_G_System_PNP_DEVICE_DRIVER.ClassGuid, SMS_G_System_PNP_DEVICE_DRIVER.Name, SMS_R_System.ClientType, SMS_R_System.LastLogonUserName, SMS_R_System.LastLogonUserDomain, SMS_R_System.OperatingSystemNameandVersion from  SMS_R_System inner join SMS_G_System_PNP_DEVICE_DRIVER on SMS_G_System_PNP_DEVICE_DRIVER.ResourceID = SMS_R_System.ResourceId
where SMS_G_System_PNP_DEVICE_DRIVER.ClassGuid = "{f01a9d53-3ff6-48d2-9f97-c8a7004be10c}"
```

Here is what it looks like in the Configuration Manager console:
![NPU WQL Query](/.Images/ConfigMgr_WQL_QueryLanguage.png)

Running the report will generate a list of devices that have the NPU class GUID.
![NPU Query Results](/.Images/ConfigMgr_WQL_Results.png)