# Repair-Office-365-After-Patch-Tuesday
Repair Office 365 After Patch Tuesday


This will make a schdudled task to set the channel to current then force an update and quick repair office after updating. 

Why repair: because office updating can cause issues with apps that use excel or the update didn't apply well. It's just to make sure everything runs smooth. 

This will run every tue at 6am using the default settings 

Run as admin

```powershell
.\Office365PatchTues-UpdateRepair.ps1
```
