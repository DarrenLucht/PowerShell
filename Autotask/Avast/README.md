# (Autotask) datto RMM Avast Device Monitor

Currently datto RMM is able to detect Avast Cloudcare for client workstations and AEM Reports will indicate which devices have it installed, and if they are up to date or not.

Unfortunately server devices are not showing Antivirus Product and Antivirus Status. Avast Cloudcare is not one of the 9 AV solutions that datto RMM monitors.

The Avast Antivirus Monitor.cpt was created from the AvastRMM powershell script as a device monitor component within datto RMM and downloaded. To deploy within RMM, goto Components > Import Component > Choose file > Upload.

If you wish to create your own custom device monitor component, use the code within the AvastRMM.ps1 powershell script to roll your own.

** You must create a client site monitoring policy within datto RMM and it must be enabled. **
