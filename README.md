# Environment Scripts

You can automatically keep the ControlUp organizational tree up-to-date with the ongoing changes in your environment topology. Our synchronization scripts are run automatically as a Windows scheduled task to read your topology and update ControlUp with added or removed machines. Those changes are automatically reflected in the ControlUp organizational tree and don't have to be made manually. 

You can continuously monitor the actual machines in your environment and remediate any issues, saving you time and resources. 

- Our sync scripts are written in PowerShell and stored in this GitHub repository.
- Depending on your VDI, you may have to run special credentials scripts to enable running the sync scripts. These are detailed in the [knowledge base articles](https://support.controlup.com/hc/en-us) covering each environment.
You set the Windows scheduled task to automatically run the sync script on the ControlUp monitor machine. This procedure is detailed [here](https://support.controlup.com/hc/en-us/articles/360015854278).
