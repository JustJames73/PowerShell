# PowerShell
Over the years I've written a few good utility scripts. Some have been improved with the help of AI. Either way these are the scripts I am proud of, and use frequently. 

I use *ConverTo-PowerShellArray.ps1* (alias CTA) almost daily to transform a list from the clipboard or a list into an array. 
The *Test-MultipleConnections.ps1* (alias TMC) has been helpful in testing VPN and Firewall rules by outputing a timestamped list for comparrision against logs, or just documenting that they were sucessful. It runs additional tests on web servers to verify that it's more than the 80/443 parts are open. 
The *rbac_analysis_script* was a pet project many years in the making; it took me a long time to figure out the math for a useful evaluation. Claude.AI took the base code and turned it into a useful tool for recomending which security groups could be used for AGUDLP style RBAC groups. 
*SearchForCommandFromModule* is helpful when updating scripts that are using depreciated commands or modules by finding the lines with commands from a specific module. If you remember the Quest AD Modules (back in the PoSh 2.0 days), I needed to find all the references to update our scripts once those features became native to powershell. 
*ou permissions.ps1* is an other that is useful when you need it, because figureing out the permissions applied to an OU can be a maddening task!
