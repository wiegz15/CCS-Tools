ZTOOLS launcher should be working now also. 


**Download this whole repository to use this.  I am working on creating a launcher exe or file to make it easier to run.**


This version uses check boxes and has a select all if you need.  We need to move the last few scripts because they need some work and actually make changes on the hosts so try not to run the select all until I get those moved. 

This will also create the tools.ini file in the bin folder.  It checks for this and if it does not exist it will ask you to create.  You will enter the VCenter server address.   I added this file to the gitignore so we do not upload the vcenter address to the main repository.  It will not run if this file does not exist but will take you through creating.  

Let me know if you run into anything. 


**Plan to Add a check to install ImportExcel or PowerCLI if needed.  Should help when setting up in new env**
