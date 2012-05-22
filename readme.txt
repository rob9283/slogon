slogon = Simple Logon

Two simple vbs logon scripts for Windows clients in Windows domain environments.  One maps printers, one maps shared drives.  Neat, right?

I first wrote these scripts years ago to overcome the severe limitations in Windows 2000 and Windows 2003.  The Group Policy Preferences that came with Server 2008 are good when they work, but a Google search for "group policy preferences mapping not working" brings up too many hits.  Maybe Windows Server 2012 will obviate scripts for good, but until then...  slogon.  

Rob Pennoyer
rob ]at[ robpennoyer ]dotcom[


_______________________________________________________________________

slogon shares.vbs
Maps drive shares, and can also map specific shares by user security group membership in active directory.  Can also map home folders, assuming the folders are named after the username and all located in the same share.

slogon printers.vbs
Maps printers by user security group membership, computer security group membership, or by active directory site.  It can force a specific default printer or retain the user's preference, again either for all users or by user group, computer group, or by active directory site.

Both scripts:
 - log their activities in the Application log, including the script file name such that if you maintain several versions you'll never be unclear about which one is running
 - trap errors, so typos don't throw error messages up on users' screens
 - are designed so you have have one version running for all users.  Stop maintaining separate scripts for different departments and users groups.

_______________________________________________________________________



vbscripts are preferable to simple shell scripts for a variety of reasons, not least of which is that they don't pop up a command prompt while they're running.  Most vbscript logon scripts available on the net require a lot of coding experience to adapt them to your environment.  Slogon scripts are designed to wrap the more complicated vbscripting up in simple, plain-English commands.  

In vbscript...
objNetwork.MapNetworkDrive "f:","\\server\share" 
...isn't impossible to understand, but did you instantiate the object correctly?  What if the drive already exists?  What if there's a typo in the share name?  What if you want to check group membership first?  

In slogon...
MapFolder "f:","\\server\share" 
..is as simple, but also traps errors and records its activity in the Application log so you can see what's happening.  


