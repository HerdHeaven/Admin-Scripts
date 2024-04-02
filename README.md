Below is a shortoverview of each script contained in this repository:

Send-MagicPacket - This script builds a magic packet for a specific MAC Address to be used when Wake-on-LAN is enabled on the targeted device. 

Verify-HashValue - This was a quick script I've used for years to verify provided hash values when downloading things from the internet. Point 
                   it at a file and then paste the provided hash value for verification. 


Export-EntraGroupReport - #  THIS IS SET UP TO USE CERTIFICATE BASED AUTH  # Simple enough to change it around but if you need assistance please message me :)
                          Exports all groups using the MgGraph and ImportExcel modules. Creates a separate worksheet for each group containing a table of the 
                          members. Creates a worksheet containing all of the basic group overview and each groups display name links to the 
                          corresponding worksheet. Each worksheet has a Title in cell "A1" that links back to the overview worksheet. 
