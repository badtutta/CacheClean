<#


               Name:    CacheClean.ps1                                                               
                                                                                                   
          Author(s):    Navy Flank Speed MCS                                                                                                                                                 
                                                                                                       
            Version:    1.0.4                                                                        
                                                                                                  
            Created:    01.18.2023                                                                 
           Modified:    12.21.2023                                                                  
                                                                                                  
                                                                                                  
                                                                                                  
          Info:                                                                                   
                                                                                                  
          This script can determine the size of the user's profile and be used to perform a disk cleanup 
          to free up space via CleanMgr.exe, as well as, backup of Microsoft Edge favorites to an html file.
	  It will all perform the cleaning of the following caches...
		  
		        -  PowerBI
			-  Teams
			-  Outlook
			-  OneDrive

	  It will also back up favorites from the Edge browser
          

          
          Parameters:
    
          	  -ProfileSize = Determines disk space used by the user's profile
		  -Clean = Cleans with default options
		  -Options = Opens GUI and allows selection of more options to clean
		  -MSTeams = Cleans Microsoft Teams cache
		  -OneDrive = Cleans Microsoft OneDrive cache
		  -PowerBI = Cleans Microsoft PowerBI cache
		  -Outlook = Cleans Microsoft Outlook cache
		  -Edge = Backs up Edge Favorites to location of user's choosing
		  
		  
          

          Version Log:

          	  .1 - Initial script
		  .2 - Added relaunch of apps after cache cleaning
		  .3 - Added Edge Favorites backup ability
		  .4 - Added support for new Teams version
                        

#>
