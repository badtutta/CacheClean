# CacheClean

               Name:    CacheClean.ps1                                                               
                                                                                                   
          Author(s):    Navy Flank Speed MCS                                                                                                                                                 
                                                                                                       
            Version:    1.0.x                                                                        
                                                                                                  
            Created:    01.18.2023                                                                 
           Modified:    03.07.2023                                                                  
                                                                                                  
                                                                                                  
                                                                                                  
          Info:                                                                                   
                                                                                                  
          This script can determine the size of the user's profile and be used to perform a disk cleanup 
          to free up space via CleanMgr.exe, as well as, cleaning of the following cache...
		  
		  -  PowerBI
			-  Teams
			-  Outlook
			-  OneDrive
          

          
          Parameters:
    
          -ProfileSize = Determines disk space used by the user's profile
		  -Clean = Cleans with default options
		  -Options = (used with -Clean) Opens GUI and allows selection of more options to clean
		  -MSTeams = Cleans Microsoft Teams cache
		  -OneDrive = Cleans Microsoft OneDrive cache
		  -PowerBI = Cleans Microsoft PowerBI cache
		  -Outlook = Cleans Microsoft Outlook cache
