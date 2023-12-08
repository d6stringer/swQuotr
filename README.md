# swQuotr

Created by Daniel Woodson  
 Started On: 12/5/2023  
 From Sources:  
 https://stackoverflow.com/questions/53682824/convert-all-solidworks-files-in-folder-to-step-files-macro  
 https://help.solidworks.com/2022/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDocExtension~SaveAs3.html  
 V0.1: 12/6/2023  
 Init functional copy: creates PDFs from solidworks drawing files and step files from solidworks part and assembly files. Very little testing done,  
 probably still very buggy.  
    - Assms not tested  
    - Long file paths may be a problem  

# Usage
    1. Move all SW files you wish to convert to a single directory/folder
    2. Open SW, close any open files
    3. Start the console application
    4. Enter the directory location
    5. Press enter, wait for system to finish conversion
    6. Check your files

# Build
Dependancies can normally be found somewhere like this:  
C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist\

SolidWorks.Interop.swconst  
SolidWorks.Interop.sldworks  

