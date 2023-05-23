# ATC-Tool-manager
Excel files to keep track of tools used in your cnc mills ATC.


 The python script will need extensions downloaded to be able to use ( watchdog, Pywin32).
  Once you setup python and the extensions. You can open your favorite code editor and edit the path to your NC file and to you excel files of you machgines toollists.
   You would also want to edit the ReadNCFile macro in excel as well.
   
    If all is edited correctly the script will keep eatch over the NC file for any changes and open the correct excel file to the machgine your posting to. It will then post the new tool numbers and tool desription in the correct columns. When a new NC file geets posted the file will open back up and replace the tool if it was changed to the current tool being used.
    
     You will also have edit your NC post to post machine type (vendor: Haas VF-3..etc) and tool number and description ( T2 H2 etc etc etc etc )
