# Thanks to:
# http://stackoverflow.com/a/31624001

import sys
import os
import time
import glob
import win32com.client

FILENAME = "Understanding Git"

def main():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    success = PPTtoPDF(dir_path + "\\" + FILENAME + ".pptx", dir_path + "\\" + FILENAME + ".pdf")

    if success:
        print("Success!")
    else:
        print("Could not copy file.")
    

def PPTtoPDF(inputFileName, outputFileName, formatType = 32):
    try:
        if outputFileName[-3:] != 'pdf':
            outputFileName = outputFileName + ".pdf"

        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        powerpoint.Visible = 1
        
        deck = powerpoint.Presentations.Open(inputFileName)
        time.sleep(0.5)
        deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
        deck.Close()
        
        powerpoint.Quit()

        return True
    
    except:
        print("Error converting file: " + filename)
        return False



if __name__ == "__main__":
    main()
