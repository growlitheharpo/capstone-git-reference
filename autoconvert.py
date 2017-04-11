# Some help from an outdated post at:
# http://stackoverflow.com/a/31624001

import os
import time
import win32com.client

FILENAME = "Understanding Git"

def main():
    dir_path = os.path.dirname(os.path.realpath(__file__)) + "\\"
    success = convertPPTtoPDF(dir_path + FILENAME + ".pptx", dir_path + FILENAME + ".pdf")

    return success

def convertPPTtoPDF(inputFileName, outputFileName):
    try:
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        
        # call open, then sleep to give powerpoint the time to process the command
        deck = powerpoint.Presentations.Open(inputFileName)
        time.sleep(0.5)

        # save the file as the pdf. formatType is 32 for pdf (according to MS docs)
        deck.SaveAs(outputFileName, 32)
        deck.Close()
        
        powerpoint.Quit()

        return True
    
    except:
        print("Error converting file: " + inputFileName)
        return False



if __name__ == "__main__":
    main()
