# -*- coding: utf-8 -*-
"""
PyDF to Excel    : A program that reads PDF files as inputs, and then outputs them
                   in Excel format.

IO_Wrapper.py    : This module wraps around camelot; handles all input/output
                   
Created on       : Sat Sep 12 12:00:01 2020

@author: kevinhhl
Source code is publicly available on https://github.com/kevinhhl
"""
import pandas as pd
import pandas.io.formats.excel
from datetime import datetime
from pathlib import Path
import os
import camelot_modified        # Originally: camelot
                               # Modified it to allow a GUI instance to be passed in; for purpose of calling UI.progressBar.update() from camelot.PDF_Handler


TIMESTAMP_FORMAT = "%Y.%m.%d_%H%M%S"

def define_output_dir():
    #TODO make this dynamic, allow user to change:
    dest_dir = str(Path.home()) + '\\Desktop\\PyDF to Excel - Outputs\\'
    if not os.path.isdir(dest_dir):
        os.makedirs(dest_dir)
    return dest_dir


def _parsePageInstructions(pageInstructions):
    #pageInstructions = self.lineEdit_3.text().strip('\n').strip("\"")
    rngList = []
    for s in pageInstructions.replace(" ","").split(","):
        rngList.append(s.strip())
    
    return rngList  


def _countTotalPages(pageInstructions):
    #pageInstructions = self.lineEdit_3.text().strip('\n').strip().strip("\"").strip()
    countAllPages = 0
    for pgRange in pageInstructions.split(","):
        strBoundsInput = pgRange.split("-")
        
        if len(strBoundsInput) != 2 and len(strBoundsInput) != 1:
            print("invalid range provided: " + pgRange.strip())

        countRng = 0
        if len(strBoundsInput) == 1:
            countRng += 1
        else:
            for pg in range(int(strBoundsInput[0]),int(strBoundsInput[1])+1):
                countRng += 1

        countAllPages += countRng

    return countAllPages


def process_PDF(Ui_dialog_instance, pgInstructions, src, dest_dir):
    timeZero=datetime.now() # run time timer
    #print("process_PDF: Target pdf={0}; Dest={1}".format(src, dest))    
    rngList = _parsePageInstructions(pgInstructions)
    #print("process_PDF: Page #s to process: "+str(pgList))

    # Show UIs related to progress bar
    filenameTokens = src.split("\\")
    fName = filenameTokens[len(filenameTokens)-1]
    Ui_dialog_instance.label_9.setText("File in progress: {0}".format(fName))
    Ui_dialog_instance.label_9.show()
    Ui_dialog_instance.progressBar.show()
    
    # Define output filename & absolute path
    timestamp = datetime.now()
    outputfileName = "PyDF_Output_{0}.xlsx".format(timestamp.strftime("%Y.%m.%d_%H%M%S"))

    # Iterate through each page in pgList
    writer = pd.ExcelWriter(dest_dir + outputfileName, engine='xlsxwriter')
    skipped_pages = [] #int pg. number
    
    #Set max value of progress
    Ui_dialog_instance.progressBar.setMaximum(_countTotalPages(pgInstructions))
    
    progressCounter = 0
    for rng in rngList:     
        try:
            tables = camelot_modified.read_pdf(Ui_dialog_instance, src, pages = rng, flavor='stream', edge_tol=500)
            print("Reading:{0}\n>>> Length of tables={1}".format(src, len(tables)))
            
            for i in range(len(tables)):
                # Parse src PDF pages
                table = tables[i]
                df = table.df

                # TODO for cell values that are numbers, convert values to float

                metadataDict = table.parsing_report
                print("@Page={2}, Accuracy={0}, Order={1}, Whitespace={3}".format(metadataDict["accuracy"], \
                                        metadataDict["order"], metadataDict["page"], metadataDict["whitespace"]))
                                
                df.to_excel(writer, sheet_name="Pg.{0} Table{1}".format(metadataDict["page"],metadataDict["order"]))

        except Exception as e:
            errMsg = "Error: page {0} not found in file: {1}; Details={2}".format(rng, fName, e)
            print(errMsg)
            skipped_pages.append(errMsg)   
        progressCounter += 1
        
    # Save output file
    try:          
        # Exception report
        if len(skipped_pages) > 0:
            skippedPgDFRemarks = pd.DataFrame({'Remarks:': skipped_pages})
            skippedPgDFRemarks.to_excel(writer, sheet_name="_skipped pages",index=False)
        
        writer.save()
        print("Ouput file saved to: {0}".format(dest_dir + outputfileName))
        print("Done.")
        #subprocess.run(['open', dest_dir + outputfileName], check=True)
    except Exception as e:
        print("[ERROR] cannot save output file: {0}\nCaught error:{1}.".format(outputfileName, str(e)))

        
    print("Runtime[H:M:S]="+str((datetime.now()-timeZero)))
    
    # Reset progress bar UI
    print("Done : {0}".format(fName))
    Ui_dialog_instance.count_processedPages = 0
    Ui_dialog_instance.count_totalPages = 0
    Ui_dialog_instance.label_9.hide()
    Ui_dialog_instance.progressBar.hide()
    Ui_dialog_instance.pushButton.setEnabled(True)

class SrcFileNotFoundException():
    pass
