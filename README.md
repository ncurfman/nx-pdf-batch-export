# nx-pdf-batch-export
Easily export PDF files for all loaded NX files

# What this script does:
This NX Journal file will look at all files loaded in NX and export any associated drawing files to PDFs. This tool can be used to export a very specific list of drawing files, or every single PDF associated with a top level assembly model. Before exporting, you'll have the option to select an output folder and set watermark text.

# How it works:
The script looks for any part or assembly files that are loaded and opens all the related drawing files. After that, the script will step through every open drawing file and export it to a PDF. The script supports multiple drawing datasets per model or assembly, but will only look for drawing datasets with "DWG" in the name. If you are using NX and Teamcenter in the normal fashion, this limitation won't affect you, but be careful if you have any drawing datasets with custom names.

# How to use it:
Download the batch export NX journal by going to the source code [here](https://github.com/ncurfman/nx-pdf-batch-export/blob/main/NX_Export_PDF_All_Open_Parts_multiple_drawings.vb)  and then clicking the "Download raw file" button on the top right.

You may receive a warning that this filetype could be dangerous, as Visual Basic code files (what this journal is written in) can be used for malicious pruposes. Since this file is coming from a trusted source (a colleague at Fermilab) this is not an issue, however if you have any concerns you should reach out to the Fermilab IT team for clarification or contact the author at ncurfman@fnal.gov.

Move the file from Downloads to somewhere you'll be able to find it later, like your Documents folder. 

Next, open NX through Teamcenter. From here, follow either the "Generate PDFS for a list of files" instructions, or the "Generate all drawings associated with an assembly" instructions.

# Known limitations:
* Loading drawings with 3D centerlines using "Structure Only" as described in the "Generate PDFs for a list of files" causes the 3D centerlines to disassociate since the referenced geometry is not loaded. Production drawings with 3D centerlines will need to be opened with "Fully Load" and exported manually.

# Generate PDFs for a list of files
* First, make sure you've closed all open parts in NX to prevent PDFs of loaded but not displayed parts from being generated unintentionally by going to File -> Close -> All Parts.

* Next, set your assembly load options as shown below. Make sure "Load" is set to "Structure Only" and "Load Interpart Data" is UNchecked. as shown below:

![Export only files intentionally opened](https://github.com/user-attachments/assets/e1be0678-04e7-4cf8-8fc9-e669884e5007)

* Now, in Teamcenter, highlight any parts you would like a PDF of, right click, and select copy.

* Go to NX and open the Teamcenter Navigator --> Open the Clipboard and select Refresh. The Items that were selected are now open in the Teamcenter navigator.

* From the clipboard menu, select the top item, then while holding the Shift  key, select the bottom item to select all the items in the clipboard.

* Right click on the highlighted files and select "Open". All models will open in NX.

* Go to the "Tools" tab and click the "Play" button in the "Journal" section.

* Click browse, and navigate to the script you downloaded previously.

* With the script shown under "File Name" click the "Run" button. You will prompted to select an output folder and add a watermark if desired.

* After the watermark prompt the script will begin opening drawings and creating PDFs. Please be patient and do not use your computer while PDF generation is running.

* All files should now have PDFs in the output folder you selected.

# Generate all drawings assocatied with an assembly:

Be Careful! This procedure generates a file for every part that's loaded in NX including ALL assembly child components! That can be a lot of PDFs!

* First, make sure you've closed all open parts in NX to prevent PDFs of loaded but not displayed parts from being generated unintentionally by going to File -> Close -> All Parts.

* Next, set your assembly load options as shown below. Make sure "Load" is set to "All Components" and "Load Interpart Data" is checked. as shown below:

![Export all loaded parts](https://github.com/user-attachments/assets/49700d9c-a0e4-4d33-a8c9-92e49f12222e)

* Open the assembly you would like to generate PDFs from in NX.

* Go to the "Tools" tab and click the "Play" button in the "Journal" section.

* Click browse, and navigate to the script you downloaded previously.

* With the script shown under "File Name" click the "Run" button. You will prompted to select an output folder and add a watermark if desired.

* After the watermark prompt the script will begin opening drawings and creating PDFs. Please be patient and do not use your computer while PDF generation is running.

* All files should now have PDFs in the output folder you selected.
