import openpyxl
import win32com.client as win32
import shutil
import os
import fitz  # PyMuPDF
import time
from datetime import datetime


# Get the current working directory
cwd = os.getcwd()



def get_all_files(folder):
    all_files = []
    for root, dirs, files in os.walk(folder):
        for file in files:
            all_files.append(os.path.join(root, file))
    return all_files

def filter_files(files, keyword):
    return [file for file in files if keyword in file]

def findReplace(wApp, wDoc, find_str, replace_str='none'):
    ''' replace all occurrences of `find_str` w/ `replace_str` in `word_file`
        or alternatively, finds a text and selects that'''
    wCont = wDoc.Content
    wdFindContinue = 1
    wdReplaceAll = 2
    if replace_str == 'none':
        wApp.Selection.Find.Execute(find_str, False, False, False, False, False, True, wdFindContinue)
        return wApp.Selection
    else:
        wCont.Find.Execute(find_str, False, False, False, False, False, True, wdFindContinue, False, replace_str, wdReplaceAll)


def create_word_documents(sMeetsetFolder, location, weeks_with_year, knmi_data_used, pWord, retry_count=1):
    # filepath_datadir = cwd + '/ExcelCalculations/Regulier/Uurwaarden/1hour - RV - weekno - 36Energy Balance - 02-09-2024 - 08-09-2024.xlsx'
    excel_filename = 'Uitwerk light uurbasis zonder koeler RM.xlsm'
    excel_filepath = os.path.abspath(os.path.join(cwd, 'Automatic excel calculations', 'Input', excel_filename))
    # Construct the TrustedDocuments path
    cwd_parts = cwd.split('\\')
    trusted_documents_dir = '\\'.join(cwd_parts[:3] + ['Documents'] + ['TrustedDocuments'])

    filepath_datadirs = []
    # Process the weeks in week_with_year:
    for week in weeks_with_year:
        # Create the path to the Excel file
        created_excel_output_folder = os.path.join(cwd, 'Automatic excel calculations', 'Input', sMeetsetFolder, week)
        files_in_folder = get_all_files(created_excel_output_folder)
        filepath_datadir = filter_files(files_in_folder, '1hour - RV')  # Filter files to include only those that contain '1hour - RV'
        if len(filepath_datadir) == 0: # Check if the folder is empty
            print(f"No files found in '{created_excel_output_folder}' that contain '1hour - RV'.")
            continue
            
        # Append the file to the list
        filepath_datadirs.append(filepath_datadir[0])
    if len(filepath_datadirs) == 0:
        print(f"No files found in {created_excel_output_folder} that contain '1hour - RV'.")

    output_folder = os.path.join(pWord, sMeetsetFolder)
    template_file = os.path.join(cwd, 'Automatic excel calculations', 'Input', 'Template document.docx')
    # Create the output directory if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    xlApp = None
    wApp = None

    # Loop throught the weeks specified in all_weeks, from HPImport.py:
    for filepath_datadir in filepath_datadirs:
        print(f"Processing file: {filepath_datadir}")
        attempt = 0
        finished = False
        while (attempt <= retry_count) and not finished:
            try:
                # For DNV laptops, Excel macros can only be used if the Excel file is in the TrustedDocuments directory
                # However, this directory does not exist for the 'meetlaptops' and in that case is the TrustedDocuments directory not needed
                
                # Therefore, check if the TrustedDocuments directory exists, if not than use a temporary directory
                if os.path.exists(trusted_documents_dir):
                    # Copy the Excel file to the TrustedDocuments directory
                    filename_data = os.path.join(trusted_documents_dir, excel_filename)          
                else:
                    filename_data = os.path.join(cwd, 'Automatic excel calculations', 'TemporaryStoredFiles', excel_filename)

                # Delete the file if it already exists
                if os.path.exists(filename_data):
                    os.remove(filename_data)

                # Wait for 2 seconds to observe the changes
                time.sleep(2)
                shutil.copy2(excel_filepath, filename_data)

                filepath_exceltool = os.path.abspath(filename_data)

                # Open the Excel file
                xlApp = win32.Dispatch('Excel.Application')
                wb = xlApp.Workbooks.Open(filename_data)
                xlApp.Visible = True
                wb.Application.Run('Start.plaatspad')
                wb.Application.Run('Datain.data_in', filepath_datadir, excel_filename)  # Pass the file path as an argument

                # Wait for 10 seconds to observe the changes
                time.sleep(10)
                wb.Application.ScreenUpdating = True
                xlApp.Visible = True

                # Run the second macro
                print("Running Grafieken.pasgrafiekenaan")
                wb.Application.Run('Grafieken.pasgrafiekenaan')

                # Get the 'Stat' worksheet
                ws = wb.Worksheets('Stat')

                # Copy the data from the worksheet, including the formatting
                # ws.Range("A1").CurrentRegion.CopyPicture()
                ws.Range("B1:O74").CopyPicture()

                # Create a new Word application
                wApp = win32.Dispatch('Word.Application')
                wApp.Visible = True
                # Get the week number from the "Control" sheet
                ws_control = wb.Worksheets('Control')
                weekNo = str(int(ws_control.Range("B4").Value))

                # Use a temporary output file for intermediate operations
                temp_output_file = os.path.join(output_folder, 'temp_output.docx')

                # Load the template document
                shutil.copy2(template_file, temp_output_file)
                wDoc = wApp.Documents.Open(temp_output_file)
                current_day = str(datetime.now().strftime('%d-%m-%Y'))
                findReplace(wApp, wDoc, '{DATE}', current_day)
                findReplace(wApp, wDoc, '{LOCATION}', location)

                # Insert the week number into the Word document
                findReplace(wApp, wDoc, '{WEEK_NUMBER}', weekNo)
                
                if knmi_data_used:
                    findReplace(wApp, wDoc, '{NOTE_KNMI}', 'Note that due to missing weather data, KNMI data was used.')
                else:
                    findReplace(wApp, wDoc, '{NOTE_KNMI}', '')

                # Insert the Stat sheet at the location marked by '{STAT_SHEET}'
                findReplace(wApp, wDoc, '{STAT_SHEET}')
                wApp.Selection.Paste()

                temp_dir = os.path.join(cwd, 'Automatic excel calculations', 'TemporaryStoredFiles')
                if not os.path.exists(temp_dir):
                    os.makedirs(temp_dir)
                # # Save the Excel file as a PDF
                # temporary_pdf_path = os.path.join(temp_dir, 'temp.pdf')  
                # wb.ExportAsFixedFormat(0, temporary_pdf_path)

                # # Open the PDF file
                # doc_pdf = fitz.open(temporary_pdf_path)

                # Insert the PDF pages as images at the location marked by '{PDF_PAGES}'
                findReplace(wApp, wDoc, '{PDF_PAGES}')


                # for page_num in range(doc_pdf.page_count - 3, doc_pdf.page_count):
                #     page_pdf = doc_pdf.load_page(page_num)
                #     mat = fitz.Matrix(300 / 72, 300 / 72)  # Create a zoom matrix
                #     pix = page_pdf.get_pixmap(matrix=mat)  # Get the pixmap with the zoom matrix
                #     img_file = os.path.join(temp_dir, f"page_{page_num + 1}.png")
                #     pix.save(img_file)
                #     pic = wApp.Selection.InlineShapes.AddPicture(img_file)
                #     #pic.LockAspectRatio = 0  # Allow the picture to be resized
                #     #pic.Width = wDoc.PageSetup.PageWidth - wDoc.PageSetup.LeftMargin - wDoc.PageSetup.RightMargin
                #     #pic.Height = pic.Width * pic.Height / pic.Width
                #     wApp.Selection.TypeParagraph()

                lstCharts = ["Chart1", "Chart2", "Chart3"]
                for sChartName in lstCharts:
                    wb.Charts(sChartName).ChartArea.Copy()
                    wApp.Selection.PasteSpecial(IconIndex=0, Link=False, Placement=0, DisplayAsIcon=False, DataType=9)
                    wApp.Selection.TypeParagraph()
                # Save the Word document
                wDoc.Save()

                # Close the Word document before renaming
                wDoc.Close(SaveChanges=0)
                wApp.Quit()
                del wApp  # Clean up the Word COM object
                
                # Get the directory name
                directory_name_excel = os.path.dirname(filepath_datadir)

                # Get the folder name
                year_weekstr = os.path.basename(directory_name_excel)
                if not os.path.exists(os.path.join(output_folder, year_weekstr)):
                    os.makedirs(os.path.join(output_folder, year_weekstr))

                # Rename the temporary output file to the final output file
                final_output_file = os.path.join(output_folder, year_weekstr, f"Weekrapport-{location}-week{weekNo}.docx")
                if os.path.exists(final_output_file):
                    os.remove(final_output_file)
                os.rename(temp_output_file, final_output_file)

                # # Close the PDF file
                # doc_pdf.close()

                # Close the Excel file
                fpathExcelOutput = os.path.join(output_folder,year_weekstr,f"Week{weekNo}-{excel_filename}")
                if os.path.exists(fpathExcelOutput):
                    os.remove(fpathExcelOutput)
                wb.SaveAs(fpathExcelOutput)
                wb.Close(SaveChanges=False)
                xlApp.Quit()
                del xlApp  # Clean up the Excel COM object

                # Delete all files in the TemporaryStoredFiles directory
                for file in os.listdir(temp_dir):
                    file_path = os.path.join(temp_dir, file)
                    if os.path.isfile(file_path):
                        os.remove(file_path)

                print('Sleep for 2 seconds such that excel can be closed')
                time.sleep(2)

                # Delete the file in TrustedDocuments at the very end
                try:
                    os.remove(filename_data)
                except FileNotFoundError:
                    print(f"File {filename_data} not found, cannot delete")

                attempt = 0  # Reset the attempt counter
                finished = True

            except Exception as e:
                print(f"An error occurred: {e}")
                attempt += 1
                print("Attempting to close the Excel and Word applications.")
                try:
                    if xlApp is not None:
                        wb.Close(SaveChanges=False)
                        xlApp.Quit()
                        del xlApp
                    if wApp is not None:
                        wApp.Quit()
                        del wApp  # Clean up the Word COM object
                except Exception as e:
                    print(f"An error occurred while trying to close the Excel and Word applications: {e}")
                if attempt > retry_count:
                    print("Max retries reached. Exiting.")
                else:
                    print(f"Retrying... ({attempt}/{retry_count})")
                    time.sleep(5)  # Wait for 5 seconds before retrying