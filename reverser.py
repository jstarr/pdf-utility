# get_doc_info.py
from PyPDF2 import PdfFileReader, PdfFileWriter
import PySimpleGUI as sg
from datetime import datetime
from time import strptime
import time


path_source = '/Users/jimstarr/Shir Tikvah/CST-SIDDUR-FINALhi-res.pdf'
sourceTip = 'The path to the file to be reversed.'
pathDest = '/Users/jimstarr/Shir Tikvah/CST-SIDDUR-FINALhi-res-forward.pdf'
destTip = 'The path for the output file.'
reverseButtonTip = 'Start the process'
monthNames = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

def convertDate(info):
    '''Convert the PDF file's date format, (D:YYYYMMDDHHmmSSOHH'mm'), to a datetime object'''

    datestring = info['/CreationDate'][2:-7]
    date_object = datetime.strptime(datestring, '%Y%m%d%H%M%S')
    return date_object

def displaySourceData(window, fsource):
    '''Post the metadata for a pdf file'''

    '''NOTE: setting strict to False prevents PdfReadWarning.  See https://github.com/mstamy2/PyPDF2/issues/36 '''
    PDFSource = PdfFileReader(fsource, strict=False)
    info = PDFSource.getDocumentInfo()

    window['-NUMPAGES-'].update(PDFSource.getNumPages())
    window['-FIELDS-'].update(PDFSource.getFields(fsource))
    window['-ENCRYPTED-'].update(PDFSource.isEncrypted)
    window['-AUTHOR-'].update(info.author)
    window['-CREATIONDATE-'].update(str(convertDate(info)))
    window['-CREATOR-'].update(info.creator)
    window['-PRODUCER-'].update(info.producer)
    window['-SUBJECT-'].update(info.subject)
    window['-TITLE-'].update(info.title)
    return PDFSource

def buildWindow():
    ofile_tip = 'Full path to the INPUT file.'
    dfile_tip = 'Full path to the OURPUR file.'
    layout = [[sg.Text('PDF Reverse file')],
              [sg.Text('Original File: '), sg.InputText(size=(40, 1), key='-OFILE-',default_text=path_source, enable_events=False, tooltip=ofile_tip)],
              [sg.Text('Dest File: '), sg.InputText(size=(40, 1), key='-DFILE-', default_text=pathDest, enable_events=False, tooltip=dfile_tip)],
              [sg.Text('Page: '), sg.Text(size=(9, 1), key='-PAGENO-')],
              [sg.Text('Number of Pages: '), sg.Text(size=(9, 1), key='-NUMPAGES-')],
              [sg.Text('Fields: '), sg.Text(size=(30, 1), key='-FIELDS-')],
              [sg.Text('Encrypted: '), sg.Text(size=(6, 1), key='-ENCRYPTED-')],
              [sg.Text('Author: '), sg.Text(size=(40, 1), key='-AUTHOR-')],
              [sg.Text('Creation Date: '), sg.Text(size=(40, 1), key='-CREATIONDATE-')],
              [sg.Text('Creator: '), sg.Text(size=(40, 1), key='-CREATOR-')],
              [sg.Text('Producer: '), sg.Text(size=(40, 1), key='-PRODUCER-')],
              [sg.Text('Subject: '), sg.Text(size=(40, 1), key='-SUBJECT-')],
              [sg.Text('Title: '), sg.Text(size=(40, 1), key='-TITLE-')],
              [sg.MLine(key='-ML1-'+sg.WRITE_ONLY_KEY, size=(50,8), reroute_cprint=True, enable_events=False)],
              [sg.Cancel(), sg.Button(button_text='Reverse', key='-REVERSE-', tooltip=reverseButtonTip)]
         ]
    window = sg.Window('Reverse PDF File', layout, return_keyboard_events=True, resizable=True)

    return window

def reverse_file(path_source, path_destination, window):
    # try:
    with open(path_source, 'rb') as fsource:
        PDFSource = displaySourceData(window, fsource)
        pdf_writer = PdfFileWriter()
        # try:
        with open(path_destination, 'wb') as dSource:
            for page_no in range(PDFSource.getNumPages()-1, -1, -1):
                window['-PAGENO-'].update(page_no)
                window.Finalize()

                try:
                    current_page = PDFSource.getPage(page_no)
                    pdf_writer.addPage(current_page)
                except IndexError:
                    sg.cprint(f'Problem in page {page_no}', c=('red on white'), key='-ML1-')
            pdf_writer.write(dSource)
    sg.cprint(f'Created output file {path_destination}\nfrom {path_source}.', c=('blue on white'))


if __name__ == '__main__':
    window = buildWindow()
    try:
        with open(path_source, 'rb') as fsource:
            while True:
                event, values = window.read()

                if event == sg.WIN_CLOSED or event == 'Exit' or event == 'Cancel':
                    break

                if event == '-REVERSE-':
                    path_source = values['-OFILE-']
                    path_destination = values['-DFILE-']
                    reverse_file(path_source, path_destination, window)
                else:
                    if (event==None  or  event==''):
                        print('None Event')
                    path_source = values['-OFILE-']
                    with open(path_source, 'rb') as fsource:
                        PDFSource = displaySourceData(window, fsource)

    finally:
        window.close()
