import os
import re
import logging
import argparse
import pandas as pd
import PyPDF2
import docx2txt
from configparser import ConfigParser


class TextSearch():
    ''' Class to search keyword in a folder recursively.
        It requires the following parameter by default
        path, searchString and caseSensitive'''

    def __init__(self, path, searchString, ignoreFile=[], ignoreDir=[],
                 caseSensitive=False):

        # dictionary for reading files differently
        self.extentionCommand = {
            "xlsx": self.searchExcelFile,
            "txt": self.searchTextFile,
            "pdf": self.searchPdfFile,
            "doc": self.searchDocFile,
            "docx": self.searchDocFile,
        }

        self.ignoreFile = ignoreFile

        self.ignoreDir = ignoreDir

        # initialising the folder path
        self.path = path
        self.searchString = searchString

        # checking for Case Sensitive flag
        if caseSensitive:
            self.regString = re.compile(self.searchString)
        else:
            self.regString = re.compile(self.searchString, re.IGNORECASE)

    def directoryWalk(self):

        # walking through the directory tree
        for dirName, subdirList, fileList in os.walk(self.path):

            subdirList = [subdir for subdir in subdirList[:]
                          if subdir not in self.ignoreDir]

            fileList = [file for file in fileList[:]
                        if file.split('.')[-1] not in self.ignoreFile]

            for file in fileList:
                # creating file path
                filePath = self.pathMaker(dirName, file)

                # reading different file formats
                self.extentionCommand.get(
                    filePath.split('.')[-1], self.searchTextFile)(filePath)

    def pathMaker(self, dirName, fileName):
        # function to create path of files
        return os.path.join(dirName, fileName)

    def searchTextFile(self, filename):
        '''search keyword in a text file i.e .txt. It reads the
        file and searches each line for the keyword'''

        try:
            # open the file in read mode and
            with open(filename, 'r', encoding='utf-8', errors='ignore') as fin:
                for index, line in enumerate(fin):
                    if self.searchString in line:
                        print(' {} : Line no. {} : {}'.format(
                            filename, index, line))
        except Exception:
            # skip the raise error because of corrupted Excel files
            pass

    def searchExcelFile(self, filename):

        # function to search string in excel files .xlsx
        print('Excel file : {}'.format(filename))
        logging.info('running the search on Excel File')
        try:
            xl = pd.read_excel(filename)
            xl.apply(self.searchSeries, axis=1)
            del xl
        except Exception:
            # skip the raise error because of corrupted Excel files
            pass

    def searchPdfFile(self, filename):
        '''search the keyword in a PDF file i.e .pdf. It uses the
        PyPDF2 library (replace this libary if you find a better one)'''

        try:
            # open the file in binary read mode
            with open(filename, 'rb') as fin:

                # read the pdf file
                pdfReader = PyPDF2.PdfFileReader(fin)
                print('PDF File {}'.format(filename))

                # search each page for the search string
                for pageNum in range(pdfReader.numPages):
                    pageText = pdfReader.getPage(pageNum).extractText()
                    if self.regString.search(pageText):
                        print('Found on Page {}'.format(pageNum))
        except Exception:
            # skip the raise error because of corrupted PDF files
            pass

    def searchDocFile(self, filename):
        '''search the search string in a Word Document i.e .docx. It uses the
        docx2text library (replace this libary if you find a better one)'''

        try:
            # convert the document to string
            docText = docx2txt.process(filename)

            # if the regex matches print the search string and the filename
            if self.regString.search(docText):
                print(' {} Found In Document File :  {}'.format(
                    self.searchString, filename))
        except Exception:
            # skip the raise error because of corrupted Doc files
            pass

    def searchCsvFile(self, filename):
        ''' search the search string in a CSV file. It loads the CSV
        in a Pandas Dataframe and then searches for the search string'''

        print('Csv file : {}'.format(filename))

        try:
            logging.info('running the search on Csv File')

            # load the CSV in the dataframe
            xl = pd.read_csv(filename)

            # search the Dataframe
            xl.apply(self.searchSeries, axis=1)

            # delete the Dataframe from the memory
            del xl
        except Exception:
            # skip the raise error because of corrupted csv files
            pass

    def searchSkipFile(self, fileName):
        ''' skip search in all files .all-other'''
        pass

    def searchSeries(self, xlRow):
        ''' search pandas series for search string'''

        # Convert the pandas series in a string
        strRow = " ".join(xlRow.values.astype(str))

        # search the string for search string
        if self.regString.search(strRow):
            print('Line In Excel : {}'.format(strRow))
            return True
        else:
            return False


if __name__ == '__main__':

    # parse the configuration file
    configParser = ConfigParser()

    configFilePath = os.path.join(os.path.dirname(__file__), 'setting.ini')

    configParser.read(configFilePath)

    # list containing extention of file to skip
    skipFileList = configParser.get('options', 'skipFileList').split(',')

    # argument parser for command line
    parser = argparse.ArgumentParser(
        description='Process the filename and the search string')

    # Directory argument
    parser.add_argument('--dir', metavar='searchDirectory',
                        type=str, nargs='?', default='.',
                        help='directory in which you need to search keyword')

    # Keyword argument
    parser.add_argument('--key', metavar='searchString', required=True,
                        type=str, nargs='?', help='the search string')

    # Case Sensitivity argument
    parser.add_argument('--case', default=False, nargs="?",
                        help='case sensitivity flag')

    # Skip Directory argument
    parser.add_argument('--skipDir', default=[], nargs="?",
                        help='skip Directory argument')

    # Skip Files argument
    parser.add_argument('--skipFile', default=[], nargs="+",
                        help='skip files  with given extentions argument')

    # parse the arguments
    args = parser.parse_args()

    print("Searching the Directory {} for the Keyword {} and Case Sensitivity\
          flag as {} ".format(args.dir, args.key, args.case))

    # get absoulte path of the directory
    dirPath = os.path.abspath(args.dir)

    # merge the two skip file lists
    if args.skipFile:
        skipFileList = skipFileList + args.skipFile

    # create an instance of the TextSearch Class
    search = TextSearch(dirPath, args.key, skipFileList,
                        args.skipDir, args.case)

    # walk the directory and search for the keyword in given files
    search.directoryWalk()

    exit
