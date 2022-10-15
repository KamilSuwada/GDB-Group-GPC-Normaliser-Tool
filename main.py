from openpyxl import Workbook, load_workbook
from xlsxfile import XLSXFile
from datamanipulator import DataManipulator
from tkinter.messagebox import askquestion
from tkinter import Tk
from tkinter.filedialog import askdirectory, askopenfile, askopenfilename, askopenfilenames, asksaveasfilename
import os

Tk().withdraw()
CWD = os.path.dirname(__file__)
DAT = DataManipulator()








def main():
    paths = askopenfilenames()
    files = [XLSXFile(path=path) for path in paths]
    MAN = DataManipulator()
    start = 0
    stop = 28
    
    for file in files:
        MAN.extract_data(file=file, start=start, stop=stop)




if __name__ == "__main__":
    main()