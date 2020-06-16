import wx, os, sys, win32com.client, PyPDF2
from time import strftime

class windowClass(wx.Frame):
    
    def __init__(self, *args, **kwargs):
        super(windowClass, self).__init__(*args, **kwargs, size=(600,200))

        self.basicGUI()

    def basicGUI(self):

        # Menu Bar
        panel = wx.Panel(self)
        menuBar = wx.MenuBar()
        fileButton = wx.Menu()

        ExitItem = wx.MenuItem(fileButton, wx.ID_EXIT, 'Zamknij\tCtrl+Q')
        fileButton.Append(ExitItem)

        toolBar = self.CreateToolBar()

        menuBar.Append(fileButton, 'Plik')
    
        self.SetMenuBar(menuBar)
        
        self.Bind(wx.EVT_MENU, self.Quit, ExitItem)
        # Menu Bar end
        
        wx.StaticText(panel, -1, 'Ścieżka folderu:', (5,0))
        self.pathFolder = wx.TextCtrl(panel, pos=(5, 25), size=(550, 18))   

        self.convertBtn = wx.Button(panel, label = 'Konwertuj do pdf', pos = (5, 50), size = (100,50))
        self.Bind(wx.EVT_BUTTON, self.convert, self.convertBtn)

        self.mergeBtn = wx.Button(panel, label = 'Połącz pdf', pos = (105, 50), size = (100,50))
        self.Bind(wx.EVT_BUTTON, self.merger, self.mergeBtn)
              
        self.SetTitle('Doc2Pdf Merger')
        self.Show(True)

    def Quit (self, e):
        self.Close()
      
    def convert(self, e):

        pathFolder = self.pathFolder.GetValue()

        def count_files(filetype):
            
            count_files = 0
            for files in os.listdir(pathFolder):
                filesLow = files.lower()
                if filesLow.endswith(filetype):
                    count_files += 1
            return count_files

        # Function "check_path" is used to check whether the path the user provided does
        # actually exist. The user is prompted for a path until the existence of the
        # provided path has been verified.

        def check_path(prompt):
        ##    (str) -> str
        ##    Verifies if the provided absolute path does exist.
            
            if os.path.exists(pathFolder) != True:
                wrongPath = wx.MessageBox("The specified path does not exist.",\
                    "Error", style = wx.OK_DEFAULT | wx.ICON_ERROR | wx.CENTRE)

            return pathFolder    
            
        pathFolder = check_path("Provide absolute path for the folder: ")

        # Change the directory.

        os.chdir(pathFolder)

        # Count the number of docx and doc files in the specified folder.

        num_docx = count_files(".docx")
        num_doc = count_files(".doc")

        # Check if the number of docx or doc files is equal to 0 (= there are no files
        # to convert) and if so stop executing the script. 

        if num_docx + num_doc == 0:
            print("\nThe specified folder does not contain docx or docs files.\n")
            print(strftime("%H:%M:%S"), "There are no files to convert. BYE, BYE!.")
            exit()
        else:
            print("\nNumber of doc and docx files: ", num_docx + num_doc, "\n")
            print(strftime("%H:%M:%S"), "Starting to convert files ...\n")

        # Try to open win32com instance. If unsuccessful return an error message.

        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False

            for f in os.listdir(pathFolder):
                in_file = os.path.abspath(pathFolder + "\\" + f)
                print(in_file)
                if in_file.lower().endswith(".docx"):
                    new_name = f.replace(in_file[-5:], r".pdf")
                    new_file = os.path.abspath(pathFolder + "\\" + new_name)
                    doc = word.Documents.Open(in_file)
                    print(strftime("%H:%M:%S"), " docx -> pdf ", os.path.relpath(new_file))
                    doc.SaveAs(new_file, FileFormat = 17)
                    doc.Close()
                if in_file.lower().endswith(".doc"):
                    new_name = f.replace(in_file[-4:], r".pdf")
                    new_file = os.path.abspath(pathFolder + "\\" + new_name)
                    doc = word.Documents.Open(in_file)
                    print(strftime("%H:%M:%S"), " doc  -> pdf ", os.path.relpath(new_file))
                    doc.SaveAs(new_file, FileFormat = 17)
                    doc.Close()

        except:
            print("Coś poszło nie tak.")
        finally:
            word.Quit()

        print("\n", strftime("%H:%M:%S"), "Finished converting files.")  

        # Count the number of pdf files.

        num_pdf = count_files(".pdf")   

        print("\nNumber of pdf files: ", num_pdf)

        # Check if the number of docx and doc file is equal to the number of files.

        if num_docx + num_doc == num_pdf:
            print("\nNumber of doc and docx files is equal to number of pdf files.")
        else:
            print("\nNumber of doc and docx files is not equal to number of pdf files.")

    def merger(self, e):
        def load_pdf (filename):
            f = open(filename, 'rb')
            return(PyPDF2.PdfFileReader(f))

        pathFolder = self.pathFolder.GetValue()
        
        def add_to_writer(pdf, writer):
            for i in range(pdf.getNumPages()):
                writer.addPage(pdf.getPage(i))

        writer = PyPDF2.PdfFileWriter()
        out_file = pathFolder + '\\Merged_pdf.pdf'
        output_file = open(out_file, 'wb')

        for f in os.listdir(pathFolder):
            in_file = os.path.abspath(pathFolder + "\\" + f)
            val = pathFolder + '\\' + f
            if f.lower().endswith(".pdf") and val != out_file:
                pdfFile = load_pdf(in_file)
                pdfPages = pdfFile.getNumPages()
                add_to_writer(pdfFile, writer)

        writer.write(output_file)
        output_file.close()

def main():
    app = wx.App()
    windowClass(None)
    
    app.MainLoop()
    
main()
