from src2docx import *
import tkinter.filedialog
import tkinter.messagebox
import tkinter.ttk


class MainForm(tkinter.Tk):
    def __init__(self):
        super().__init__()
        self.title("src2docx")
        self.geometry("220x180")
        self.resizable(0, 0)
        self.directoryLabel = tkinter.ttk.Label(self, text="소스 코드가 있는 폴더")
        self.directoryEntry = tkinter.ttk.Entry(self)
        self.directoryBrowseButton = tkinter.ttk.Button(self, text="찾아보기", command=self.onDirectoryBrowseButtonClicked)
        self.outputLabel = tkinter.ttk.Label(self, text="Word 문서 파일의 이름")
        self.outputEntry = tkinter.ttk.Entry(self)
        self.outputBrowseButton = tkinter.ttk.Button(self, text="찾아보기", command=self.onOutputBrowseButtonClicked)
        self.src2DocxButton = tkinter.ttk.Button(self, text="src2docx", command=self.onSrc2DocxButtonClicked)
        self.directoryLabel.pack()
        self.directoryEntry.pack()
        self.directoryBrowseButton.pack()
        tkinter.ttk.Label(self, text="↓").pack()
        self.outputLabel.pack()
        self.outputEntry.pack()
        self.outputBrowseButton.pack()
        self.src2DocxButton.pack(side="bottom")
        tkinter.ttk.Separator(self).pack(side="bottom", fill="x")
        self.mainloop()

    def onDirectoryBrowseButtonClicked(self):
        directory = tkinter.filedialog.askdirectory()
        self.directoryEntry.delete(0, tkinter.END)
        self.directoryEntry.insert(0, directory)

    def onOutputBrowseButtonClicked(self):
        filename = tkinter.filedialog.asksaveasfilename(filetypes=(("Word 문서", "*.docx"), ("모든 파일", "*.*")))
        self.outputEntry.delete(0, tkinter.END)
        self.outputEntry.insert(0, filename)

    def onSrc2DocxButtonClicked(self):
        directory = self.directoryEntry.get().strip()
        output = self.outputEntry.get().strip()
        if directory == "" or output == "":
            tkinter.messagebox.showerror(title="src2docx", message="값을 입력해주세요.")
            return
        src2Docx = Src2Docx(directory, output)
        src2Docx.run()
        tkinter.messagebox.showinfo(title="src2docx", message="완료되었습니다.")


if __name__ == "__main__":
    mainForm = MainForm()
