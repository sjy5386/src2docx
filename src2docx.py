import argparse
import docx
import os


class Src:
    def __init__(self, filename):
        self.filename = filename
        f = open(filename, "r", encoding="utf-8")
        self.content = f.read()
        f.close()


def searchSrc(dirname, array):
    cwd = os.getcwd()
    for (path, dirs, files) in os.walk(dirname):
        for filename in files:
            os.chdir(path)
            print(filename)
            f = Src(str(filename))
            array.append(f)
    os.chdir(cwd)


def makeTable(doc, src):
    table = doc.add_table(rows=2, cols=1)
    table.rows[0].cells[0].text = src.filename
    table.rows[1].cells[0].text = src.content
    table.style = "Table Grid"
    doc.add_paragraph()


if __name__ == "__main__":
    argumentParser = argparse.ArgumentParser()
    argumentParser.add_argument("directory", type=str, help="directory")
    argumentParser.add_argument("output", type=str, help="output")
    args = argumentParser.parse_args()
    directory = args.directory
    output = args.output
    files = []
    searchSrc(directory, files)
    document = docx.Document()
    for file in files:
        makeTable(document, file)
    document.save(output)
