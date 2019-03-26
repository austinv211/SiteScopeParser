import pdfkit
import sys
from bs4 import BeautifulSoup
from os import remove

#edit the html title
def editHTML(fileName):
    try:
        html_doc = open(fileName, "r+")
        soup = BeautifulSoup(html_doc,'html.parser')
        res = soup.find("h2")
        res.contents[0].replaceWith("Azure Monthly Exception Report")
        html_doc.close()
        html = soup.prettify("utf-8")
        with open("HTML/editedHTML.html", "wb") as file:
            file.write(html)
        return 0
    except:
        print("Error Editing HTML file", file=sys.stderr)
        exit(1)

#convert to pdf function
def convertToPDF(outputPath):
    try:
        pdfkit.from_file("HTML/editedHTML.html", outputPath)
        remove("HTML/editedHTML.html")
        return 0
    except:
        print("Error converting to PDF, please make sure html file exists and PDF is not open")
        exit(1)
    
if __name__ == "__main__":
    #check the commandline args
    if len(sys.argv) != 3:
        print("Incorrect Number of Args provided", file=sys.stderr)
        exit(1)
    if sys.argv[1].endswith(".html"):
        print(sys.argv[1])
    else:
        print("First Arg Must be an html file", file=sys.stderr)
        exit(1)
    if sys.argv[2].endswith(".pdf"):
        print(sys.argv[2])
    else:
        print("Second arg must be a pdf file", file=sys.stderr)
        exit(1)
    #edit the html
    if editHTML(sys.argv[1]) == 0:
        print("Completed HTML Edit")
    #call convert to pdf
    if convertToPDF(sys.argv[2]) == 0:
        print("DONE!")
