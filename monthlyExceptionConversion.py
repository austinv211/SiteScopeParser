import pdfkit
import sys
from bs4 import BeautifulSoup
from os import remove

#function to change header
def changeHeader(soup):
    #find our header
    header = soup.find("h2")
    header.contents[0].replaceWith("Azure Monthly Exception Report")

def removeDataTables(soup):
    tables = soup.select(".data-table")
    for i in range(0, len(tables)):
        if i != 2:
            tables[i].decompose() 

#function to change the logo
def changeLogo(soup):
    #find the h1
    h1 = soup.find("h1")
    img_soup = BeautifulSoup("<img src='https://perspecta.com/sites/default/files/perspecta_logo_tm_0.png' height='100' >", 'html.parser')
    img = img_soup.img
    h1.replaceWith(img)

#function to remove ping monitor graphs
def removePingMonitorGraphs(soup):
    h3_ping_find = soup.select("center > h3")
    for h3 in h3_ping_find:
        if "Ping Monitor" in h3.contents[0]:
            parent = h3.find_parent("center")
            parent.decompose()

#edit the html title
def editHTML(fileName):
    try:
        #open the doc and create a soup out of it
        html_doc = open(fileName, "r+")
        soup = BeautifulSoup(html_doc,'html.parser')
        #change the header
        changeHeader(soup)
        #change the logo
        changeLogo(soup)
        #remove ping graphs
        removePingMonitorGraphs(soup)
        #remove data tables
        removeDataTables(soup)
        #close and format the doc
        html_doc.close()
        html = soup.prettify("utf-8")
        #write output as editedHTML.html
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
