import keyboard
import os
import configparser
import xml.etree.ElementTree as ET
from colorama import Fore, Style, init, AnsiToWin32
from htmldocx import HtmlToDocx
import re
import os
import shutil  # save img locally
import uuid
from datetime import datetime
import requests  # request img from web
from bs4 import BeautifulSoup

# Read config
def ReadConfig(filepath):
    print(f'{Fore.YELLOW}Start ReadConfig')

    if os.path.exists(filepath):
        config = configparser.ConfigParser()
        config.read(filepath, "utf8")
        config.sections()

        global glhost
        glhost = config.has_option("Settings", "host") and config.get("Settings", "host") or None

        global glapikey
        glapikey = config.has_option("Settings", "apikey") and config.get("Settings", "apikey") or None

        return True
    else:
        print(f'{Fore.YELLOW}Start create_config')
        config = configparser.ConfigParser()
        config.add_section("Settings")

        config.set("Settings", "host", 'http://192.168.1.1')
        config.set("Settings", "apikey", 'dq3inqgnqe8igqngninkkvekmviewrgir9384')

        with open(filepath, "w") as config_file:
            config.write(config_file)

        print(f'{Fore.GREEN}Create config: {Fore.BLUE}{filepath},{config_file}')
        return False

# Write DOCX
def WriteDocx(issueDescription, filename):
    print(f'{Fore.YELLOW}Write *.docx: {filename}')

    new_parser = HtmlToDocx()
    docx = new_parser.parse_html_string(issueDescription)
    docx.save(filename)

# Download IMG
def DownloadIMG(description, imgScrs, downloadDirectory, imagesavelist):
    for item in imgScrs:
        imgScr = item.get("src")

        pathImg = f'{glhost}/{imgScr}'
        if imgScr.startswith('http'):
            pathImg = imgScr

        print(f'{Fore.BLUE}{pathImg=}')

        xfilename, xfile_extension = os.path.splitext(pathImg)
        newImgName = f'{uuid.uuid4()}{xfile_extension}'
        newImPath = os.path.join(downloadDirectory, newImgName)
        imagesavelist.append(newImPath)  # save list download Img file
        print(f'{Fore.BLUE}Download Img: {pathImg} -> {newImPath=}')

        dowloadLink = f"{pathImg}?key={glapikey}"
        if imgScr.startswith('http'):
            owloadLink = f"{pathImg}"

        print(f'{dowloadLink=}')
        print(f'{Fore.MAGENTA}Replace: {str(imgScr)} to {str(newImPath)}')

        # Download img
        res = requests.get(dowloadLink, stream=True)
        if res.status_code == 200:
            with open(newImPath, 'wb') as f:
                shutil.copyfileobj(res.raw, f)

        # Replace Img link to local path
        description = description.replace(str(imgScr), str(newImPath))

    return description


#  1) Delete emty string
#  2) LowDown <h > на 1
def editblock(curblok):
    sresult = ""
    for curstr in curblok.split('\n'):
        if curstr.strip() != '':
            hlist = re.findall(r'<h\d*', curstr)
            if len(hlist) > 0:
                hitem = hlist[0].replace("<h", "").strip()
                newhitem = int(hitem) + 1
                curstr = curstr.replace(f"<h{hitem}", f"<h{newhitem}").replace(f"/h{hitem}>", f"/h{newhitem}>")
            sresult = sresult + curstr + '\n'

    return sresult


def processfile(cssfilename, filename):
    # necessarily namespace for find elements
    namespaces = '{http://www.w3.org/2005/Atom}'

    tree = ET.parse(cssfilename)
    root = tree.getroot()
    text = "<html><body>"

    #  find all entry
    for entry in root.findall(f'{namespaces}entry'):
        for blok in entry:

            #  Get title every one entry - it H1
            if blok.tag == f'{namespaces}title':
                text = text + f'<h1>{blok.text.strip()}</h1>'

            #  body
            if blok.tag == f'{namespaces}content':
                curblok = blok.text
                text = text + editblock(curblok)

    text = text + "</html></body>"

    text = text.replace("&para;", "")
    text = text.replace("<p> </p>", "")
    text = text.replace('<a name="XSD"></a>', "")

    now = datetime.now()
    downloadDirectory = os.path.join(currentDirectory, str(f'{now.strftime("%Y-%m-%d")}-{now.strftime("%H-%M-%S")}'))
    if not os.path.exists(downloadDirectory):
        os.makedirs(downloadDirectory)
        print(f"{Fore.CYAN}{downloadDirectory=}")

    imagesavelist = []
    # Process image
    # 2.1 Collect Scr list
    soup = BeautifulSoup(text, "lxml")
    imgScrs = soup.findAll("img")

    # 2.2 Download image from Src
    text = DownloadIMG(text, imgScrs, downloadDirectory, imagesavelist)

    # Delete exist file
    docxfilename = f'{downloadDirectory}\\{filename}.docx'
    if os.path.exists(docxfilename):
        os.remove(docxfilename)

    # Write new file
    print(f'{Fore.GREEN}\n\nCreate: {docxfilename}')
    WriteDocx(text, docxfilename)

    # Delete Img file
    for file in imagesavelist:
        if os.path.isfile(file):
            os.remove(file)

def main():
    print(f'{currentDirectory=}')
    for file in os.listdir(currentDirectory):
        if os.path.isfile(file) and file.lower().endswith(".css"):
            processfile(f'{currentDirectory}\\{file}', file.split('.')[0])


if __name__ == "__main__":
    print(f"{Fore.CYAN}Last update: Cherepanov Maxim masygreen@gmail.com (c), 01.2022")
    print(f"{Fore.CYAN}Convert file *.html to *.docx")

    currentDirectory = os.getcwd()
    configfilepath = os.path.join(currentDirectory, 'config.cfg')

    if ReadConfig(configfilepath):
        main()
    else:
        print(f'{Fore.RED}Pleas edit default Config value: {Fore.BLUE}{configfilepath}')
        keyboard.wait("space")