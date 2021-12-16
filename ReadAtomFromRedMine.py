import os
import xml.etree.ElementTree as ET
from colorama import Fore, Style, init, AnsiToWin32
from htmldocx import HtmlToDocx
import re


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
    htmlfilename = f'{cur_dir}\\{filename}.html'
    if os.path.exists(htmlfilename):
        os.remove(htmlfilename)

    print(f'{Fore.YELLOW}\n\nProcess: {cssfilename} => {htmlfilename}')
    namespaces = '{http://www.w3.org/2005/Atom}'  # обязательно namespace для поиска elements

    tree = ET.parse(cssfilename)
    root = tree.getroot()
    text = "<html><body>"

    #  проходим по всем записям entry
    for entry in root.findall(f'{namespaces}entry'):
        for blok in entry:

            #  Забираем title каждой записи - это заголовок H1
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

    with open(htmlfilename, 'w', encoding='utf-8') as htmlf:
        htmlf.write(text)

    docxfilename = f'{cur_dir}\\{filename}.docx'
    if os.path.exists(docxfilename):
        os.remove(docxfilename)

    print(f'{Fore.GREEN}\n\nProcess: {htmlfilename} => {docxfilename}')

    new_parser = HtmlToDocx()
    docx = new_parser.parse_html_string(text)
    docx.save(docxfilename)


def main():
    print(f'{cur_dir=}')
    for file in os.listdir(cur_dir):
        if os.path.isfile(file) and file.lower().endswith(".css"):
            processfile(f'{cur_dir}\\{file}', file.split('.')[0])


if __name__ == "__main__":
    cur_dir = os.getcwd()
    print(f"{Fore.CYAN}Last update: Cherepanov Maxim masygreen@gmail.com (c), 12.2021")
    print(f"{Fore.CYAN}Convert file (*.css) to (*.html, *.docx)")
    main()