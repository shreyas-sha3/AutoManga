import os
import time
import shutil
import smtplib
import requests
import pyautogui
import subprocess
from pathlib import Path
from bs4 import BeautifulSoup
from optparse import OptionParser
from importlib import import_module
from email.message import EmailMessage
from epubMaker import EPubMaker, CmdProgress

#Variables
mangaName = 'null'
chNo = '0'
to= open("Config/ReceiverList.ini","r").read()
Resolution = open("Config/Resolution.ini", "r")
Width = int(Resolution.readline().split()[2])
Height = int(Resolution.readline().split()[2])

#Making new Directory
try:
    os.mkdir('Downloaded Mangas')
except OSError:
    None

#Loop for different mangas
fav = open("Config/MangaList.ini","r").read()
numOfMangas= len(fav.split(","))
for i in range(0, numOfMangas):
    mangaName=(fav.split(",")[i])
    i = 1
    
    #Gettting Latest Chapter
    URL = 'https://w13.mangafreak.net/Manga/'+ mangaName
    print(URL)
    r = requests.get(URL)
    soup = BeautifulSoup(r.content, 'lxml') 
    recent_chaps = soup.find('div', class_='series_sub_chapter_list')
    latest_chap =  recent_chaps.a.text.split()[1]
    chNo = latest_chap
    prev_chap = recent_chaps.find_all('a')[1].text.split()[1]
    newURL = 'https://w13.mangafreak.net' + recent_chaps.a['href']
    mangaDir = 'Downloaded Mangas/' + mangaName + '_'+ chNo + '/'

    #Check if new Manga is released
    try:
        os.mkdir(mangaDir)
    except:
        print("New chapter is not released for " + mangaName)
        continue

    #Downloading Individual Pages from manga
    rs = requests.get(newURL)
    soup = BeautifulSoup(rs.content, 'lxml') 
    for item in soup.find_all('img',id='gohere'):
        print("Downloading Page: "+ str(i))
        imgUrl = requests.get(item['src'])
        open(mangaDir+ str(i) +".jpg","wb").write(imgUrl.content)
        i = i + 1
    
    #Conversion to EPUB
    if __name__ == '__main__':
        parser = OptionParser(
            usage='usage: %prog [--cmd] [--progress] --dir DIRECTORY --file FILE --name NAME\n'
              '   or: %prog [--progress] DIRECTORY DIRECTORY ... (batchmode, implies -c)'
    )
    parser.add_option(
        '-c', '--cmd', action='store_true', dest='cmd', default=True, help='Start without gui'
    )
    parser.add_option(
        '-p', '--progress', action='store_true', dest='progress', default=True,
        help='Show a nice progressbar (cmd only)'
    )
    parser.add_option(
        '-d', '--dir', dest='input_dir', metavar='DIRECTORY', default=mangaDir ,help='DIRECTORY with the images'
    )
    parser.add_option(
        '-f', '--file', dest='file', metavar='FILE', default= "Downloaded Mangas/" + mangaName + '_'+ chNo + '.epub',help='FILE where the ePub is stored'
    )
    parser.add_option(
        '-n', '--name', dest='name', metavar='NAME',default= mangaName + '_'+ chNo , help='NAME of the book'
    )
    parser.add_option(
        '-g', '--grayscale', dest='grayscale', default=False, action='store_true',
        help="Convert all images to black and white before adding them to the ePub.",
    )
    parser.add_option(
        '-W', '--max-width', dest='max_width', default=Width, type="int",
        help="Resize all images to have the given maximum width in pixels."
    )
    parser.add_option(
        '-H', '--max-height', dest='max_height', default=Height, type="int",
        help="Resize all images to have the given maximum height in pixels."
    )
    parser.add_option(
        '--wrap-pages', dest='wrap_pages', action='store_true',
        help="Wrap the pages in a separate file. Results will vary for each reader. (Default)"
    )
    parser.add_option(
        '--no-wrap-pages', dest='no_wrap_pages', action='store_true',
        help="Do not wrap the pages in a separate file. Results will vary for each reader."
    )
    (options, args) = parser.parse_args()

    if options.wrap_pages and options.no_wrap_pages:
        parser.error("options --wrap-pages and --no-wrap-pages are mutually exclusive")

    if not options.input_dir and not options.file and not options.name:
        if not all(os.path.isdir(elem) for elem in args):
            parser.error("Not all given arguments are directories!")

        directories = []
        for elem in args:
            path = Path(args)
            if not path.is_dir():
                parser.error(f"The following path is not a directory: {path}")
            if not path.name:
                parser.error(f"Could not get the name of the directory: {path}")
            directories.append(path)

        for path in directories:
            EPubMaker(
                master=None, input_dir=path, file=path.parent.joinpath(path.name + '.epub'), name=path.name or "Output",
                grayscale=options.grayscale, max_width=options.max_width, max_height=options.max_height,
                progress=CmdProgress(options.progress), wrap_pages=not options.no_wrap_pages
            ).run()
    elif options.input_dir and options.file and options.name:
        if options.cmd:
            if args or not options.input_dir or not options.file or not options.name:
                parser.error("The '--dir', '--file', and '--name' arguments are required.")

            EPubMaker(
                master=None, input_dir=options.input_dir, file=options.file, name=options.name,
                grayscale=options.grayscale, max_width=options.max_width,
                max_height=options.max_height, progress=CmdProgress(options.progress),
                wrap_pages=not options.no_wrap_pages
            ).run()
    

    #Sending Mail
    print("Sending Mail...")
    msg=EmailMessage()
    msg['Subject']=mangaName + ' Chapter ' + chNo
    msg['From']='Auto Manga'
    msg['To']= to

    with open("Downloaded Mangas/"+ mangaName+ "_" +chNo+ ".epub","rb") as manga:
        file_data=manga.read()
        file_name=mangaName + ' Chapter ' + chNo + '.epub'
        msg.add_attachment(file_data,maintype="application",subtype="epub",filename=file_name)

    server =smtplib.SMTP_SSL("smtp.gmail.com", 465)
    server.login("weeklyautomanga@gmail.com", "testaccpass123")
    server.send_message(msg)
    print(mangaName+" Manga Sent")

    #Removing previous chapters
    shutil.rmtree("Downloaded Mangas/"+ mangaName + '_' + chNo)
    os.remove("Downloaded Mangas/"+ mangaName + '_' + chNo + '.epub')
    try:
        shutil.rmtree("Downloaded Mangas/"+ mangaName + '_' + chNo)
    except:
        None

print("ALL MANGAS SENT")
time.sleep(10)
















