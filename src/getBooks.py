import urllib.request, subprocess, platform, ssl, os
from hurry.filesize import size
from bs4 import BeautifulSoup
import pandas as pd
import sys

# File Data Books
BOOKS = 'dataBooks.xlsx'
COLS = 80
DOWNLOADTITLE = '{:*^'+str(COLS)+'}'
DOWNLOADFORMAT = '{:<17} {:^20} {:^20} {:>20}'
LINES = 50
MESSEGE = 'This script is designed by Giomar Osorio to automatically download the free books offered by Springer Nature, from his web portal.\n\nCheck the main page of Springer Nature for more info:\nhttps://www.springernature.com/gp/librarians/news-events/all-news-articles/industry-news-initiatives/free-access-to-textbooks-for-institutions-affected-by-coronaviru/17855960\n\n'
NAMECATEGORYCOLUMN = 'English Package Name'
SEPARATORLINE = '-'*COLS
SERVER = 'https://link.springer.com'

# Ignore SSL certificate errors
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE


def getData(filename):
    try:
        return pd.read_excel(filename)
    except:
        print('\nplease make sure the "%s" file is in the same folder as the program, then relaunch the program.' % (filename))
        input('press enter to exit...')
        sys.exit()


def getPath(category, foldername):
    parent_dir = os.path.join(os.getcwd(), 'books')
    return os.path.join(parent_dir, os.path.join(category,foldername))


def cleanConsole():
    if platform.system()=="Windows":
        subprocess.Popen("cls", shell=True).communicate() #I like to use this instead of subprocess.call since for multi-word commands you can just type it out, granted this is just cls and subprocess.call should work fine 
    else: #Linux and Mac
        print("\033c", end="")


def printUI():
    print(DOWNLOADTITLE.format('DOWNLOADING BOOKS')+'\n\n')
    print(DOWNLOADFORMAT.format('Book name', 'Category', 'File zise', 'Progress'))
    print(SEPARATORLINE)


def categoriesToPrint(data):
    categories = {}
    categoriesIndex = 0
    for index,row in data:
        if index not in categories.values():
            categories[str(categoriesIndex)] = index
            categoriesIndex+=1
    print('Please select the categories you want to download: (eg. 1,5,7 / blank = all)')
    for key,value in categories.items():
        if int(key)<10:
            print(' %s - %s' % (key, value))
            continue
        print('%s - %s' % (key, value))
    return categories


def categoriesToDownload(data):
    categories = categoriesToPrint(data)
    validCategories = False
    noError = True
    while not validCategories:
        print('', end='\r')
        options = input('Enter categories: ')
        if options!='':
            categoriesToDownloadList = options.split(',')
            for item in categoriesToDownloadList:
                if item.isdigit():
                    if item not in categories.keys():
                        print('The category %s are not valid, please try again.' % (item))
                        noError = False
                        break
                    continue
                else:
                    noError = False
                    print('The options entered are not valid, please try again.')
                    break
            if noError:
                validCategories = not validCategories
        else:
            categoriesToDownloadList = [key for key in categories.keys()]
            validCategories = not validCategories
        noError = True
    print('You have selected "%s" categories, the download will begin shortly.' % (len(categoriesToDownloadList)))
    return [categories[i] for i in categoriesToDownloadList if i in categories.keys()]


def createFolder(category, foldername):
    path = getPath(category, foldername)
    try: 
        os.makedirs(path, exist_ok = True) 
        return True
    except FileExistsError:
        return True
    except OSError as error: 
        return False


def downloadFile(url, path, filename, category):
    if os.path.isfile(os.path.join(path, filename)):
        print(DOWNLOADFORMAT.format(filename[:12]+'...', category[:15]+'...', 'File exist!', 'SKIPPING FILE'))
    else:
        if url != None:
            u = urllib.request.urlopen(SERVER+url)
            f = open(os.path.join(path, filename) , 'wb')
            file_size = int(u.headers['content-length'])
            

            file_size_dl = 0
            block_sz = 8192
            while True:
                buffer = u.read(block_sz)
                if not buffer:
                    break

                file_size_dl += len(buffer)
                f.write(buffer)
                status = r"%3.2f%%" % (file_size_dl * 100. / file_size)
                print(DOWNLOADFORMAT.format(filename[:12]+'...', category[:15]+'...', size(file_size), status), end='\r')

            print('')
            f.close()


def downloadBooks(data,categories):
    dataGroupByCategories = data.groupby(NAMECATEGORYCOLUMN)
    for index,row in dataGroupByCategories:
        if index in categories:
            for index,rowBooks in row.iterrows():
                html = urllib.request.urlopen(rowBooks['OpenURL'], context=ctx).read()
                soup = BeautifulSoup(html, 'html.parser')

                tags = soup.find_all("a", attrs = {'title': 'Download this book in PDF format'})
                if len(tags)>0:
                    if createFolder(rowBooks[NAMECATEGORYCOLUMN], rowBooks['Book Title']):
                        downloadFile(tags[0].get('href', None), getPath(rowBooks[NAMECATEGORYCOLUMN], rowBooks['Book Title'].rstrip()), rowBooks['Book Title'].rstrip()+".pdf", rowBooks[NAMECATEGORYCOLUMN])


def main():
    os.system('mode con: cols=%s lines=%s' % (COLS, LINES))
    print(MESSEGE)
    print('Taking data from Excel file...')
    df = getData(BOOKS)
    print('Extracting the categories from the data...\n')
    todownload = categoriesToDownload(df.groupby(NAMECATEGORYCOLUMN))
    input('press enter to continue...')
    cleanConsole()
    printUI()
    downloadBooks(df, todownload)
    input('all downloaded books, press enter to continue...')


if __name__ == '__main__':
    main()
