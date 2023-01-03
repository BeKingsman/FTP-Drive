from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import os
import requests
import json
import xlsxwriter

gauth = GoogleAuth()
drive = GoogleDrive(gauth)

folder_id = "1sYQcl2wWI2AWpxJcAgeePA_9UabuXuF9"
path = str(input("Enter Image Folder Path: "))
all_links = []
image_per_product = int(input("Number of Images Per Row: "))


def ListFolder(parent):
    filelist = []
    file_list = drive.ListFile(
        {'q': "'%s' in parents and trashed=false" % parent}).GetList()
    for f in file_list:
        if f['mimeType'] == 'application/vnd.google-apps.folder':  # if folder
            filelist.append(
                {"id": f['id'], "title": f['title'], "list": ListFolder(f['id'])})
        else:
            filelist.append(
                {"title": f['title'], "link": f['alternateLink']})
    return filelist


def get_file_link(gfile):
    access_token = gauth.credentials.access_token
    file_id = gfile['id']
    url = 'https://www.googleapis.com/drive/v3/files/' + \
        file_id + '/permissions?supportsAllDrives=true'
    headers = {'Authorization': 'Bearer ' +
               access_token, 'Content-Type': 'application/json'}
    payload = {'type': 'anyone', 'value': 'anyone', 'role': 'reader'}
    res = requests.post(url, data=json.dumps(payload), headers=headers)

    link = gfile['alternateLink']
    return link


def upload_images():
    images = os.listdir(path)
    for img in images:
        if img.endswith(".jpg") or img.endswith(".jpeg") or img.endswith(".JPG") or img.endswith(".JPEG") or img.endswith(".png") or img.endswith(".PNG"):
            try:
                gfile = drive.CreateFile(
                    {'parents': [{'id': folder_id}]})

                gfile.SetContentFile(img)
                gfile.Upload()
                link = get_file_link(gfile)
                all_links.append(link)
            except Exception as e:
                print(str(e))
                all_links.append("-")


def output_links():
    workbook = xlsxwriter.Workbook('image_links.xlsx')
    worksheet = workbook.add_worksheet()

    for i in range(len(all_links)):
        row = int(i/image_per_product)
        col = int(i % image_per_product)
        worksheet.write(row, col, all_links[i])
        print(all_links[i])
    workbook.close()


upload_images()
output_links()
