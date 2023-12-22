# %%
import datetime
import glob
import os
import pymsteams
import shutil


# %%
#Define các biến ban đầu
yesterday = datetime.date.today()
file_pattern = '*{}_{:02d}_{:02d} 18*.BAK'.format(yesterday.year, yesterday.month, yesterday.day)
bravoBackUpPath = 'D:/BAKUPNEW/'
bravoOneDrivePath = 'C:/Users/Administrator/OneDrive - CONG TY TNHH HTAUTO VIET NAM/BAK_bravo/'

numberOfRequiredFile = 7
listBackUpPath = glob.glob(bravoBackUpPath + file_pattern)
listBackUpFile = [os.path.basename(x) for x in listBackUpPath]

warningWebHookURL = 'https://htautovietnam.webhook.office.com/webhookb2/96106e17-e26a-42cc-a5f2-77a6f3cc98d1@a83657d0-03aa-4e23-b7b8-27d218b177a8/IncomingWebhook/38e1f18bb8aa498487b125150289cd1d/95700413-d3f7-4ae2-a9c5-b2a613f29d96'
successWebHookURL = 'https://htautovietnam.webhook.office.com/webhookb2/96106e17-e26a-42cc-a5f2-77a6f3cc98d1@a83657d0-03aa-4e23-b7b8-27d218b177a8/IncomingWebhook/400ef5db810540cdb3399208c1cb2ecf/95700413-d3f7-4ae2-a9c5-b2a613f29d96'

# %%
#Check xem đã đủ số lượng file chưa
a=0
while len(listBackUpPath) != numberOfRequiredFile:
    listBackUpPath = glob.glob(bravoBackUpPath + file_pattern)
    listBackUpFile = [os.path.basename(x) for x in listBackUpPath]
    if a==0:
        myTeamsMessage = pymsteams.connectorcard(warningWebHookURL)
        myTeamsMessage.text("Số lượng file trong thư mục back up Bravo hiện đang không khớp ghi nhận: {} so với {} file <br><br> Danh sách các file trong thư mục back up hiện có {} <br><br> Yêu cầu bổ sung thêm".format(len(listBackUpPath), numberOfRequiredFile, str(listBackUpFile)))
        myTeamsMessage.send()
        a = a + 1


if len(listBackUpPath) == numberOfRequiredFile:
    for i in range(0, numberOfRequiredFile):
        shutil.copy(listBackUpPath[i], bravoOneDrivePath + listBackUpFile[i])


myTeamsMessage = pymsteams.connectorcard(successWebHookURL)
myTeamsMessage.text("Đã đẩy file back up lên onedrive. Tổng có {} file <br><br>{}".format(len(listBackUpFile),str(listBackUpFile)))
myTeamsMessage.send()






