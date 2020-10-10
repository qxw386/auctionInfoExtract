import cv2
import re, os, sys
import xlsxwriter
import pytesseract
import pandas as pd
import numpy as np
from pytesseract import Output
import configparser
import multiprocessing as mp
from tqdm import tqdm

isDebug = False



def grabtextfromImages(imageDiagnose, dateUploaded, isDebug, ctrlDist):
    image = cv2.imread(imageDiagnose)
    img = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)


    # Trace text character 
    # h, w, c = img.shape
    # boxes = pytesseract.image_to_boxes(img) 
    # for b in boxes.splitlines():
    #     b = b.split(' ')
    #     img = cv2.rectangle(img, (int(b[1]), h - int(b[2])), (int(b[3]), h - int(b[4])), (0, 255, 0), 2)

    # cv2.imshow('img', img)
    # cv2.waitKey(0)


    # # Trace text word
    # d = pytesseract.image_to_data(img, output_type=Output.DICT)
    # # print(d.keys())
    # # n_boxes = len(d['text'])
    # n_boxes = len(d['level'])
    # for i in range(n_boxes):
    #     if int(d['conf'][i]) > 0:
    #         (x, y, w, h) = (d['left'][i], d['top'][i], d['width'][i], d['height'][i])
    #         img = cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 2)
    # cv2.imshow('img', img)
    # cv2.waitKey(0)

    # Trace text regex
    predictionTextphase = 1
    bidPriceBottom = 0
    temp_area = ""
    temp_addr = ""
    temp_type = ""
    temp_bidprice = ""
    temp_marketValue = ""
    temp_propertyID = ""
    temp_auctionDate = ""
    temp_tenure = ""
    temp_restriction = ""
    temp_landArea = ""
    gotProperty = False
    gotAuctionDate = False
    gotTenurePD = False
    gotRestrictiony = False
    gotLandArea = False
    marketValueIndexDel = []

    d = pytesseract.image_to_data(img, lang='eng', output_type='data.frame')
    d = d[d['conf'] > 50]
    d = d.dropna()
    d = d[(~d.text.str.contains(":")) & (~d.text.str.contains(" ")) ]
    d = d.reset_index(drop=True)

    if isDebug.lower() == 'true':
        for i in range(len(d['text'])):
            (x, y, w, h) = (d['left'][i], d['top'][i], d['width'][i], d['height'][i])
            img = cv2.rectangle(img, (x, y), (x + w, y + h), (0, 255, 0), 2)

    # First we isolate and remove MarketValue data
    marketValuePD = d[d['left'] > 290]
    for index, row in marketValuePD.iterrows():
        if (re.match('[rR][mM].[0-9]{1,}[kK]|[rR][mM][0-9]{1,}\,[0-9]{1,}|[rR][mM][0-9]{1,}[kK]|[rR][mM][0-9]{1,}|MV', row['text'])):
            marketValueIndexDel.append(index)
        if (re.match('[rR][mM].[0-9]{1,}[kK]|[rR][mM][0-9]{1,}\,[0-9]{1,}|[rR][mM][0-9]{1,}[kK]|[rR][mM][0-9]{1,}', row['text'])):
            if len(temp_marketValue) > 0 :
                temp_marketValue = temp_marketValue + " " + row['text']
            else:
                temp_marketValue =  row['text']
    if isDebug.lower() == 'true':
        print(marketValuePD)
        print("marketValueIndexDel:", marketValueIndexDel)
    d = d.drop(marketValueIndexDel)
    d = d.reset_index(drop=True)

    d['bottom'] = d['top'] + d['height']
    for i in range(1, len(d)):
        if i == 1:
            d['distance'] = 0
        else:
            d.loc[i, 'distance'] = 0 if d.loc[i, 'top'] - d.loc[i-1, 'top'] < 5 else d.loc[i, 'top'] - d.loc[i-1, 'bottom']
    
    for j in range(0, len(d)):
        if d.loc[j, 'distance'] < int(ctrlDist) \
            and ( d.loc[j, 'text'] != 'No.' or int(predictionTextphase) != 1 )\
            and ( d.loc[j, 'text'] != 'Double'  or int(predictionTextphase) != 2)\
            and ( d.loc[j, 'text'] != 'Single'  or int(predictionTextphase) != 2)\
            and ( d.loc[j, 'text'] != 'Triple'  or int(predictionTextphase) != 2)\
            and ( d.loc[j, 'text'] != 'Flat'  or int(predictionTextphase) != 2)\
            and ( d.loc[j, 'text'] != 'Service'  or int(predictionTextphase) != 2)\
            and ( d.loc[j, 'text'] != 'Terrace'  or int(predictionTextphase) != 2)\
            :
            d.loc[j, 'paragraph'] = int(predictionTextphase)
        elif d.loc[j, 'distance'] > int(ctrlDist) \
            and predictionTextphase == 1  and d.loc[j, 'block_num'] == d.loc[j-1, 'block_num'] and abs(int(d.loc[j, 'distance']) - int(ctrlDist)) <= 6 :
            d.loc[j, 'paragraph'] = int(predictionTextphase)
        else:
            predictionTextphase = predictionTextphase + 1
            d.loc[j, 'paragraph'] = int(predictionTextphase)

    reodercolumns = ['level', 'page_num', 'block_num', 'par_num', 'line_num', 'word_num',\
                     'top', 'height', 'bottom', 'distance' , 'paragraph', 'left', 'width', 'conf', 'text']
    d = d[reodercolumns]
    d = d.astype({"paragraph":'int64'})

    for paragraphIndex in range(d['paragraph'].max()):
        if paragraphIndex == 0:
            areaPD= d[d['paragraph']==paragraphIndex+1]
            for index, row in areaPD.iterrows():
                if len(temp_area) > 0 :
                    temp_area = temp_area + " " + row['text']
                else:
                    temp_area =  row['text']
        elif paragraphIndex == 1:
            addressPD= d[d['paragraph']==paragraphIndex+1]
            for index, row in addressPD.iterrows():
                if len(temp_addr) > 0 :
                    temp_addr = temp_addr + " " + row['text']
                else:
                    temp_addr =  row['text']
        elif paragraphIndex == 2:
            typePD= d[d['paragraph']==paragraphIndex+1]
            for index, row in typePD.iterrows():
                if len(temp_type) > 0 :
                    temp_type = temp_type + " " + row['text']
                else:
                    temp_type =  row['text']
        elif paragraphIndex == 3:
            bidpricePD= d[d['paragraph']==paragraphIndex+1]
            for index, row in bidpricePD.iterrows():
                if (re.match('[rR][mM] {0,}[0-9]{1,},[0-9]{1,}.[0-9]{1,}|[rR][mM]$|[rR][mM] $|[0-9]{1,},[0-9]{1,}.[0-9]{1,}', row['text']) ):
                    if len(temp_bidprice) > 0 :
                        temp_bidprice = temp_bidprice + " " + row['text']
                        bidPriceBottom = row['bottom']
                    else:
                        temp_bidprice =  row['text']
                        bidPriceBottom = row['bottom']
        elif paragraphIndex > 3:
            if bidPriceBottom > 0 :
                onlyDetaildataBottome = bidPriceBottom
            else:
                onlyDetaildataBottome = d['bottom']

            leftDetailData = d[(d['top'] > onlyDetaildataBottome)  & (d['left'] < 219) & (d['paragraph']==paragraphIndex+1)]

            if not leftDetailData[leftDetailData['text'].str.lower() == 'property'].empty:
                gotProperty = True
            if not leftDetailData[leftDetailData['text'].str.lower() == 'auction'].empty:
                gotAuctionDate = True
            if not leftDetailData[leftDetailData['text'].str.lower() == 'tenure'].empty:
                gotTenurePD = True
            if not leftDetailData[leftDetailData['text'].str.lower() == 'restriction'].empty:
                gotRestrictiony = True
            if not leftDetailData[leftDetailData['text'].str.lower() == 'land'].empty:
                gotLandArea = True
            if not leftDetailData[leftDetailData['text'].str.lower() == 'built'].empty:
                gotLandArea = True
    
  
    rightDetailData= d[(d['top'] > onlyDetaildataBottome)  & (d['left'] >= 219)]
    if isDebug.lower() == 'true':
        print(gotProperty,gotAuctionDate,gotTenurePD,gotRestrictiony,gotLandArea)
        print(rightDetailData)

    if gotProperty:
        for index, row in rightDetailData.iterrows():
            if (re.match('^[0-9][0-9]{4,}', row['text']) ):
                if len(temp_propertyID) > 0 :
                    temp_propertyID = temp_propertyID + " " + row['text']
                else:
                    temp_propertyID =  row['text']
    if gotAuctionDate:
        for index, row in rightDetailData.iterrows():
            if (re.match('[0-9][0-9]-[aA-zZ][aA-zZ][aA-zZ]-[0-9][0-9]|\([aA-zZ][aA-zZ][aA-zZ]\)', row['text']) ):
                if len(temp_auctionDate) > 0 :
                    temp_auctionDate = temp_auctionDate + " " + row['text']
                else:
                    temp_auctionDate =  row['text']
    if gotTenurePD:
        for index, row in rightDetailData.iterrows():
            if (re.match('[lL][eE][aA][sS][eE][hH][oO][lL][dD]|[fF][rR][eE][eE][hH][oO][lL][dD]', row['text']) ):
                if not (re.match('Tenure', row['text'])):
                    if len(temp_tenure) > 0 :
                        temp_tenure = temp_tenure + " " + row['text']
                    else:
                        temp_tenure =  row['text']
    if gotRestrictiony:
        for index, row in rightDetailData.iterrows():
            if (re.match('[aA-zZ]{3,}', row['text']) ):
                if not (re.match('[lL][eE][aA][sS][eE][hH][oO][lL][dD]|[fF][rR][eE][eE][hH][oO][lL][dD]', row['text'])):
                    if len(temp_restriction) > 0 :
                        temp_restriction = temp_restriction + " " + row['text']
                    else:
                     temp_restriction =  row['text']
    if gotLandArea:
        for index, row in rightDetailData.iterrows():
            if ( re.match('[0-9]\,[0-9]{1,}|[0-9][0-9][0-9]$|[0-9][0-9]\,[0-9]{1,}|[0-9][0-9][0-9]$', row['text']) \
                    or re.match('[sS][qQ]\.[fF][tT]|\:.[sS][qQ]\.[fF][tT]', row['text']) ):
                if len(temp_landArea) > 0 :
                    temp_landArea = temp_landArea + " " + row['text']
                else:
                    temp_landArea =  row['text']

    Area = temp_area
    Address = temp_addr
    housetype = temp_type
    marketValue = temp_marketValue
    bidprice = temp_bidprice
    propertyID = temp_propertyID
    auctionDate = temp_auctionDate
    tenure = temp_tenure.replace('^ ', '')
    restriction = temp_restriction
    landArea = temp_landArea
    LinkToImage = '=HYPERLINK("' + imageDiagnose + '")'
    checkState = Address.lower().replace(' ','')
    if checkState.__contains__('kualalumpur'): state="Kuala Lumpur"
    elif checkState.__contains__('selangor'): state="Selangor"
    elif checkState.__contains__('negerisembilan'): state="Negeri Sembilan"
    elif checkState.__contains__('johor'): state="Johor"
    elif checkState.__contains__('kelantan'): state="Kelatan"
    elif checkState.__contains__('pahang'): state="Pahang"
    elif checkState.__contains__('terengganu'): state="Terengganu"
    elif checkState.__contains__('perlis'): state="Perlis"
    elif checkState.__contains__('penang'): state="Penang"
    elif checkState.__contains__('melaka'): state="Melaka"
    elif checkState.__contains__('perak'): state="Perak"
    elif checkState.__contains__('sarawak'): state="Sarawak"
    elif checkState.__contains__('sabah'): state="Sabah"
    elif checkState.__contains__('kedah'): state="Kedah"
    else: state="Unknown"
    if isDebug.lower() == 'true':
        print(d)
        print("Area: " + Area)
        print("Address: " + Address)
        print("State: " + state)
        print("House Type: " + housetype)
        print("Bidprice: " + bidprice)
        print("Market Value: " + marketValue)
        print("Property ID: " + propertyID)
        print("Auction Date: " + auctionDate)
        print("Tenure: " + tenure)
        print("Restriction:" + restriction)
        print("LandArea: " + landArea)
        cv2.imshow('img', img)
        cv2.waitKey(0)

    listResult = [dateUploaded,Area,Address,state,housetype,bidprice,marketValue,propertyID,auctionDate,tenure,restriction,landArea,LinkToImage]
    return listResult

def write2Excel(myWorkbook,auctionlistHeader, auctionlisttBody, sheetname):
    worksheet = myWorkbook.add_worksheet(sheetname)
    col_num = 0
    for tempHead in auctionlistHeader:
     worksheet.write(0, col_num, tempHead)
     col_num += 1

    row_num = 1
    for i in range(len(auctionlisttBody)):
     col_num = 0
     for j in range(len(auctionlisttBody[i])):
      worksheet.write(row_num, col_num, auctionlisttBody[i][j])
      col_num += 1
     row_num += 1

if __name__ == '__main__':
    mp.freeze_support()
    images2process = []
    auctionlist = []
    barCounter = 0

    config = configparser.ConfigParser()
    if not os.path.isfile("auctionImageCapture.properties"):
        print("Error: Config File auctionImageCapture.properties is not found !!")
        sys.exit()
    config.read('auctionImageCapture.properties')
    InputPath = config['auctionConf']['InputPath']
    isDebug = config['auctionConf']['isDebug']
    ctrlDist = config['auctionConf']['distance']
    FileBassedParallelizing = config['MIDPParallelizing']['FileBassedParallelizing']
    FileBassedThreads = config['MIDPParallelizing']['FileBassedThreads']

    for f in os.listdir(InputPath):
        if os.path.isdir(InputPath + f):
            for g in os.listdir(InputPath + f):
                if g.endswith('.jpg'):
                    images2process.append([f,g])
                        # auctionlist.append(grabtextfromImages(InputPath + f + '/'+ g , f))
    # for a,b in images2process:
    #     print(a,b)

    

    if FileBassedParallelizing.lower() == 'true':
        print("Info: Parallelize Mode!")
        if int(FileBassedThreads) > mp.cpu_count():
            print("Error: FileBassedThreads is set more than CPUs on this machine!" + "[" + str(mp.cpu_count()) + "]")
            sys.exit()
        elif int(FileBassedThreads) == 0:
            if len(images2process) < mp.cpu_count():
                pool = mp.Pool(len(images2process))
            else:
                pool = mp.Pool(mp.cpu_count())
        else:
            pool = mp.Pool(int(FileBassedThreads))
        try:
            print("Proceeding..............................")
            result = (pool.starmap(grabtextfromImages, [ (InputPath + dateUploaded + "/" + fileCSV,dateUploaded, isDebug, ctrlDist) for dateUploaded,fileCSV in images2process] ))
            pool.close()
            if len(auctionlist) == 0:
                auctionlist = result
            else:
                auctionlist.append(result)
        except Exception as e:
         print("Error: Fail to parse with error below")
         print(e)
         pool.terminate()
         sys.exit()
    else:
        print("Info: serialize Mode!")
        print("Proceeding..............................")
        for dateUploaded,fileCSV in images2process:
            result = grabtextfromImages(InputPath + dateUploaded + "/" + fileCSV,dateUploaded, isDebug, ctrlDist)
            auctionlist.append(result)

    if isDebug.lower() == 'false':
        # print(auctionlist)
        auctionlistHeader=["Date Uploaded","Area","Address","State","House Type","Bidprice","Market Value","Property ID","Auction Date","Tenure","Restriction","LandArea","LinkToImage"]
        workbook = xlsxwriter.Workbook( "auctionList.xlsx")
        write2Excel(workbook,auctionlistHeader, auctionlist, "auctionList")
        workbook.close()