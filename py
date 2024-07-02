import cv2
import pytesseract
import os
import pandas as pd

pytesseract.pytesseract.tesseract_cmd = r"Tesseract-OCR\tesseract.exe"

def parse(x,y,w,h):
    try:
        pay_to = thresh[y:y+h,x:x+w]
        data = pytesseract.image_to_string(pay_to, lang='eng',config='--psm 6')
        data = data.replace("\n\x0c", "")
        return data
    except Exception as e:
        print(e)
        return "#EANF"

def parse2(x,y,w,h):
    try:
        pay_to = thresh[y:y+h,x:x+w]
        data = pytesseract.image_to_string(pay_to, lang='e13b',config='--psm 6')
        data = data.replace("\n\x0c", "")
        return data
    except Exception as e:
        print(e)
        return "#EANF"

final=[]
listfile = os.listdir('images')
for img in listfile:

    if '.png' in img or '.jpg' in img or '.jpeg' in img:
        each={}
        each['images file'] = img
        img = os.path.join('images',img)
        image = cv2.imread(img, 0)

        thresh = 255 - cv2.threshold(image, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
        pay = parse(50,200,600,170)
        pay = pay.replace("\n"," ")
        price = parse(860,150,200,200)
        date = parse(760,90,300,100).replace("Date:","").replace("DATE;","")
        btm_number = parse2(100,420,800,300)
        remiter = parse(15,140,600,80)
        up_number = parse(880,45,200,50)
        print("="*30)
        print("date : ",date)
        print("pay to :", pay)
        print("amount : ", price)
        print("btm :",btm_number)
        print("up number :",up_number)
        print("remiter :", remiter)
        print("="*30)

        each['date'] = date
        each['pay_to'] = pay
        each['amount'] = price
        each['btm_number'] = btm_number
        each['up_number'] = up_number
        each['remiter'] = remiter

        final.append(each)

df=pd.DataFrame(final)
df.to_excel("resultOCR.xlsx",index=False)

    #cv2.waitKey()
