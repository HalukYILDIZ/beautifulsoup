from xlwt import Workbook,Formula,easyxf

import requests
import datetime
from bs4 import BeautifulSoup
wb=Workbook()
sheet1=wb.add_sheet('sheet 1')
style1=easyxf('pattern:pattern solid, fore_colour yellow;'
			      'font: colour black, bold True;')
style2=easyxf('pattern: pattern solid, fore_colour light_green;'
                              'font: colour black, bold False;')
for i in range(1,31):
    sheet1.col(i).width=7000
sheet1.write(0,0,"index",style1)
sheet1.write(0,1,"Baslik",style1)
sheet1.write(0,2,"Internet adresi",style1)
sheet1.write(0,3,"Sahibi",style1)
sheet1.write(0,4,"Telefonu",style1)
sheet1.write(0,5,"Fiyati",style1)
sheet1.write(0,6,"Ev adresi",style1)
toplamUrun=0
ozellik_no=0

for sayfaNo in range(1,40):
    url = "https://www.hurriyetemlak.com/ankara-satilik-sahibinden/isyeri?utm_expid=.NIZPTyhmTEqACLpdTPxKHw.0&utm_referrer=https%3A%2F%2Fwww.hurriyetemlak.com%2Fsatilik-sahibinden%2Fisyeri&page="+str(sayfaNo)
    r= requests.get(url)
    soup= BeautifulSoup(r.content,"lxml")
    villalar=soup.find_all("div",attrs={"class":"list-item timeshare clearfix"})

    for villa in villalar:
        urunAdi=villa.a.get("title")
        urunLink=villa.a.get("href")
        urun_url="https://www.hurriyetemlak.com"+urunLink

       # print(urunAdi)
      #  print(urun_url)
        try:
            urun_r = requests.get(urun_url)
            toplamUrun +=1
            sheet1.write(toplamUrun,1,urunAdi,style2)
            sheet1.write(toplamUrun,2,urun_url)
        except Exception:
            print ("Ürün bulunamadı.")
        try:
            urun_soup=BeautifulSoup(urun_r.content,"lxml")
            owner=urun_soup.find("div",attrs={"class":"owner-name jsDataOwnerName"}).text
            price=urun_soup.find("span").text
            tel=urun_soup.find("ul",attrs={"class":"phone-numbers"}).a.get("href")
            imgs=urun_soup.find_all("figure",limit=5)

        #print(tel)
            sheet1.write(toplamUrun, 4, tel)
        #print(price)
            sheet1.write(toplamUrun, 5, price,style2)
        #print(owner)
            sheet1.write(toplamUrun, 3, owner,style2)
            try:
                adress=urun_soup.find("span",attrs={"class":"address-line-breadcrumb"}).text
            except Exception as e:
                adress="bulunamadı"
            sheet1.write(toplamUrun, 6, adress)
            ozellikler=urun_soup.find("li",attrs={"class":"info-line"})
            try:
                o_liler=ozellikler.find_all("li",limit=25)
            except Exception as e:
                print(e)
                print(urunLink)
                print (urunAdi)
            o=7
            for o_li in o_liler:
                sheet1.write(toplamUrun, o, o_li.text)
                o+=1
            image_no=35
            for img in imgs:
                img_url=img.img.get("src")
                sheet1.write(toplamUrun,image_no,img_url)
                print(img_url)
                image_no+=1
        except Exception as e:
            print(e)
      #  ozellik_no=4

             #   sheet1.write(toplamUrun,ozellik_no,urun_label+":"+urun_data)
            #    ozellik_no +=1
       # print("#"*60)
print ("Toplam {} kadar ürün bulundu.".format(toplamUrun))
for index in range(1,toplamUrun+1):
    sheet1.write(index,0,index)
an = datetime.datetime.now()
dt=datetime.datetime.strftime(an, '%d %m %Y')
print(dt)
wb.save('hurriyetsahibindenankaraisyerleri'+dt+'.xls')