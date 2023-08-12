from openpyxl.workbook import Workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

#For Mail
import smtplib
from email.message import EmailMessage

driver = webdriver.Chrome()
driver.implicitly_wait(20)
driver.maximize_window()
driver.get("https://www.amazon.in/")
driver.implicitly_wait(20)
driver.find_element(By.ID,"twotabsearchtextbox").send_keys("Iphone")
driver.find_element(By.XPATH,"//input[@type='submit']").click()
#driver.find_element(By.XPATH,"//span[text()='Apple']").click()
phonenames = driver.find_elements(By.XPATH,"//span[@class='a-size-medium a-color-base a-text-normal']")
prices = driver.find_elements(By.XPATH,"//div[@class='a-row a-size-base a-color-base']//span[@class='a-price-whole']")

myphone=[]
myprice=[]

for phone in phonenames:
    print(phone.text)
    myphone.append(phone.text)

print("*" * 50)

for price in prices:
    print(price.text)
    myprice.append(price.text)

print("myphone",myphone)
print("myprice",myprice)

finallist = zip(myphone,myprice)

driver.quit()
print("Part 1")

wb = Workbook()
sh1 = wb.active
wb['Sheet'].title = "Amazon Samsung"

sh1.append(['Name' , 'Price'])
for x in list(finallist):
    sh1.append(x)

wb.save("FinalRecordsNew.xlsx")

print("Part 2")

msg = EmailMessage()
msg['Subject'] = 'Iphone Phone team'
msg['From'] = 'Automation Team'
msg['To'] = 'rahimataur929@gmail.com'

with open('EmailTemplate.text') as myfile:
    data = myfile.read()
    msg.set_content(data)

with open("FinalRecordsNew.xlsx","rb") as f:
    file_data = f.read()
    file_name=f.name
    msg.add_attachment(file_data,maintype="application",subtype="xlsx",filename=file_name)
    print("Attachment attached in message",file_name)


with smtplib.SMTP_SSL('smtp.gmail.com',465) as server:
    print("Hiii")
    server.login("rahimataur929@gmail.com","swdupxbxuhlhgkym")
    server.send_message(msg)

print("Email sent")

