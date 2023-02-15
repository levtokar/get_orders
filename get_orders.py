"""
This is a Python script that uses the Selenium package to get data from different websites. The script is divided into several functions that perform different tasks, including logging in to the website, navigating to different pages, and copying data to the clipboard.

The selenium.webdriver module is used to interact with the web browser, and the webdriver_manager.chrome module is used to automatically download and install the Chrome driver. The pynput.keyboard module is used to simulate keyboard events, and the pyautogui module is used to take screenshots.

The script defines a class called fumes_class, which represents a single row of data that will be extracted from the website. The script also defines a list called kaspi_ids, which will hold the IDs of the orders that have been scraped.

The script uses several external libraries, including ctypes, os.path, subprocess, and time.
"""

# Import necessary modules
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from inspect import getsourcefile
from os.path import abspath
import subprocess
from pynput.keyboard import Key, Controller
from threading import Timer
import sys
import os.path
import time
import pyautogui
from os.path import isfile, join
import ctypes 
import time

# Create a keyboard object to simulate keyboard input
keyboard = Controller()


# Define function to print the full exception and exit the program
def show_exception_and_exit(exc_type, exc_value, tb):
    import traceback
    traceback.print_exception(exc_type, exc_value, tb)
    sys.exit(-1)

sys.excepthook = show_exception_and_exit

# Define class to store information for each fumes order
class fumes_class:
    id2: str
    name: str
    surname: str
    order_kaspi: str
    order_insales: str
    adress: str
    phone: str
    title: str
    sku: str
    quantity: str
    price: str
    date: str
    delivery: str
    sign:str

# Initialize kaspi_ids and fumes arrays
kaspi_ids = [""] * 999
fumes = [fumes_class()] * 999

# Initialize max_id and cur_id to keep track of which kaspi_ids have been processed
max_id=0
cur_id=0

# Define delay functions to wait for page to load
def delay2(dwMilliseconds,text):
  iStart = time.perf_counter()
  while 1==1:
    s=driver.page_source
    iStop = time.perf_counter()
    if ((iStop - iStart) >= dwMilliseconds) or (text in s):
       break

def delay(dwMilliseconds):
  iStart = time.perf_counter()
  while 1==1:
    iStop = time.perf_counter()
    if ((iStop - iStart) >= dwMilliseconds):
       break  

# Define function to log in to kaspi.kz
def login_kaspi():
    driver.get("https://kaspi.kz/mc/#/login")

    time.sleep(3) 

    driver.execute_script("document.getElementsByClassName('is-active').item(0).nextSibling.firstChild.childNodes[2].click()")

    time.sleep(1) 
    driver.find_element(By.ID, "user_email").send_keys('ikkert.p@gmail.com')
    
    driver.execute_script("document.getElementsByClassName('button is-primary').item(0).click()")
    
    time.sleep(1) 
    driver.find_elements(By.CLASS_NAME , "text-field")[2].send_keys('password') 

    time.sleep(1)  
    driver.execute_script("document.getElementsByClassName('button is-primary').item(0).click()")
    time.sleep(1)  




# function to get kaspi IDs
def get_ids_kaspi():
    # wait for 3 seconds
    time.sleep(3) 

    # get the page source
    s = driver.page_source
    
    # declare max_id as a global variable
    global max_id
    
    # loop until 'class="rows">' is no longer found in s
    while 'class="rows">' in s: 
        # update s to the next substring after 'class="rows">'
        s = s[s.index('class="rows">') + 8:]
                
        # get the substring after 'href="#/' and before the next '/'
        s4 = s[s.index('href="#/') + 11:]
        s4 = s4[s4.index('/') + 1:]
        s4 = s4[:s4.index('"')]
        
        # print max_id and s4
        print(max_id)
        print(s4)
                
        # add the kaspi ID and "Нет" (meaning "no" in Russian) to kaspi_ids dictionary using max_id as the key
        kaspi_ids[max_id] = s4 + '$#$' + 'Нет'

        # increment max_id
        max_id += 1

# function to add a kaspi ID manually
def get_ids_kaspi2():
    # declare max_id as a global variable
    global max_id
    
    # manually set s4 to '182990380'
    s4 = '182990380'
    
    # add the kaspi ID and "Нет" to kaspi_ids dictionary using max_id as the key
    kaspi_ids[max_id] = s4 + '$#$' + 'Нет'
    
    # increment max_id
    max_id += 1 

# function to get the last page
def get_lastpage():
    # wait for 3 seconds
    time.sleep(3)
    
    # get the page source
    s = driver.page_source

    # if 'pagination-previous' is found in s, update s to the substring after 'pagination-previous'
    if 'pagination-previous' in s: 
        s = s[s.index('pagination-previous'):]
    else:
        # if 'pagination-previous' is not found, set s to '123'
        s = '123';     
    
    # if s is '123', return False
    if s == "123":
        return False
    # if 'disabled="disabled' is found in s, return False
    elif 'disabled="disabled' in s:
        return False
    else: 
        # if neither of the above conditions is true, return True
        return True

# function to check if an element exists by xpath
def check_exists_by_xpath(xpath):
    try:
        driver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True


# define a function called next_kaspi_id_write
def next_kaspi_id_write():
    # sleep for 5 seconds
    time.sleep(5) 
    
    # use the global variable cur_id to get the current kaspi_id
    global cur_id
    s=kaspi_ids[cur_id]

    # extract the order ID from the kaspi_id
    s3=s
    s=s[:s.index('$#$')]
    s=s[:s.index('?')]

    # increment cur_id
    cur_id=cur_id+1
    
    # if the condition is True
    if 1==1:

        # set a flag to True
        flag1=True

        # create three directories if they don't already exist
        global dir_path
        if not os.path.exists(dir_path+"\kaspi_ids"):
           os.mkdir(dir_path+"\kaspi_ids")
    
        if not os.path.exists(dir_path+"\kaspi_ids2"):
           os.mkdir(dir_path+"\kaspi_ids2")       
        
        if not os.path.exists(dir_path+"\kaspi_ids3"):
           os.mkdir(dir_path+"\kaspi_ids3")  

        # read the contents of the ll_olds.txt file
        f4 = open(dir_path+"\ll_olds.txt", "r", encoding='utf-8')            
        s_olds=f4.read() 
        f4.close()

        # if the order file exists and the order is not in the ll_olds.txt file
        if isfile(dir_path+"\kaspi_ids2\\" + str(s) + ".txt") and not(s  in s_olds):
            # if the order file has content
            if (os.stat(dir_path+"\kaspi_ids2\\" + str(s) + ".txt").st_size>0):
                # log in to the kaspi.kz website
                driver.get('https://kaspi.kz/mc/#/orders/' + s)
                delay(5)
                s=driver.page_source 
                # if the order can be accepted
                if "<span> Изменить состав </span>" in s:
                    driver.execute_script('document.getElementsByClassName("button").item(3).click()') 
                else:
                    driver.execute_script('document.getElementsByClassName("button").item(2).click()') 
                # write the contents of s to ll_olds.txt
                f4 = open(dir_path+"\ll_olds.txt", "a", encoding='utf-8')            
                f4.write(s+"\n") 
                f4.close()
                flag1=False
        # if the order file exists and the order is in the ll_olds.txt file
        elif isfile(dir_path+"\kaspi_ids2\\" + str(s) + ".txt"):
            flag1=False

        

        

        if flag1:
            # Go to the web page and extract the data
            driver.get('https://kaspi.kz/mc/#/orders/' + s)
            s4 = s3
            delay(5)
            delay2(20,"Название в Kaspi Магазине")
            s = driver.page_source
            print(s)
            
           


            # Check the data for the desired values
            if (s4 in s) and ('Отменен покупателем' not in s):
              # Write the data to files
              
              
            
              

              f3 = open(dir_path + "\kaspi_ids3\\" + str(s4) + ".txt", "w", encoding='utf-8')         
              #f3.write(s+"\n")    #uncomment for debugging
              f3.close()
              f1 = open(dir_path + "\kaspi_ids\\" + str(s4) + ".txt", "w", encoding='utf-8') 
              f2 = open(dir_path + "\kaspi_ids2\\" + str(s4) + ".txt", "w", encoding='utf-8')         
      
              stemp=s      
              s=stemp
              s=s[s.index('>Имя')+2:]
              s=s[s.index('ck">')+4:]
              s=s[:s.index('<')]
              name2=s
              s=s[:s.index(' ')]
              s=s.replace("&nbsp;", "") 
              f1.write(s+"\n") 
              f2.write(s+"\n") 
        

              s=name2
              s=s[s.index(' ')+1:]
              s=s.replace("&nbsp;", "") 
              f1.write(s+"\n") 
              f2.write(s+"\n") 

              
              s=stemp         
              s=s[s.index('Адрес доставки')+2:]
              s=s[s.index('="">')+4:]
              s=s[:s.index('<')]
              s=s.replace("&nbsp;", "")   
              f1.write(s+"\n") 
              f2.write(s+"\n") 
        

              s=stemp
              s=s[s.index('Номер телефона')+2:]  
              s=s[s.index('tel:')+4:]
              s='+7'+s[:s.index('"')]
              s=s.replace("&nbsp;", "") 
              f1.write(s+"\n") 
              f2.write(s+"\n") 
        

              s=stemp
              s=s[s.index('Состав заказа')+2:]
              s=s[:s.index('Информация о курьерской службе')+2]

              
             
            
              # Reset variables 
              title=''
              sku=''
              price=''
              quantity=''
              quant_package=0
              

              
              # Loop through the string 
              while '<tr draggable="' in s:  
                 # Extract product title
                 s=s[s.index('Название в Kaspi Магазине')+2:]    
                 s=s[s.index('blank">')+8:]
                 s2=s
                 s=s[:s.index(' <')]
                 s=s.replace("&amp;", "") 

                 if len(title)>0:
                     title=title + '$#$' + s
                 else:
                     title=s
        
                 # Extract product SKU                  
                 s=s2        
                 s=s[s.index('Артикул')+4:] 
                 s=s[s.index('"">')+4:]  
                 s2=s
                 s=s[:s.index(' <')]
                 if len(sku)>0:
                      sku=sku + '$#$' + s
                 else:
                     sku=s

                 # Extract product quantity
                 s=s2      
                 s=s[s.index('Кол-во')+4:] 
                 s=s[s.index('--->')+5:] 
                 s2=s
                 s=s[:s.index(' <')]
                 quant_package=quant_package+int(s)
                 if len(quantity)>0:
                      quantity=quantity + '$#$' + s
                 else:
                     quantity=s
        
                 # Extract product price
                 s=s2          
                 s=s[s.index('Цена')+4:] 
                 s=s[s.index('">')+3:]   
                 s2=s
                 s=s[:s.index('<')-2]            
                 s=s.replace("&nbsp;", "") 
                 if len(price)>0:
                      price=price + '$#$' + s
                 else:
                     price=s  
       
                 s = s2



              # remove &nbsp; symbols
              title=title.replace("&nbsp;", "") 
              sku=sku.replace("&nbsp;", "") 
              quantity=quantity.replace("&nbsp;", "") 
              price=price.replace("&nbsp;", "")  


              # Write data to file 1
              f1.write(title+"\n") 
              f1.write(sku+"\n")   
              f1.write(quantity+"\n") 
              f1.write(price+"\n") 
        
              sku=sku + '$#$' + 'Parfum-kz-pac' 
              quantity=quantity + '$#$' + str(quant_package)
              price=price + '$#$' + '0'
              title=title + '$#$' + 'package'
        
        
              # Write data to file 2
              f2.write(title+"\n") 
              f2.write(sku+"\n") 
              f2.write(quantity+"\n") 
              f2.write(price+"\n")     
        
              # Extract order date and write to both files
              s=stemp
              s=s[s.index('Дата поступления заказа')+2:]
              s=s[s.index('ms;">')+6:]
              s=s[:s.index(' <')]
              s=s.replace("&nbsp;", "") 
              f1.write(s+"\n") 
              f2.write(s+"\n") 
        

              # Write empty line to both files
              s="empty"
              f1.write(s+"\n")
              f2.write(s+"\n")      
        
              # Write special values to file 1 based on the value of s3
              if ('Нет' in s3) or ('No' in s3):
                  f1.write("1\n") 
                  f1.write("1\n") 
              else:        
                  f1.write("2\n") 
                  f1.write("2\n") 
            
              f1.close()      
              f2.close()

            # Print the value of name2  
            if len(name2)>0:
                print(name2) 
    
         





# This line sets the directory path to the current working directory. 
dir_path ='C:/Release'

#dir_path = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__))) #change for debug
#dir_path = os.getcwd()  #change for EXE









# This line creates a web driver service for Chrome using ChromeDriverManager to install the driver.
s=Service(ChromeDriverManager().install())

#driver = webdriver.Firefox()

# This line creates a new Chrome webdriver with the service created above.
driver = webdriver.Chrome(service=s)





  
 
try:
   f4 = open(dir_path+"\ll_olds.txt", "a", encoding='utf-8') 
   f4.writelines("z")           
   f4.close()  
   login_kaspi()
   #The above line calls a function to log in to the website.


      
   time.sleep(5) 
   # This line navigates the driver to the page with NEW status orders.
   driver.get("https://kaspi.kz/mc/#/orders?status=NEW")
   time.sleep(5)     
     
   
   #Next lines iterate through all the pages orders and get their IDs.
   while get_lastpage():

      get_ids_kaspi()
      time.sleep(1)
      driver.execute_script("document.getElementsByClassName('pagination-link').item(1).click()") 
   
   #time.sleep(555) 
    

   get_ids_kaspi()
   time.sleep(1) 

   
   #Next lines iterate through all the pages with KASPI_DELIVERY_CARGO_ASSEMBLY status orders and get their IDs.
   driver.get("https://kaspi.kz/mc/#/orders?status=KASPI_DELIVERY_CARGO_ASSEMBLY") #  Change for auto accept
   time.sleep(5) 
   while get_lastpage():
      get_ids_kaspi()
      time.sleep(1) 
      driver.execute_script("document.getElementsByClassName('pagination-link').item(1).click()") 

   get_ids_kaspi()
   time.sleep(1) 

   # next lines go through the orders, process them and get their IDs.
   while len(kaspi_ids[cur_id])>0:
      next_kaspi_id_write() 
      driver.get("https://kaspi.kz/mc/#/orders?status=NEW")
      time.sleep(5)   
   driver.close()
   f4 = open(dir_path+"\\" + "reboot_get_id" + ".txt", "w", encoding='utf-8')                 
   f4.write("111"+"\n")  
   f4.close()


   #os._exit(0) 
except:
    pass    #Reverse1 


# close the driver and write to a file.
driver.close()
f4 = open(dir_path+"\\" + "reboot_get_id" + ".txt", "w", encoding='utf-8')                 
f4.write("111"+"\n")  
f4.close()


#instructions to create executable file through pyinstaller in CMD

#pyinstaller test.py  --onefile --add-binary "./driver/chromedriver.exe;./driver"
#pyinstaller test2.py --onefile
