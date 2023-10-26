import time
import csv
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from termcolor import colored
from selenium.webdriver.common.by import By # https://stackoverflow.com/questions/69875125/find-element-by-commands-are-deprecated-in-selenium
import numpy as np
from selenium.webdriver.common.proxy import *
from selenium import webdriver
import random
from selenium.webdriver.common.proxy import Proxy, ProxyType
print("Current selenium version is:", selenium.__version__)
# Select webdriver profile to use in selenium or not.
import sys
import os
import undetected_chromedriver.v2 as uc
import time
# ux-ui
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import scrolledtext 
from tkinter import messagebox 
from termcolor import colored

companies = {}

def driver_Profile(Profile_name):
    if Profile_name == "Yes":
               ### Selenium Web Driver Chrome Profile in Python
        # set proxy and other prefs.
        profile = webdriver.FirefoxProfile()
        profile.set_preference("network.proxy.type", 1)
        profile.set_preference("network.proxy.http",allproxs[0]['IP Address'])
        profile.set_preference("network.proxy.http_port", int(allproxs[0]['Port']))
        # update to profile. you can do it repeated. FF driver will take it.
        profile.set_preference("network.proxy.ssl", allproxs[0]['IP Address']);
        profile.set_preference("network.proxy.ssl_port", int(allproxs[0]['Port']))
        # You would also like to block flash
        profile.set_preference("media.peerconnection.enabled", False)
        # save to FF profile
        profile.update_preferences()
        driver = webdriver.Firefox(options=firefox_options)
        ### Check the version of Selenium currently installed, from Python
        print("Current selenium version is:", selenium.__version__)
        print("Current web browser is", driver.name)
        ### datetime object containing current date and time
        now = datetime.now()
        ### dd/mm/YY H:M:S
        dt_string = now.strftime("%B %d, %Y    %H:%M:%S")
        print("Current date & time is:", dt_string)
        print(colored("You are using webdriver profile!", "red"))

    else:
        firefox_options = webdriver.FirefoxOptions()
        firefox_options.headless = True  # Bật chế độ headless
        driver = webdriver.Firefox(options=firefox_options)
        ### Check the version of Selenium currently installed, from Python
        print("Current selenium version is:", selenium.__version__)
        print("Current web browser is", driver.name)
        now = datetime.now()
        ### dd/mm/YY H:M:S
        dt_string = now.strftime("%B %d, %Y    %H:%M:%S")
        print("Current date & time is:", dt_string)
        print(colored("You are NOT using webdriver profile!", "red"))
    return driver

def open_file_dialog():
    # Mở hộp thoại chọn tệp CSV
    csv_file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if csv_file:
        # Hiển thị tên tệp đã chọn trên giao diện
        selected_file_label.config(text="Selected CSV File: " + csv_file)
        global selected_csv_file
        selected_csv_file = csv_file
def save_file_dialog():
    # Mở hộp thoại chọn nơi lưu tệp npdetails.csv
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if save_path:
        # Lưu đường dẫn đến tệp đã chọn vào biến global
        save_file_label.config(text="Selected Save CSV File: " + save_path) 
        global save_csv_file
        save_csv_file = save_path

def save_file_function(companies):
    if companies:
        npdetails = pd.DataFrame.from_dict(companies).transpose()
        if 'save_csv_file' in globals():
            npdetails.to_csv(save_csv_file)
            messagebox.showinfo("Information", "Save Success.")
            return
        else :
            messagebox.showinfo("Information", "Please select a CSV file to save.")
            return
    else : 
        messagebox.showinfo("Information", "Please run program before")
        return
    
# place holder
def on_entry_click(event):
    if positions_entry.get("1.0", "end-1c") == "President, Chief Executive Officer, Chief Technology Officer, Chief Operating Officer":
        positions_entry.delete("1.0", "end-1c")  # Xóa placeholder
        positions_entry.config(fg="black")  # Thay đổi màu văn bản thành đen

def on_focus_out(event):
    if not positions_entry.get("1.0", "end-1c").strip():  # Nếu không có văn bản, thêm lại placeholder
        positions_entry.config(fg="grey")  # Thay đổi màu văn bản thành xám

def Google_search(driver,board_of_directors, csv_file_path, save_csv_file):
    # board_of_directors = ["President", "Chief Executive Officer", "Chief Technology Officer",
    if csv_file_path:
        with open(csv_file_path, 'r') as csvfile: 
            # 
            if save_csv_file:
                with open(save_csv_file, 'w', newline='') as output_csv:
                    fieldnames = ['Company', 'Position', 'Name', 'LinkedIn Link']
                    writer = csv.DictWriter(output_csv, fieldnames=fieldnames)
                    writer.writeheader()
    # with open('nonprofit.csv', 'r') as csvfile:
            datareader = csv.reader(csvfile)
            for row in datareader:
                company = ''.join(row)
                companies[company] = {}
                for position in board_of_directors:
                    url = "http://www.google.com/search?q=" + " " + company + " " + position + " linkedin"
                    driver.get(url)
                    print(url)
                    time.sleep(3)
                    # # Tìm phần tử đầu tiên có thẻ <div> và class="yuRUbf" bằng XPath
                    first_div_element = driver.find_elements(By.XPATH,'//div[@class="yuRUbf"]')
                    time.sleep(2)
                    soup = BeautifulSoup(driver.page_source, 'html.parser')
                    website = soup.find('div', class_="yuRUbf")
                    time.sleep(2)
                    if type(website) != None:
                        link = website.a.get('href')
                    if first_div_element:
                        # Tìm tất cả các phần tử <h3> trong phần tử <div>
                        span_elements = first_div_element[0].find_elements(By.TAG_NAME, 'h3')
                        if span_elements:
                            name = span_elements[0].text.split('-')[0]
                    
                    # 
                    if save_csv_file:
                        with open(save_csv_file, 'a', newline='') as output_csv:
                            writer = csv.DictWriter(output_csv, fieldnames=fieldnames)
                            writer.writerow({'Company': company, 'Position': position, 'Name': name, 'LinkedIn Link': link})
                    # 
                    companies[company][position] = name + link
                    time.sleep(3)
    else:
        messagebox.showinfo("Thông báo", "File csv không khả dụng")
    # Lưu trạng thái vào tệp

    return companies

def start_search():
    start_time = datetime.now()
    # Kiểm tra xem đã chọn tệp CSV chưa
    if 'selected_csv_file' not in globals() and 'save_csv_file' not in globals():
        messagebox.showinfo("Information", "Please select a CSV file input and CSV file output.")
        return
    driver = driver_Profile('No')
    positions = positions_entry.get("1.0", "end-1c").split(",") 
    try:
        companies = Google_search(driver, positions,selected_csv_file, save_csv_file)
    except Exception as e:
        messagebox.showinfo("Information", str(e))
    # npdetails.to_csv('npdetails.csv')
    driver.quit()
    end_time = datetime.now()
    result_label.config(text='Duration time: {} seconds'.format(end_time - start_time), fg="blue")
# Tạo cửa sổ giao diện
window = tk.Tk()
window.title("Google Search Tool")
window.geometry("400x600") 

# Label và Entry cho việc chọn profile
profile_label = tk.Label(window, text="Use WebDriver Profile")
profile_label.pack()

# Label và Entry để nhập danh sách vị trí
positions_label = tk.Label(window, text="Enter Positions (comma-separated):")
positions_label.pack()

# positions_entry = tk.Text(window, width=30, height=10)
# positions_entry.pack()
positions_entry = tk.Text(window, width=30, height=10, fg="grey")
positions_entry.insert("1.0", "President, Chief Executive Officer, Chief Technology Officer, Chief Operating Officer")  # Thêm placeholder
positions_entry.bind("<FocusIn>", on_entry_click)
positions_entry.bind("<FocusOut>", on_focus_out)
positions_entry.pack()

# Button để bắt đầu tìm kiếm
search_button = tk.Button(window, text="Start Search", command=start_search)
search_button.pack()

# Nút cho việc chọn tệp
choose_file_button = tk.Button(window, text="Choose CSV File", command=open_file_dialog, width=150)
choose_file_button.pack(padx=20, pady=5)
# Label để hiển thị tệp đã chọn
selected_file_label = tk.Label(window, text="Selected CSV File: None")
selected_file_label.pack()

# Nút cho việc chọn nơi lưu tệp npdetails.csv
save_file_button = tk.Button(window, text="Select Save Path", command=save_file_dialog, width=150)
save_file_button.pack(padx=20, pady=5)
# Label để hiển thị nơi lưu tệp npdetails.csv
save_file_label = tk.Label(window, text="Selected Save Path: None")
save_file_label.pack()

# Nút cho lưu file:
btn_save = tk.Button(window, text="Save", command=lambda:save_file_function(companies), width=150)
btn_save.pack(padx=20, pady=5)
# Label để hiển thị kết quả
result_label = tk.Label(window, text="")
result_label.pack()

# Bắt đầu ứng dụng
start_time = None
if __name__ == '__main__':
    window.mainloop()

