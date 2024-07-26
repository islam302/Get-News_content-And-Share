# -*- coding: utf-8 -*-

from tkinter import Tk, ttk, filedialog, simpledialog, messagebox, Button, Label, PhotoImage, Frame, font, Entry, Canvas, NW
from bs4 import BeautifulSoup
import time
import pyautogui
import random
import ttkbootstrap as TTK
from ChromeDriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException, ElementClickInterceptedException
import re
from selenium.webdriver.common.action_chains import ActionChains
from tkinter import filedialog
from tkinter import messagebox
import tkinter as tk
from requests.packages.urllib3.exceptions import InsecureRequestWarning
from PIL import Image, ImageTk
from selenium.webdriver.common.action_chains import ActionChains
from docx import Document
import pandas as pd
import xlsxwriter
import logging
import chardet
import datetime
import psutil
import urllib
import requests
import base64
import glob
import sys
import os
import codecs
from urllib.parse import quote, unquote, urlparse


class SearchAboutNews(Tk):

    def __init__(self):
        super().__init__()

        self.include_iframe_var = tk.BooleanVar()
        self.include_iframe_var.set(True)
        self.title("SHARE NEWS")
        self.geometry("600x900")
        self.configure(bg="#282828")
        self.style = TTK.Style()
        self.style.theme_use("darkly")
        self.style.configure('TButton', background='blue', foreground='white')

        self.user_agents = [
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
        ]

        self.driver = None
        self.current_dir = os.path.dirname(sys.argv[0])
        self.results_folder = os.path.join(self.current_dir, 'RESULTS')
        os.makedirs(self.results_folder, exist_ok=True)

        self.templates = []
        self.create_widgets()

    def create_widgets(self):
        # Set logo as window background
        current_dir = os.path.dirname(sys.argv[0])
        logo_files = glob.glob(os.path.join(current_dir, "logo.*"))
        if logo_files:
            logo_path = logo_files[0]
            try:
                logo_image = Image.open(logo_path)
                logo_image = logo_image.resize((600, 900), Image.LANCZOS)
                self.background_image = ImageTk.PhotoImage(logo_image)
                bg_label = tk.Label(self, image=self.background_image)
                bg_label.place(x=0, y=0, relwidth=1, relheight=1)
            except Exception as e:
                print(f"Error loading image: {e}")
                messagebox.showerror("Error", "Failed to load logo image")

        self.title("Searching Links Extractor")
        self.geometry("600x900")
        self.configure(bg="#282828")

        custom_font = font.Font(family="Helvetica", size=14, weight="bold")

        btn_style = {
            'bg': '#006400',
            'fg': 'white',
            'padx': 20,
            'pady': 10,
            'bd': 0,
            'borderwidth': 0,
            'highlightthickness': 0,
            'font': custom_font
        }

        self.task2_button = tk.Button(self, text='Run Searching TASK', command=self.execute_task, **btn_style)
        self.task2_button.pack(pady=(200, 10))  

        self.template_frame = tk.Frame(self, bg='#282828')
        self.template_frame.pack(pady=4)

        exit_button = tk.Button(self, text="Exit", command=self.destroy, bg="red", fg="white", font=("Arial", 20))
        exit_button.pack(pady=4)

        self.style = ttk.Style()
        self.style.configure('TButton', background='#006400', foreground='white')
        self.style.configure('TFrame', background='#282828')

    def encode_image_to_base64(self, image_path):
        with open(image_path, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
        return encoded_string

    def start_driver(self):
        self.driver = WebDriver.start_driver(self)
        return self.driver

    def killDriverZombies(self, driver_pid):
        try:
            parent_process = psutil.Process(driver_pid)
            children = parent_process.children(recursive=True)
            for process in [parent_process] + children:
                process.terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

    def select_file(self):
        file_path = filedialog.askopenfilename(title="Select Search Engine Links File",
                                               filetypes=[("All files", "*.*")])
        return file_path

    def select_folder(self):
        folder_path = filedialog.askdirectory()  # This should return a string path
        print(f"Selected folder path: {folder_path}")
        return folder_path

    def get_domains_from_file(self, excel_file_path):
        try:
            df = pd.read_excel(excel_file_path)
            domains = df['search link'].tolist()
            return domains
        except FileNotFoundError:
            print(f'File {excel_file_path} not found.')
            return []
        except Exception as e:
            print(f'An error occurred: {e}')
            return []

    def get_title(self, link, headers, retries=3, timeout=10):
        for attempt in range(retries):
            try:
                requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
                response = requests.get(link, headers=headers, verify=False, timeout=timeout)
                print(f'-{link} : {response}')
                if response.status_code == 200:
                    soup = BeautifulSoup(response.content, 'html.parser')
                    title = soup.title.string.strip() if soup.title else None
                    if not title:
                        h1_tag = soup.find('h1')
                        if h1_tag:
                            title = h1_tag.get_text(strip=True)
                    return title if title else 'No title found'
            except requests.exceptions.RequestException as e:
                print(f'Attempt {attempt + 1} failed for {link}: {e}')
                time.sleep(2)  # Wait for 2 seconds before retrying
        return 'No title found'

    def get_publish_date(self, link, headers, retries=3, timeout=10):
        for attempt in range(retries):
            try:
                requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
                response = requests.get(link, headers=headers, verify=False, timeout=timeout)
                if response.status_code == 200:
                    encoding = chardet.detect(response.content)['encoding']
                    response.encoding = encoding
                    html_content = response.text
                    soup = BeautifulSoup(html_content, 'html.parser')

                    date_patterns = [
                        r'\b\d{1,2}\s+\w+\s+\d{4}\b',
                        r'\b\d{4}/\d{2}/\d{2}\b',
                        r'\b\d{1,2}/\d{1,2}/\d{2,4}\b',
                        r'\b\d{1,2}-\d{1,2}-\d{4}\b',
                        r'\b\d{1,2}\s+\w+\s+\d{4}\b',
                        r'\b\d{4}-\d{2}-\d{2}\b'
                    ]

                    time_tags = soup.find_all('time')
                    for time_tag in time_tags:
                        datetime_attr = time_tag.get('datetime')
                        if datetime_attr:
                            return datetime_attr.strip()

                    for pattern in date_patterns:
                        date_match = re.search(pattern, html_content, re.IGNORECASE)
                        if date_match:
                            return date_match.group().strip()

                    link_date_match = re.search(r'(\d{4}-\d{2}-\d{2})', link)
                    if (link_date_match):
                        return link_date_match.group()

                    return 'No date found'
            except requests.exceptions.RequestException as e:
                print(f'Attempt {attempt + 1} failed for {link}: {e}')
                time.sleep(2)  # Wait for 2 seconds before retrying
        return 'No date found'

    def get_article_content(self, link, headers, retries=3, timeout=10):
        for attempt in range(retries):
            try:
                requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
                response = requests.get(link, headers=headers, verify=False, timeout=timeout)
                if response.status_code == 200:
                    soup = BeautifulSoup(response.content, 'html.parser')

                    # ?????? ????? ????? ??????? ???????
                    content = soup.find('article') or soup.find('div', class_='content') or soup.find('main')
                    if not content:
                        content = soup.find('div', class_=re.compile(r'.*content.*', re.IGNORECASE))

                    if content:
                        # ????? ??????? ??? ??????? ????
                        for tag in content.find_all(['script', 'style', 'aside', 'footer']):
                            tag.decompose()
                        for tag in content.find_all(True, {
                            'class': re.compile(r'.*(ad|advertisement|sidebar|footer).*', re.IGNORECASE)}):
                            tag.decompose()

                        article_text = content.get_text(separator='\n', strip=True)
                        article_text = re.sub(r'\s+', ' ', article_text).strip()
                        return article_text[:10000]
                    else:
                        return 'No article content found'
            except requests.exceptions.RequestException as e:
                print(f'Attempt {attempt + 1} failed for {link}: {e}')
                time.sleep(2)  # Wait for 2 seconds before retrying
        return 'No article content found'

    def get_image(self, link, headers, retries=3, timeout=10):
        for attempt in range(retries):
            try:
                requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
                response = requests.get(link, headers=headers, verify=False, timeout=timeout)
                if response.status_code == 200:
                    soup = BeautifulSoup(response.content, 'html.parser')

                    # Check for Open Graph image
                    og_image = soup.find('meta', property='og:image')
                    if og_image and og_image.get('content'):
                        return og_image['content']

                    # Check for Twitter image
                    twitter_image = soup.find('meta', attrs={'name': 'twitter:image'})
                    if twitter_image and twitter_image.get('content'):
                        return twitter_image['content']

                    # Fallback to searching for the largest image in the article
                    images = soup.find_all('img')
                    if images:
                        main_image = max(images, key=lambda img: int(img.get('width', 0)) * int(img.get('height', 0)),
                                         default=None)
                        if main_image and main_image.get('src'):
                            return main_image['src']

                return 'No image found'
            except requests.exceptions.RequestException as e:
                print(f'Attempt {attempt + 1} failed for {link}: {e}')
                time.sleep(2)  # Wait for 2 seconds before retrying
        return 'No image found'

    def get_photos_from_folder(self, folder_path):
        # List all files in the folder
        files = os.listdir(folder_path)

        # Filter out only JPEG files
        jpg_files = [f for f in files if f.lower().endswith('.jpg')]

        # Sort files numerically by their number in the name
        jpg_files.sort(key=lambda x: int(os.path.splitext(x)[0]))

        # Load and return images
        images = []
        for file_name in jpg_files:
            file_path = os.path.join(folder_path, file_name)
            try:
                img = Image.open(file_path)
                images.append(img)
            except IOError as e:
                print(f"Error loading image {file_name}: {e}")

        return images

    def update_excel_with_content(self, links, folder_name_input):
        if isinstance(links, list):
            data = []
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
            }

            for link in links:
                title = self.get_title(link, headers)
                date = self.get_publish_date(link, headers)
                article = self.get_article_content(link, headers)
                image = self.get_image(link, headers)
                data.append({
                    'Link': link,
                    'Title': title,
                    'Article': article,
                    'Date': date,
                    'Image': image
                })

            df = pd.DataFrame(data)

            # Create folder name based on the current date and the provided input
            now = datetime.datetime.now()
            formatted_now = now.strftime("%Y-%m-%d_%H-%M")
            folder_name = f'Data-{formatted_now}-{folder_name_input}'.replace(':', '-').replace('"', '')
            folder_path = os.path.join(self.results_folder, folder_name)

            # Create the directory if it doesn't exist
            os.makedirs(folder_path, exist_ok=True)

            # Define the path for the Excel file
            file_name = f'data-{formatted_now}.xlsx'
            excel_path = os.path.join(folder_path, file_name)

            # Save DataFrame to the Excel file
            df.to_excel(excel_path, index=False)

            print(f'Data successfully written to {excel_path}')

            # Return the data for sharing
            return data
        else:
            print('No valid links to process.')
            return []

    def execute_task(self):
        links = self.get_domains_from_file('domains.xlsx')
        folder_name_input = 'extracted_data'

        if not folder_name_input:
            messagebox.showinfo("Error", "Folder name cannot be empty")
            return

        folder_phoro_path = self.select_folder()
        print(f"Selected folder path: {folder_phoro_path}")  # Debugging statement

        # Ensure folder_phoro_path is a string
        if not isinstance(folder_phoro_path, str):
            messagebox.showinfo("Error", "Invalid folder path selected.")
            return

        data = self.update_excel_with_content(links, folder_name_input)
        self.data = data  # Assign data to self.data

        self.share(data, folder_phoro_path)

        if self.data:
            self.share(data, folder_phoro_path)  # Ensure to pass the required arguments
            messagebox.showinfo("Task Completed", "Task completed successfully!")
        else:
            messagebox.showinfo("Error", "Failed to update Excel with content")

        messagebox.showinfo("Task Completed", "Task completed successfully!")

    def share(self, data, photos_folder_path):
        # Print debug information
        print(f"Type of photos_folder_path: {type(photos_folder_path)}")
        print(f"Value of photos_folder_path: {photos_folder_path}")

        # Check if data exists
        if not data:
            messagebox.showinfo("Error", "No data to share. Please run the task first.")
            return

        # Ensure photos_folder_path is a string
        if not isinstance(photos_folder_path, str):
            raise ValueError(f"Expected photos_folder_path to be a string, got {type(photos_folder_path).__name__}")

        driver = self.start_driver()

        # Open the login page
        driver.get('https://e.afaqai.com/wp-login.php')

        # Log in
        username_field = driver.find_element(By.ID, 'user_login')
        username_field.send_keys('editor-1')  # Replace with your username or email
        password_field = driver.find_element(By.ID, 'user_pass')
        password_field.send_keys('7f(I$QCg*ty#jo3JIwFo96bw')  # Replace with your password
        remember_me_checkbox = driver.find_element(By.ID, 'rememberme')
        if not remember_me_checkbox.is_selected():
            remember_me_checkbox.click()
        submit_button = driver.find_element(By.ID, 'wp-submit')
        submit_button.click()

        driver.refresh()
        # Wait for login to complete
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'wpadminbar'))
        )

        # Open the new post page
        driver.get('https://e.afaqai.com/wp-admin/post-new.php')

        # Wait for the page to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'title'))
        )

        # Loop through each row in the data and create a post
        for index, item in enumerate(data):
            # Enter the title
            title_input = driver.find_element(By.ID, 'title')
            title_input.clear()
            title_input.send_keys(item['Title'])

            # Switch to the content editor frame
            driver.switch_to.frame(driver.find_element(By.ID, 'content_ifr'))

            # Enter the content
            content_frame = driver.find_element(By.ID, 'tinymce')
            content_frame.clear()
            content_frame.send_keys(item['Article'])

            # Switch back to the main content
            driver.switch_to.default_content()

            # Click the "Add Media" button
            try:
                add_photo = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.ID, 'insert-media-button'))
                )
                add_photo.click()
            except TimeoutException:
                print("Timed out waiting for 'Add Media' button to be clickable.")
                driver.save_screenshot('screenshot.png')  # Save screenshot for debugging
                continue
            except ElementClickInterceptedException:
                print("Element click intercepted. Trying to handle the issue.")
                driver.execute_script("arguments[0].scrollIntoView(true);", add_photo)
                driver.execute_script("arguments[0].click();", add_photo)
                continue
            # هنا المشكلة ..؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟؟
            upload_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH,
                                            "//button[@type='button' and @role='tab' and contains(@class, 'media-menu-item') and @id='menu-item-upload' and @aria-selected='true' and text()='رفع ملفات']"))
            )
            upload_button.click()

            # Wait for the file upload dialog to appear
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'button.browser'))
            )

            select_files_button = driver.find_element(By.CSS_SELECTOR, 'button.browser')
            select_files_button.click()

            # Give time for the file dialog to open
            time.sleep(2)

            # Construct the full path for the image
            image_filename = f'photo{index + 1}.webp'  # Assuming photos are named photo1.jpg, photo2.jpg, etc.
            image_path = os.path.join(photos_folder_path, image_filename)
            image_path = os.path.normpath(image_path)  # Normalize the path

            # Verify the constructed path
            if not os.path.isfile(image_path):
                print(f"Image file not found: {image_path}")
                continue

            # Use pyautogui to interact with the file dialog
            pyautogui.write(image_path)
            pyautogui.press('enter')

            button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(@class, 'media-button-insert') and text()='إدراج في المقالة']"))
            )
            # Click the button
            button.click()



            button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH,
                                            "//input[@type='submit' and @name='save' and @id='save-post' and @value='حفظ المسودة' and contains(@class, 'button')]"))
            )
            # Click the input button
            button.click()

            print(f"Published post: {item['Title']}")

        # Close the browser
        driver.quit()

if __name__ == "__main__":
    app = SearchAboutNews()
    app.mainloop()

# def get_domains_from_file(excel_file_path):
#     try:
#         df = pd.read_excel(excel_file_path)
#         domains = df['search link'].tolist()
#         return domains
#     except FileNotFoundError:
#         print(f'File {excel_file_path} not found.')
#         return []
#     except Exception as e:
#         print(f'An error occurred: {e}')
#         return []
#
#
# def get_title(self, link, headers, retries=3, timeout=10):
#     for attempt in range(retries):
#         try:
#             requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
#             response = requests.get(link, headers=headers, verify=False, timeout=timeout)
#             print(f'-{link} : {response}')
#             if response.status_code == 200:
#                 soup = BeautifulSoup(response.content, 'html.parser')
#                 title = soup.title.string.strip() if soup.title else None
#                 if not title:
#                     h1_tag = soup.find('h1')
#                     if h1_tag:
#                         title = h1_tag.get_text(strip=True)
#                 return title if title else 'No title found'
#         except requests.exceptions.RequestException as e:
#             print(f'Attempt {attempt + 1} failed for {link}: {e}')
#             time.sleep(2)  # Wait for 2 seconds before retrying
#     return 'No title found'
#
#
# def get_publish_date(self, link, headers, retries=3, timeout=10):
#     for attempt in range(retries):
#         try:
#             requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
#             response = requests.get(link, headers=headers, verify=False, timeout=timeout)
#             if response.status_code == 200:
#                 encoding = chardet.detect(response.content)['encoding']
#                 response.encoding = encoding
#                 html_content = response.text
#                 soup = BeautifulSoup(html_content, 'html.parser')
#
#                 date_patterns = [
#                     r'\b\d{1,2}\s+\w+\s+\d{4}\b',
#                     r'\b\d{4}/\d{2}/\d{2}\b',
#                     r'\b\d{1,2}/\d{1,2}/\d{2,4}\b',
#                     r'\b\d{1,2}-\d{1,2}-\d{4}\b',
#                     r'\b\d{1,2}\s+\w+\s+\d{4}\b',
#                     r'\b\d{4}-\d{2}-\d{2}\b'
#                 ]
#
#                 time_tags = soup.find_all('time')
#                 for time_tag in time_tags:
#                     datetime_attr = time_tag.get('datetime')
#                     if datetime_attr:
#                         return datetime_attr.strip()
#
#                 for pattern in date_patterns:
#                     date_match = re.search(pattern, html_content, re.IGNORECASE)
#                     if date_match:
#                         return date_match.group().strip()
#
#                 link_date_match = re.search(r'(\d{4}-\d{2}-\d{2})', link)
#                 if (link_date_match):
#                     return link_date_match.group()
#
#                 return 'No date found'
#         except requests.exceptions.RequestException as e:
#             print(f'Attempt {attempt + 1} failed for {link}: {e}')
#             time.sleep(2)  # Wait for 2 seconds before retrying
#     return 'No date found'
#
# def get_article_content(self, link, headers, retries=3, timeout=10):
#     for attempt in range(retries):
#         try:
#             requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
#             response = requests.get(link, headers=headers, verify=False, timeout=timeout)
#             if response.status_code == 200:
#                 soup = BeautifulSoup(response.content, 'html.parser')
#
#                 # ?????? ????? ????? ??????? ???????
#                 content = soup.find('article') or soup.find('div', class_='content') or soup.find('main')
#                 if not content:
#                     content = soup.find('div', class_=re.compile(r'.*content.*', re.IGNORECASE))
#
#                 if content:
#                     # ????? ??????? ??? ??????? ????
#                     for tag in content.find_all(['script', 'style', 'aside', 'footer']):
#                         tag.decompose()
#                     for tag in content.find_all(True, {
#                         'class': re.compile(r'.*(ad|advertisement|sidebar|footer).*', re.IGNORECASE)}):
#                         tag.decompose()
#
#                     article_text = content.get_text(separator='\n', strip=True)
#                     article_text = re.sub(r'\s+', ' ', article_text).strip()
#                     return article_text[:10000]
#                 else:
#                     return 'No article content found'
#         except requests.exceptions.RequestException as e:
#             print(f'Attempt {attempt + 1} failed for {link}: {e}')
#             time.sleep(2)  # Wait for 2 seconds before retrying
#     return 'No article content found'

# def get_article_content(link, headers, retries=3, timeout=10):
#     for attempt in range(retries):
#         try:
#             requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
#             response = requests.get(link, headers=headers, verify=False, timeout=timeout)
#             if response.status_code == 200:
#                 soup = BeautifulSoup(response.content, 'html.parser')
#                 content = soup.find('article') or soup.find('div', class_='content') or soup.find('main')
#
#                 if content:
#                     article_text = content.get_text(separator='\n', strip=True)
#                     for tag in soup.find_all(True, {'class': re.compile(r'.*(ad|advertisement|sidebar|footer).*')}):
#                         tag.decompose()
#                     article_text = article_text.replace('\n', ' ').strip()
#                     return article_text[:10000]
#                 else:
#                     article_text = soup.get_text(separator='\n', strip=True)
#                     article_text = article_text.replace('\n', ' ').strip()
#                     return article_text[:10000]
#             return 'No article content found'
#         except requests.exceptions.RequestException as e:
#             print(f'Attempt {attempt + 1} failed for {link}: {e}')
#             time.sleep(2)  # Wait for 2 seconds before retrying
#     return 'No article content found'
# def get_main_image(link, headers, retries=3, timeout=10):
#     for attempt in range(retries):
#         try:
#             requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
#             response = requests.get(link, headers=headers, verify=False, timeout=timeout)
#             if response.status_code == 200:
#                 soup = BeautifulSoup(response.content, 'html.parser')
#
#                 # Check for Open Graph image
#                 og_image = soup.find('meta', property='og:image')
#                 if og_image and og_image.get('content'):
#                     return og_image['content']
#
#                 # Check for Twitter image
#                 twitter_image = soup.find('meta', attrs={'name': 'twitter:image'})
#                 if twitter_image and twitter_image.get('content'):
#                     return twitter_image['content']
#
#                 # Fallback to searching for the largest image in the article
#                 images = soup.find_all('img')
#                 if images:
#                     main_image = max(images, key=lambda img: int(img.get('width', 0)) * int(img.get('height', 0)), default=None)
#                     if main_image and main_image.get('src'):
#                         return main_image['src']
#
#             return 'No image found'
#         except requests.exceptions.RequestException as e:
#             print(f'Attempt {attempt + 1} failed for {link}: {e}')
#             time.sleep(2)  # Wait for 2 seconds before retrying
#     return 'No image found'
#
# def update_excel_with_content(excel_file_path, output_file_path):
#     links = get_domains_from_file(excel_file_path)
#     if isinstance(links, list):
#         data = []
#         headers = {
#             "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
#         }
#
#         for link in links:
#             title = get_title(link, headers)
#             date = get_publish_date(link, headers)
#             article = get_article_content(link, headers)
#             image = get_main_image(link, headers)
#             data.append({
#                 'Link': link,
#                 'Title': title,
#                 'Article': article,
#                 'Date': date,
#                 'Image': image
#             })
#
#         df = pd.DataFrame(data)
#         df.to_excel(output_file_path, index=False)
#         print(f'Data successfully written to {output_file_path}')
#     else:
#         print('No valid links to process.')
#
# update_excel_with_content('domains.xlsx', 'output.xlsx')

