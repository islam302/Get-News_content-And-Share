# from tkinter import Tk, ttk, filedialog, simpledialog, messagebox, Button, Label, PhotoImage, Frame, font, Entry, Canvas, NW
# from bs4 import BeautifulSoup
# import time
# import random
# import ttkbootstrap as TTK
# from ChromeDriver import WebDriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
# import re
# import tkinter as tk
# from requests.packages.urllib3.exceptions import InsecureRequestWarning
# from PIL import Image, ImageTk
# from selenium.webdriver.common.action_chains import ActionChains
# from docx import Document
# import pandas as pd
# import xlsxwriter
# import logging
# import chardet
# import datetime
# import psutil
# import urllib
# import requests
# import base64
# import glob
# import sys
# import os
# import codecs
# from urllib.parse import quote, unquote, urlparse
#
#
# class SearchAboutNews(Tk):
#
#     def __init__(self):
#         super().__init__()
#
#         self.include_iframe_var = tk.BooleanVar()
#         self.include_iframe_var.set(True)
#         self.title("Searching Links Extractor")
#         self.geometry("600x900")
#         self.configure(bg="#282828")
#         self.style = TTK.Style()
#         self.style.theme_use("darkly")
#         self.style.configure('TButton', background='blue', foreground='white')
#
#         self.user_agents = [
#             "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
#             "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
#             "Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36"
#         ]
#
#         self.driver = None
#         self.current_dir = os.path.dirname(sys.argv[0])
#         self.results_folder = os.path.join(self.current_dir, 'RESULTS')
#         os.makedirs(self.results_folder, exist_ok=True)
#
#         self.templates = []
#         self.create_widgets()
#
#     def create_widgets(self):
#         # Set logo as window background
#         current_dir = os.path.dirname(sys.argv[0])
#         logo_files = glob.glob(os.path.join(current_dir, "logo.*"))
#         if logo_files:
#             logo_path = logo_files[0]
#             try:
#                 logo_image = Image.open(logo_path)
#                 logo_image = logo_image.resize((600, 900), Image.LANCZOS)
#                 self.background_image = ImageTk.PhotoImage(logo_image)
#                 bg_label = tk.Label(self, image=self.background_image)
#                 bg_label.place(x=0, y=0, relwidth=1, relheight=1)
#             except Exception as e:
#                 print(f"Error loading image: {e}")
#                 messagebox.showerror("Error", "Failed to load logo image")
#
#         # Main window properties
#         self.title("Searching Links Extractor")
#         self.geometry("600x900")
#         self.configure(bg="#282828")
#
#         custom_font = font.Font(family="Helvetica", size=14, weight="bold")
#
#         btn_style = {
#             'bg': '#006400',
#             'fg': 'white',
#             'padx': 20,
#             'pady': 10,
#             'bd': 0,
#             'borderwidth': 0,
#             'highlightthickness': 0,
#             'font': custom_font
#         }
#
#         self.task2_button = tk.Button(self, text='Run Searching TASK', command=self.execute_task, **btn_style)
#         self.task2_button.pack(pady=(200, 10))  # Adjust pady to move the button down
#
#         self.template_frame = tk.Frame(self, bg='#282828')
#         self.template_frame.pack(pady=4)
#
#         self.template_entries = []
#         self.add_template_entry()
#
#         add_template_button = tk.Button(self, text='+', command=self.add_template_entry, **btn_style)
#         add_template_button.pack(pady=4)
#
#         time_display_options = ['اي وقت', 'اخر سنة', 'اخر شهر', 'اخر اسبوع', 'اخر يوم', 'اخر ساعة']
#         self.time_option_var = tk.StringVar()
#         self.time_option_menu = ttk.Combobox(self, textvariable=self.time_option_var, values=time_display_options,
#                                              state="readonly", font=("Arial", 14))
#         self.time_option_menu.set(time_display_options[0])
#         self.time_option_menu.pack(pady=5)
#
#         exit_button = tk.Button(self, text="Exit", command=self.destroy, bg="red", fg="white", font=("Arial", 20))
#         exit_button.pack(pady=4)
#
#         self.style = ttk.Style()
#         self.style.configure('TButton', background='#006400', foreground='white')
#         self.style.configure('TFrame', background='#282828')
#
#     def add_template_entry(self):
#         entry = Entry(self.template_frame, width=50)
#         entry.pack(pady=3)
#         self.template_entries.append(entry)
#
#     def get_templates(self):
#         self.templates = [entry.get() for entry in self.template_entries if entry.get()]
#         return self.templates
#
#     def execute_task(self):
#         news_articles = self.get_templates()
#         if not news_articles:
#             messagebox.showinfo("Error", "No words in this file")
#             return
#
#         time_option = self.time_option_var.get()
#         if not time_option:
#             return
#
#         # Map time_option to specific values
#         time_option_map = {
#             'اخر سنة': 'y',
#             'اخر شهر': 'm',
#             'اخر اسبوع': 'w',
#             'اخر يوم': 'd',
#             'اخر ساعة': 'h',
#             'اي وقت': 'anytime'
#         }
#
#         time_option = time_option_map.get(time_option, 'anytime')
#
#         max_results = self.select_max_results()
#
#         excluded_domains = self.get_excluded_domains('black-list.txt')
#
#         file = self.select_file()
#         domains = self.get_domains_from_file(file)
#
#         # Prompt user for folder name input
#         folder_name_input = simpledialog.askstring("Folder Name", "Enter folder name for search results:")
#
#         if not folder_name_input:
#             messagebox.showinfo("Error", "Folder name cannot be empty")
#             return
#
#         # Create a single folder for all results
#         now = datetime.datetime.now()
#         formatted_now = now.strftime("%Y-%m-%d & %H-%M")
#         folder_name = f'Search-Results-{formatted_now}-{folder_name_input}'.replace(':', '-').replace('"', '').encode('utf-8').decode('utf-8')
#         folder_path = os.path.join(self.results_folder, folder_name)
#         os.makedirs(folder_path, exist_ok=True)
#
#         all_data = []
#
#         for i, news_article in enumerate(news_articles, start=1):
#             if not news_article:
#                 print(f"Skipping article {i} due to missing title")
#                 continue
#
#             search_word_data = self.main(folder_name_input, folder_path, [news_article],domains, time_option, max_results, excluded_domains)
#             all_data.extend(search_word_data)
#
#         file_name = f'Domains-Search-Data-{formatted_now}-{folder_name_input}.xlsx'
#         excel_path = os.path.join(folder_path, file_name)
#
#         # Create DataFrame from all_data list
#         df_all_data = pd.DataFrame(all_data)
#
#         # Write DataFrame to Excel
#         writer_all = pd.ExcelWriter(excel_path, engine='xlsxwriter')
#         df_all_data.to_excel(writer_all, index=False)
#         worksheet_all = writer_all.sheets['Sheet1']
#         worksheet_all.set_column('A:C', 50)
#         writer_all._save()
#
#         messagebox.showinfo("Task Completed", "Task completed successfully!")
#
#     def get_templates(self):
#         self.templates = [entry.get() for entry in self.template_entries if entry.get()]
#         return self.templates
#
#     def encode_image_to_base64(self, image_path):
#         with open(image_path, "rb") as image_file:
#             encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
#         return encoded_string
#
#     def start_driver(self):
#         self.driver = WebDriver.start_driver(self)
#         return self.driver
#
#     def get_publish_date(self, link):
#         requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
#         try:
#             response = requests.get(link)
#             if response.status_code == 200:
#                 encoding = chardet.detect(response.content)['encoding']
#                 response.encoding = encoding
#                 html_content = response.text
#                 soup = BeautifulSoup(html_content, 'html.parser')
#
#                 link_text = soup.get_text()
#                 date_match = re.search(r'\b\d{1,2}\s+\w+\s+\d{4}\b', link_text, re.IGNORECASE | re.UNICODE)
#                 if date_match:
#                     link_date = date_match.group()
#                     return link_date.strip()
#
#                 date_patterns = [
#                     r'\b(\d{4}/\d{2}/\d{2})\b',
#                     r'\b(\d{1,2}/\d{1,2}/\d{2,4})\b',
#                     r'\b(\d{1,2}\s+\w+\s+\d{2,4})\b',
#                     r'\b(\d{4}-\d{2}-\d{2})\b',
#                     r'\b(\d{1,2}\s+\w+\s+\d{4})\b',
#                     r'\b(\d{1,2}\s+(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{2,4})\b',
#                     r'\b(\d{1,2}/\d{1,2}/\d{2,4}\s+\d{1,2}:\d{2})\b',
#                     r'\b(\d{1,2}\s+\w+\s+/\s+\w+\s+\d{2,4})\b',
#                     r'\b(\d{1,2}\s+\w+\s+\d{4}\s+\d{1,2}:\d{2}:\d{2})\b',
#                     r'\b(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})\b',
#                     r'\b(\d{1,2}/\d{1,2}/\d{2,4}\s+\d{1,2}:\d{2}:\d{2})\b',
#                     r'\b(\d{1,2}\s+\w+\s+\d{4}\s+\d{1,2}:\d{2}:\d{2})\b',
#                     r'\b(\d{1,2}\s+[?-?]+\s+\d{4})\b',
#                     r'\b(\d{1,2}/\d{1,2}/\d{2,4})\s+[?-?]+\s+\d{1,2}:\d{2}\b',
#                     r'\b(\d{1,2}\s+[\u0623-\u064a]+\s+\d{4})\b',
#                     r'\b(\d{1,2}\s+[\u0623-\u064a]+\s+/\s+[\u0623-\u064a]+\s+\d{2,4})\b',
#                     r'(\d{4}/\d{2}/\d{2}/)'
#                 ]
#
#                 for pattern in date_patterns:
#                     date_match = re.search(pattern, html_content, re.IGNORECASE | re.UNICODE)
#                     if date_match:
#                         link_date = date_match.group()
#                         return link_date.strip()
#
#                 time_tags = soup.find_all('time', class_=re.compile(r'.*'))
#                 for time_tag in time_tags:
#                     datetime_attr = time_tag.get('datetime')
#                     if datetime_attr:
#                         arabic_date = time_tag.text.strip()
#                         return arabic_date
#
#                 link_date_match = re.search(r'(\d{4}-\d{2}-\d{2})', link)
#                 if link_date_match:
#                     return link_date_match.group()
#
#             return None
#         except:
#             return None
#
#     def get_title(self, link):
#         try:
#             requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
#             response = requests.get(link)
#             if response.status_code == 200:
#                 encoding = chardet.detect(response.content)['encoding']
#                 response.encoding = encoding
#                 html_content = response.text
#                 soup = BeautifulSoup(html_content, 'html.parser')
#                 title = soup.title.string.strip()
#                 return title
#         except:
#             return None
#
#     def killDriverZombies(self, driver_pid):
#         try:
#             parent_process = psutil.Process(driver_pid)
#             children = parent_process.children(recursive=True)
#             for process in [parent_process] + children:
#                 process.terminate()
#         except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#             pass
#
#     def select_file(self):
#         file_path = filedialog.askopenfilename(title="Select Search Engine Links File",
#                                                filetypes=[("All files", "*.*")])
#         return file_path
#
#     def select_max_results(self):
#         max_results = simpledialog.askinteger("Max Results",
#                                               "Enter the maximum number of results to fetch:")
#         return max_results
#
#     def get_excluded_domains(self, domains_file):
#         try:
#             with open(domains_file, 'r') as file:
#                 excluded_domains = [line.strip() for line in file.readlines()]
#             return excluded_domains
#         except FileNotFoundError:
#             print(f"Domains file '{domains_file}' not found.")
#             return []
#
#     def get_domains_from_file(self, excel_file_path):
#         try:
#             df = pd.read_excel(excel_file_path)
#             domains = df['search link'].tolist()
#             return domains
#         except FileNotFoundError:
#             print(f'File {excel_file_path} not found.')
#             return []
#         except Exception as e:
#             return False
#
#     def google_domain_search(self, domain, word, time_option='anytime', max_results=10):
#         found_links = []
#         processed_urls = set()
#         start = 0
#         error_urls = set()
#         self.driver = self.start_driver()
#
#         while len(found_links) < max_results:
#             encoded_word = quote(word)
#             search_url = f'https://www.google.com/search?q="{encoded_word}"+site:{domain}&start={start}'
#             if time_option != 'anytime':
#                 search_url += f"&tbs=qdr:{time_option}"
#
#             print(search_url)
#
#             self.driver.get(search_url)
#             time.sleep(1)
#             no_results_element_1 = self.driver.find_elements(By.XPATH,
#                                                              "//div[@aria-level='2' and @role='heading' and contains(text(), 'لا توجد نتائج')]")
#             no_results_element_2 = self.driver.find_elements(By.XPATH,
#                                                              "//div[@class='mnr-c']//p[contains(text(), 'لم ينجح بحثك عن')]")
#
#             if no_results_element_1 or no_results_element_2:
#                 print('No results found.')
#                 error_urls.add(search_url)  # Add the search URL to error_urls set
#                 break
#             else:
#                 print('Processing search results.')
#
#             try:
#                 response = requests.get(search_url, headers={'User-Agent': 'Mozilla/5.0'})
#                 time.sleep(random.uniform(1.0, 3.0))
#                 response.raise_for_status()
#                 if response.status_code == 200:
#                     soup = BeautifulSoup(response.content, "html.parser")
#
#                     search_results = soup.find_all("a", href=True)
#                     links_found = 0
#
#                     for result in search_results:
#                         href = result.get("href")
#                         if href and href.startswith("/url?q="):
#                             url = href.split("/url?q=")[1].split("&sa=")[0]
#                             url = unquote(url)
#                             if url not in processed_urls and not url.startswith(
#                                     ('data:image', 'javascript', '#', 'https://maps.google.com/',
#                                      'https://accounts.google.com/', 'https://www.google.com/preferences',
#                                      'https://policies.google.com/', 'https://support.google.com/', '/search?q=')):
#                                 link_text = result.text.strip()
#                                 found_links.append({'link': url, 'link_text': link_text})
#                                 processed_urls.add(url)
#                                 links_found += 1
#                                 if len(found_links) >= max_results:
#                                     break
#
#                     if links_found == 0:
#                         error_urls.add(search_url)  # Add the search URL to error_urls set
#                         break
#
#                 start += 10
#                 time.sleep(random.uniform(1.0, 3.0))
#
#             except requests.exceptions.HTTPError as e:
#                 print(f"HTTP Error occurred: {e}")
#                 error_urls.add(search_url)  # Add the search URL to error_urls set
#                 break
#             except Exception as e:
#                 print(f"An error occurred: {e}")
#                 error_urls.add(search_url)  # Add the search URL to error_urls set
#                 break
#
#         if self.driver:
#             driver_pid = self.driver.service.process.pid
#             self.killDriverZombies(driver_pid)
#
#         return found_links, error_urls
#
#     def main(self, file_name_input, folder_path, search_words, domains, time_option, max_results, excluded_domains):
#         all_data = []
#         error_urls = set()
#
#         try:
#             for search_word in search_words:
#                 for domain in domains:
#                     found_links_all, errors = self.google_domain_search(domain, search_word, time_option, max_results)
#
#                     error_urls.update(errors)
#
#                     filtered_links = [link for link in found_links_all if
#                                       not any(domain in link['link'] for domain in excluded_domains)]
#
#                     for link in filtered_links:
#                         content, images, links = self.extract_content(link['link'])
#                         all_data.append({
#                             'Search Word': search_word,
#                             'url': link['link'],
#                             'content': content,
#                             'images': images,
#                             'links': links
#                         })
#
#         except Exception as e:
#             print("Error", f"An error occurred: {e}")
#             messagebox.showinfo("Error", "An error occurred. Please try again.")
#
#         finally:
#             # Save errors to a txt file
#             if error_urls:
#                 txt_file_path = os.path.join(folder_path, 'error_links.txt')
#                 with open(txt_file_path, 'w') as file:
#                     for url in error_urls:
#                         file.write(url + '\n')
#
#             # Save all data to a Word document
#             word_file_path = os.path.join(folder_path, f'{file_name_input}.docx')
#             self.save_to_word(word_file_path, all_data)
#
#         return all_data
#
#     def fetch_and_save_full_html_with_selenium(self, url, output_filename):
#         try:
#
#             driver = self.start_driver()
#
#             # Open the URL
#             driver.get(url)
#
#             # Get the full HTML content
#             html_content = driver.page_source
#
#             # Save the HTML content to a file
#             with open(output_filename, 'w', encoding='utf-8') as file:
#                 file.write(html_content)
#
#             print(f"Full HTML content saved to {output_filename}")
#
#             # Quit the WebDriver
#             driver.quit()
#
#         except Exception as e:
#             print(f"Failed to retrieve the URL: {e}")
#
# if __name__ == '__main__':
#     app = SearchAboutNews()
#     app.fetch_and_save_full_html_with_selenium('https://www.alwatan.com.sa/article/1150685', 'sabq.html')

import requests
from bs4 import BeautifulSoup
from requests.packages.urllib3.exceptions import InsecureRequestWarning
import pandas as pd
import re
import chardet
import time


def get_domains_from_file(excel_file_path):
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


def get_title(link, headers, retries=3, timeout=10):
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


def get_publish_date(link, headers, retries=3, timeout=10):
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




def get_article_content(link, headers, retries=3, timeout=10):
    for attempt in range(retries):
        try:
            requests.packages.urllib3.disable_warnings(InsecureRequestWarning)
            response = requests.get(link, headers=headers, verify=False, timeout=timeout)
            if response.status_code == 200:
                soup = BeautifulSoup(response.content, 'html.parser')

                # محاولة إيجاد محتوى المقالة الرئيسي
                content = soup.find('article') or soup.find('div', class_='content') or soup.find('main')
                if not content:
                    content = soup.find('div', class_=re.compile(r'.*content.*', re.IGNORECASE))

                if content:
                    # إزالة العناصر غير المرغوب فيها
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


def get_main_image(link, headers, retries=3, timeout=10):
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
                    main_image = max(images, key=lambda img: int(img.get('width', 0)) * int(img.get('height', 0)), default=None)
                    if main_image and main_image.get('src'):
                        return main_image['src']

            return 'No image found'
        except requests.exceptions.RequestException as e:
            print(f'Attempt {attempt + 1} failed for {link}: {e}')
            time.sleep(2)  # Wait for 2 seconds before retrying
    return 'No image found'

def update_excel_with_content(excel_file_path, output_file_path):
    links = get_domains_from_file(excel_file_path)
    if isinstance(links, list):
        data = []
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }

        for link in links:
            title = get_title(link, headers)
            date = get_publish_date(link, headers)
            article = get_article_content(link, headers)
            image = get_main_image(link, headers)
            data.append({
                'Link': link,
                'Title': title,
                'Article': article,
                'Date': date,
                'Image': image
            })

        df = pd.DataFrame(data)
        df.to_excel(output_file_path, index=False)
        print(f'Data successfully written to {output_file_path}')
    else:
        print('No valid links to process.')

update_excel_with_content('domains.xlsx', 'output.xlsx')

