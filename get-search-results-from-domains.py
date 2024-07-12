import os
import requests
from bs4 import BeautifulSoup
from urllib.parse import quote, unquote
import pandas as pd
import time
import re
import random

def read_words(txt_file):
    with open(txt_file, 'r', encoding='utf-8') as file:
        words = file.read().splitlines()
    return words

def read_domains(excel_file):
    df = pd.read_excel(excel_file)
    return df['domains'].tolist()

def search_google(word, search_link, time_option='anytime', max_results=1, exact_match=False):
    found_links = []
    processed_urls = set()
    start = 0

    while len(found_links) < max_results:
        encoded_word = quote(f'"{word}"' if exact_match else word)
        search_url = f"https://www.google.com/search?q=site:{search_link}+{encoded_word}"
        if time_option != 'anytime':
            search_url += f"&tbs=qdr:{time_option}"

        try:
            response = requests.get(search_url)
            response.raise_for_status()

            if response.status_code == 200:
                soup = BeautifulSoup(response.content, "html.parser")
                search_results = soup.find_all("a")
                links_found = 0

                for result in search_results:
                    href = result.get("href")
                    if href and href.startswith("/url?q="):
                        url = href.split("/url?q=")[1].split("&sa=")[0]
                        url = unquote(url)
                        if url not in processed_urls and not url.startswith(
                                ('data:image', 'javascript', '#', 'https://maps.google.com/',
                                 'https://accounts.google.com/', 'https://www.google.com/preferences',
                                 'https://policies.google.com/', 'https://support.google.com/')):
                            found_links.append({'link': url})
                            processed_urls.add(url)
                            links_found += 1
                            if len(found_links) >= max_results:
                                break

                if links_found == 0:
                    break  # No new links found, exit the loop

            start += 10
            time.sleep(random.uniform(1.0, 3.0))

        except requests.exceptions.HTTPError as e:
            print(f"HTTP Error occurred: {e}")
            break
        except Exception as e:
            print(f"An error occurred: {e}")
            break

    return found_links

def scrape_page_content(url):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        if response.status_code == 200:
            page_content = preprocess_text(response.text)
        return page_content
    except requests.exceptions.RequestException as e:
        print(f"Failed to retrieve content from {url}: {e}")
        return None


def preprocess_text(text):
    text = text.lower()
    text = re.sub(r'\W+', ' ', text)
    return text


def main():
    txt_file = 'words.txt'
    excel_file = 'domains.xlsx'
    num_results = 5

    words = read_words(txt_file)
    domains = read_domains(excel_file)

    all_results = {}

    for word in words:
        for domain in domains:
            results = search_google(word, domain)
            all_results[(word, domain)] = results
            time.sleep(1)  # To avoid getting blocked by Google

    # Now process the results and scrape content
    for key, results in all_results.items():
        word, domain = key
        for result in results:
            url = result['link']
            print(url)

if __name__ == '__main__':
    main()
