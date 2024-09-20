import requests
from bs4 import BeautifulSoup
import pandas as pd
from urllib.parse import urlparse, urljoin
import os

def get_absolute_url(base_url, link):
    return urljoin(base_url, link)

def filter_links(links, domain_name):
    # Filter out non-HTTP/S links
    return [link for link in links if domain_name in link and link.startswith(('http://', 'https://'))]

def clean_text(text):
    # Remove '\n' and replace multiple spaces with single space
    return ' '.join(text.replace('\n', ' ').split())

def word_count(text):
    # Count the number of words in the text
    return len(text.split())

def scrape_page(session, url):
    try:
        response = session.get(url)
        response.raise_for_status()
    except requests.RequestException as e:
        print(f"Request failed: {e}")
        return

    # Parse the content with BeautifulSoup
    soup = BeautifulSoup(response.content, 'html.parser')

    # Collecting H1 to H6 headers with additional information
    headers_data = {}
    for i in range(1, 7):
        headers = [header.get_text(separator=' ') for header in soup.find_all(f'h{i}')]
        cleaned_headers = [clean_text(header) for header in headers]
        headers_data[f'H{i}'] = '\t'.join(cleaned_headers)
        headers_data[f'H{i}_count'] = len(headers)
        headers_data[f'H{i}_words'] = sum(word_count(header) for header in cleaned_headers)

    # Collecting button texts
    buttons = [button.get_text(separator=' ') for button in soup.find_all('button')]
    cleaned_buttons = [clean_text(button) for button in buttons]
    buttons_text = '\t'.join(cleaned_buttons)
    buttons_count = len(buttons)
    buttons_words = sum(word_count(button) for button in cleaned_buttons)

    # Collecting text content excluding scripts, styles, and other non-visible elements
    for script in soup(["script", "style"]):
        script.decompose()
    text_content = clean_text(soup.get_text(separator=' '))
    text_words = word_count(text_content)

    return {
        'URL': url,
        **headers_data,
        'Buttons': buttons_text,
        'Buttons_count': buttons_count,
        'Buttons_words': buttons_words,
        'Text Content': text_content,
        'Text_words': text_words,
        'Month': 1
    }

def scrape_website(url):
    session = requests.Session()
    # Set a dummy cookie to simulate accepting cookies; customize this as per actual cookie requirements
    session.cookies.set('cookie_consent', 'accepted')

    scraped_data = []
    urls_to_scrape = [url]
    scraped_urls = set()

    parsed_url = urlparse(url)
    domain_name = parsed_url.netloc

    while urls_to_scrape:
        current_url = urls_to_scrape.pop(0)
        if current_url in scraped_urls:
            continue
        
        scraped_urls.add(current_url)
        page_data = scrape_page(session, current_url)
        if page_data:
            scraped_data.append(page_data)

        try:
            response = session.get(current_url)
            response.raise_for_status()
        except requests.RequestException as e:
            print(f"Failed to fetch {current_url}: {e}")
            continue

        soup = BeautifulSoup(response.content, 'html.parser')
        links = [get_absolute_url(current_url, a.get('href')) for a in soup.find_all('a', href=True)]
        filtered_links = filter_links(links, domain_name)

        for link in filtered_links:
            if link not in scraped_urls and link not in urls_to_scrape:
                urls_to_scrape.append(link)

    return scraped_data, len(scraped_urls)

def export_to_excel(data, file_name):
    # Check if the file exists
    if os.path.exists(file_name):
        # If the file exists, load its content
        existing_df = pd.read_excel(file_name)
        # Create a new DataFrame from the new data
        new_df = pd.DataFrame(data)
        # Append the new data to the existing data
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        # If the file doesn't exist, just create a new DataFrame from the new data
        combined_df = pd.DataFrame(data)
    
    # Save the combined data back to the Excel file
    combined_df.to_excel(file_name, index=False)

def main():
    user_url = input("Enter the URL of the website to scrape: ")
    data, url_count = scrape_website(user_url)
    export_to_excel(data, 'scraped_data.xlsx')
    print(f'The number of URLs scraped: {url_count}')

if __name__ == '__main__':
    main()