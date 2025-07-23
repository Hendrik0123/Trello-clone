import requests
from bs4 import BeautifulSoup
import os
from dotenv import load_dotenv

load_dotenv()

def get_titles_from_url(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
    except Exception as e:
        print(f"Fehler beim Abrufen der URL: {e}")
        return []

    soup = BeautifulSoup(response.text, 'html.parser')
    title_list = []

    # Finde alle divs mit der Klasse 'header'
    headers = soup.find_all("div", class_="header")
    for header in headers:
        # Suche darin nach <h3><a title="...">
        h3 = header.find("h3")
        if h3:
            a_tag = h3.find("a", title=True)
            if a_tag:
                title_list.append(a_tag['title'])

    return title_list

# Beispielnutzung
if __name__ == "__main__":
    url = os.getenv("url")
    titles = get_titles_from_url(url)
    print("Gefundene Gruppen:")
    for title in titles:
        print("-", title)
