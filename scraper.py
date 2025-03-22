import requests
from bs4 import BeautifulSoup

page = requests.get("https://seekingalpha.com/")

soup = BeautifulSoup(page.text, "html.parser")