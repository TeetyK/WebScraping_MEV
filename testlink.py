import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin

main_url = "https://index.tpso.go.th/cpi/index-analysis-report/1" 
soup = BeautifulSoup(response.content, 'html.parser')

pdf_link = ""

for link in soup.find_all('a', href=True):
    href = link['href']
    if href.endswith('.pdf'):
        pdf_link = urljoin(main_url, href)
        print(f"เจอลิงก์ PDF แล้ว: {pdf_link}")
        break

if not pdf_link:
    print("ไม่เจอลิงก์ PDF ")