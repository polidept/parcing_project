import requests
from bs4 import BeautifulSoup as BS
import openpyxl

def get_html(url):
    response = requests.get(url)
    if response.status_code == 200:
        return response.text
    return None

def get_links(html):
    soup = BS(html, 'html.parser')
    container = soup.find('div', class_ = 'results-grid -show')
    posts = container.find_all('a', class_ = 'product-link lsco-col-xs-offset-1 lsco-col-lg-offset-0 lsco-col-xs-offset-right-1 lsco-col-lg-offset-right-0')
    links = []
    for post in posts:
        link = post.get('href')
        full_link = 'https://www.levi.com' + link
        links.append(full_link)
    return links

def get_posts(html):
    soup = BS(html, 'html.parser')
    title = soup.find('div', class_ = 'page-title lsco-col-xs-offset-2 lsco-col-xs-offset-right-2 lsco-col-md-offset-0 lsco-col-md-offset-right-0 hide-mobile')
    name = title.find('h1', class_ = 'product-title').text
    price = title.find('span', class_ = 'price').text.strip()
    color_section = soup.find('div', class_ = 'swatches-section')
    color = color_section.find('span', class_ = 'swatchName').text.strip()
    overview_section = soup.find('div', class_ = 'product-overview')
    try:
        overview = overview_section.find('p').text.strip()
    except:
        overview = '-'

    data = {
        'model' : name,
        'price' : price,
        'color' : color,
        'overview' : overview
    }
    return data
    
def save_to_xls(data):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet['A1'] = 'Model'
    sheet['B1'] = 'Price'
    sheet['C1'] = 'Color'
    sheet['D1'] = 'Overview'

    for i, item in enumerate(data, 4):
        sheet[f'A{i}'] = item['model']
        sheet[f'B{i}'] = item['price']
        sheet[f'C{i}'] = item['color']
        sheet[f'D{i}'] = item['overview']
    
    wb.save('jeans.xlsx')

def main():
    URL = 'https://www.levi.com/US/en_US/clothing/men/jeans/c/levi_clothing_men_jeans'
    html = get_html(URL)
    links = get_links(html)
    data = []
    for page in range(4):
        page_url = URL + f'?page={page}'
        html = get_html(page_url)
        links = get_links(html)
        for link in links:
            detail_html = get_html(link)
            data.append(get_posts(detail_html))

    save_to_xls(data = data)
    
if __name__ == '__main__':
    main()
