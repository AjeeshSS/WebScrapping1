from bs4 import BeautifulSoup
import requests, openpyxl


excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top Rated Movies'
print(excel.sheetnames)
sheet.append(['Movie name', 'Movie year', 'Movie rating'])


try:
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    }
    
    source = requests.get('https://www.imdb.com/chart/top/', headers=headers)
    source.raise_for_status() 
    
    soup = BeautifulSoup(source.text, 'html.parser')
    
    movie_items = soup.find_all('li', class_='ipc-metadata-list-summary-item')

    for item in movie_items:
        title_tag = item.find('h3', class_='ipc-title__text')
        
        year_tag = item.find('span', class_='cli-title-metadata-item')
        
        rating_tag = item.find('span', class_='ipc-rating-star--rating')

        if title_tag and year_tag and rating_tag:
            title = title_tag.text
            year = year_tag.text
            rating = rating_tag.text
            
            print(f'Title: {title}, Year: {year}, Rating: {rating}')
            sheet.append([title, year, rating])

except Exception as e:
    print(e)

excel.save('movie ratings.xlsx')