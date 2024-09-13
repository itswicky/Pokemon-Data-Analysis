import openpyxl
import requests
from bs4 import BeautifulSoup
from urllib.parse import quote_plus
import os
import openpyxl

def scrape_data():
    print("Scraping Pikalytics data")
    url = 'https://www.pikalytics.com/pokedex/gen9ou/'

    # Make request and parse HTML
    response = requests.get(url)
    print("Response status code: ", response.status_code)
    soup = BeautifulSoup(response.content, 'html.parser')

    # Find pokemon on page
    pokemon_list = soup.find_all('ul', class_='list gen90u')
    print("Found", len(pokemon_list), "pokemon.")

    data = []

    #Iterate over each item
    for pokemon in pokemon_list:
        
        try:
            pokemon_name = pokemon.find('span', class_='pokemon-name').text.strip()
        except AttributeError:
            pokemon_name = 'N/A'
        
        try:
            usage_percent = pokemon.find('span', class_='margin-right-20').text.strip()
        except AttributeError:
            usage_percent = 'N/A'
        
        data.append([pokemon_name, usage_percent])

    print('Scraping complete.')
    return data

def write_to_excel(data, output):
    print('Writing data to Excel...')
    # Delete previous if exists
    if os.path.exists(output):
        os.remove(output)
        print('Deleted previous file.')
    
    # Create workbook
    wb = openpyxl.Workbook()
    ws = wb.active

    # Write headers
    headers = ['Name', 'Usage']
    ws.append(headers)

    # Write data
    for row in data:
        ws.append(row)

    # Save workbook
    wb.save(output)
    print("Data written successfully to:", output)

if __name__ == "__main__":
    output = 'pokemon_ou_usage.xlsx'

    print('Starting scraping process...')

    data = scrape_data()
    write_to_excel(data, output)

    print('Process completed successfully.')