import requests
from bs4 import BeautifulSoup
import csv
import pandas as pd
import sys

class Product:
    def __init__(self, mpn, price=None):
        self.mpn = mpn
        self.price = price

    def __repr__(self):
        return f"Product(MPN={self.mpn}, Price={self.price})"

def convert_excel_to_csv(input_excel_file, output_csv_file):
    df = pd.read_excel(input_excel_file)
    df.to_csv(output_csv_file, index=False)

def read_mpn_from_csv(file_path):
    products = []
    with open(file_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        next(reader)  # Skip header row if present
        for row in reader:
            products.append(Product(mpn=row[0]))
    return products

def scrape_price_for_mpn(mpn):
    """
    Simulates scraping by generating a fake price. Replace this function with actual web scraping logic.
    """
    url = f"https://example.com/search?q={mpn}"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        # Simulate scraping by extracting fake price (replace with actual logic)
        price = soup.find('span', class_='price').text.strip()
        return price
    except Exception as e:
        print(f"Error scraping {mpn}: {e}")
        return None

def scrape_prices(products):
    for product in products:
        product.price = scrape_price_for_mpn(product.mpn)

def write_prices_to_csv(products, output_file):
    with open(output_file, mode='w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["MPN", "Price"])
        for product in products:
            writer.writerow([product.mpn, product.price])

def write_prices_to_excel(products, output_file):
    data = [{"MPN": product.mpn, "Price": product.price} for product in products]
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python script.py <input_excel_file>")
        sys.exit(1)

    # Input and output file paths
    input_excel_file = sys.argv[1]  # Input Excel file passed as an argument
    temp_csv_file = "mpns.csv"
    output_csv_file = "prices.csv"
    output_excel_file = "prices.xlsx"

    # Convert Excel to CSV
    convert_excel_to_csv(input_excel_file, temp_csv_file)

    # Read MPNs from CSV
    products = read_mpn_from_csv(temp_csv_file)

    # Scrape prices for each MPN
    scrape_prices(products)

    # Write aggregated prices to an output CSV file
    write_prices_to_csv(products, output_csv_file)

    # Write aggregated prices to an output Excel file
    write_prices_to_excel(products, output_excel_file)

    print(f"Prices scraped and written to {output_csv_file} and {output_excel_file}")
