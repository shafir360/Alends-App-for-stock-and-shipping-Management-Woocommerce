from woocommerce import API
import requests
import pandas as pd
import tkinter as tk
import pdf2image
import pytesseract
#import tesseract
from PIL import Image, ImageDraw, ImageFont
from fpdf import FPDF
import re
import os
from pop_ups import CustomDialog
from tkinter import ttk, filedialog, messagebox

class WooFunc:
    
    def __init__(self):
        self.load_api()
        
    def load_api(self):
        credentials_valid = False
        while not credentials_valid:
            if os.path.exists("credential"):
                with open('credential', 'r') as file:
                    lines = file.readlines()
                    store_url = lines[0].strip()
                    consumer_key = lines[1].strip()
                    consumer_secret = lines[2].strip()

                credentials_valid = self.verify_credentials(store_url, consumer_key, consumer_secret)
                if not credentials_valid:
                    os.remove("credential")  # Remove invalid credentials file
            else:
                root = tk.Tk()
                root.withdraw()  # Hide the root window
                dialog = CustomDialog(root)
                if dialog.result:
                    store_url, consumer_key, consumer_secret = dialog.result
                    with open('credential', 'w') as file:
                        file.write('\n'.join(dialog.result))
                    credentials_valid = self.verify_credentials(store_url, consumer_key, consumer_secret)
                    if not credentials_valid:
                        messagebox.showerror("Error", "Invalid credentials, please try again.")
                else:
                    print("No credentials provided")
                    return  # Exit if the user cancels the dialog

        self.wcapi = API(
            url=store_url,
            consumer_key=consumer_key,
            consumer_secret=consumer_secret,
            version="wc/v3"
        )

    def verify_credentials(self, store_url, consumer_key, consumer_secret):
        test_api = API(
            url=store_url,
            consumer_key=consumer_key,
            consumer_secret=consumer_secret,
            version="wc/v3"
        )
        try:
            response = test_api.get("products", params={"per_page": 1}).json()
            if isinstance(response, list):
                return True  # Credentials are valid
        except Exception as e:
            print(f"Verification error: {e}")
        return False  # Credentials are invalid
    
               
    def get_all_products(self):
        page = 1
        products = []
        while True:
            response = self.wcapi.get("products", params={"per_page": 100, "page": page}).json()
            if not response:
                break
            products.extend(response)
            page += 1
        return products

    
        
    def get_categories(self):
        
        cat = self.wcapi.get("products/categories").json()
        return cat
    
    def print_category(self):
        categories = self.get_categories()
        for category in categories:
            print(f"ID: {category['id']}, Name: {category['name']}")
    
    def get_product_variations(self,product_id):
        var_json =  self.wcapi.get(f"products/{product_id}/variations").json()
        return var_json
    
    def print_product_variations(self):
        all_products = self.get_all_products()
        for product in all_products:
            if product['type'] == 'variable':
                variations = self.get_product_variations(product['id'])
                for variation in variations:
                    # Extract the size information from the attributes
                    size_variations = [attr['option'] for attr in variation.get('attributes', []) if attr['name'] == 'Size']
                    size_info = ', '.join(size_variations)
                    print(f"Product ID: {product['id']}, Name: {product['name']}, Variation ID: {variation['id']}, Size: {size_info}, Regular Price: {variation['regular_price']}, Sale Price: {variation['sale_price']}")
            else:
                print(f"Single, Product ID: {product['id']}, Name: {product['name']}, Regular Price: {product['regular_price']}, Sale Price: {product['sale_price']}")

    def get_variation_id(self,product_id, size):
        variations = self.wcapi.get(f"products/{product_id}/variations").json()
        for variation in variations:
            for attribute in variation.get('attributes', []):
                if attribute.get('name').lower() == 'size' and attribute.get('option') == size:
                    return variation['id']
        return None
    
    def get_product_id(self, product_name):
        products = self.wcapi.get("products", params={"search": product_name}).json()
        print(f"Searching for product '{product_name}', found: {products}")
        return products[0]['id'] if products else None
        
    def generate_stock_report(self, filename='product_stock_by_category.xlsx',callback=None):
        categories = self.get_categories()
        products = self.get_all_products()
        writer = pd.ExcelWriter(filename, engine='openpyxl')

        for category in categories:
            category_products = [product for product in products if category['id'] in [cat['id'] for cat in product['categories']]]

            if not category_products:
                continue

            product_data = []
            for product in category_products:
                if product['type'] == 'variable':
                    variations = self.get_product_variations(product['id'])
                    for variation in variations:
                        size = next((attribute['option'] for attribute in variation['attributes'] if attribute['name'].lower() == 'size'), 'Unknown')
                        product_data.append({'Product Name': product['name'], 'Size': size, 'Stock': variation['stock_quantity']})

            df = pd.DataFrame(product_data)
            pivot_df = df.pivot_table(index='Product Name', columns='Size', values='Stock', fill_value=0)
            pivot_df.to_excel(writer, sheet_name=category['name'])

        writer.close()
        print("Reported generated")
        if callback:
            callback()
            

    def update_stock_from_excel(self, file_path, callback=None, update_gui=None):
        def update_stock(product_id, variation_id, new_stock):
            # Get the current stock quantity
            old_stock = self.wcapi.get(f"products/{product_id}/variations/{variation_id}").json().get('stock_quantity', 'Unknown')
            
            # Update the stock quantity
            data = {"stock_quantity": new_stock}
            response = self.wcapi.put(f"products/{product_id}/variations/{variation_id}", data).json()
            
            # Return old stock and response for GUI update
            return old_stock, response.get('stock_quantity', 'Unknown')

        xl = pd.ExcelFile(file_path)

        for sheet_name in xl.sheet_names:
            df = xl.parse(sheet_name)
            for index, row in df.iterrows():
                product_name = row['Product Name']
                product_id = self.get_product_id(product_name)
                if product_id is None:
                    continue

                for size, new_stock in row.items():
                    if size == 'Product Name':
                        continue

                    variation_id = self.get_variation_id(product_id, size)
                    if variation_id:
                        old_stock, updated_stock = update_stock(product_id, variation_id, new_stock)
                        if update_gui:
                            update_gui(f"Product '{product_name}' Size: '{size}' Stock updated from {old_stock} to {updated_stock}")
                    else:
                        if update_gui:
                            update_gui(f"Variation ID not found for size '{size}' in product '{product_name}'")

        if callback:
            callback()

    
    def shipping_label_update(self, filepath, output_filepath = "shipping_out.pdf", callback=None):
                
        # Convert the PDF to a list of images
        images = pdf2image.convert_from_path(filepath,dpi=300)

        # Create a new PDF
        #pdf = FPDF()
        # Create a new PDF with 4x6 inches page size
        pdf = FPDF(unit='mm', format=(4*25.4, 6*25.4))

        # Loop over the images (pages)
        for i, image in enumerate(images):
            # Convert the image to RGB
            image = image.convert('RGB')

            # Create a draw object
            draw = ImageDraw.Draw(image)

            # Extract the order number from the text in the image
            text = pytesseract.image_to_string(image)
            print(text)
            match = re.search(r'Customer Ref:\s*[\d\s]*\/\s*#(\d+)', text)
            if match:
                order_number = match.group(1)  # The order number is in the first group
            else:
                print(f"Order number not found on page {i+1}")
                continue  # Skip to the next page

            # Get order details
            #order_number = "2231"
            data = self.wcapi.get("orders/" + order_number).json()

            # Extract product names and quantities
            line_items = data['line_items']
            order_list = [(item['name'], item['quantity']) for item in line_items]

            # Prepare the custom text
            custom_text = "\n".join(product for product, quantity in order_list for _ in range(quantity))

            # Draw the custom text at the bottom left corner
            font = ImageFont.truetype("arial.ttf", size=30)
            draw.text((86, image.height - 340), custom_text, fill="black",font=font)

            # Define the position and size of the rectangle: x1, y1, x2, y2
            x1 = 900
            y1 = image.height - 925
            x2 = x1 + 225  # Width of the rectangle
            y2 = y1 + 380  # Height of the rectangle

            # Draw a white rectangle with a black border
            draw.rectangle([(x1, y1), (x2, y2)], fill="white")

            # Save the image to a temporary file in PNG format
            image_path = f"temp_page_{i+1}.png"
            image.save(image_path)

            # Add the image to the PDF
            pdf.add_page()
            pdf.image(image_path, x = 0, y = 0, w = pdf.w, h = pdf.h)

        # Save the final PDF
        pdf.output(output_filepath) 
          
        if callback:
            callback()
        




   
