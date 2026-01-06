import pdfplumber
import pandas as pd
import re
import os


def extract_data(pdf_path):
    data = {
        "billing_address": "", "shipping_address": "", "invoice_type": "",
        "order_number": "", "invoice_number": "", "order_date": "",
        "invoice_details": "", "invoice_date": "", "seller_info": "",
        "seller_pan": "", "seller_gst": "", "fssai_license": "Not Available",
        "billing_state_code": "", "shipping_state_code": "", "place_of_supply": "",
        "place_of_delivery": "", "reverse_charge": "No", "amount_in_words": "",
        "seller_name": "", "seller_address": "", "total_tax": "", "total_amount": ""
    }

    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
        lines = [line.strip() for line in text.split('\n')]
                
        # Simple Address & Seller Logic 
        for i, line in enumerate(lines):
            if "Billing Address" in line:
                addr = ", ".join(lines[i+1:i+4]).replace("*", "")
                data["billing_address"] = addr
                data["shipping_address"] = addr 
            
            if "Sold By" in line:
                data["seller_name"] = lines[i+1]
                raw_addr = ", ".join(lines[i+2:i+5]).replace("*", "")
                data["seller_address"] = raw_addr
                data["seller_info"] = f"{data['seller_name']}, {raw_addr}, India"
        
        data["invoice_type"] = "Tax Invoice" if "Tax Invoice" in text else "Bill of Supply"
        data["order_number"] = (re.search(r"Order (?:Number|Id|ID)\s*[:]?\s*(\S+)", text)).group(1) if re.search(r"Order (?:Number|Id|ID)", text) else ""
        data["invoice_number"] = (re.search(r"Invoice (?:Number|No)\s*[:]?\s*(\S+)", text)).group(1) if re.search(r"Invoice (?:Number|No)", text) else ""
        data["order_date"] = (re.search(r"Order Date\s*[:]?\s*([\d\.-]+)", text)).group(1) if "Order Date" in text else ""
        data["invoice_details"] = (re.search(r"Invoice Details\s*[:]?\s*(\S+)", text)).group(1) if "Invoice Details" in text else None
        data["invoice_date"] = (re.search(r"Invoice Date\s*[:]?\s*([\d\.-]+)", text)).group(1) if "Invoice Date" in text else ""
        data["seller_gst"] = (re.search(r"GST(?:IN| Registration No)?\s*[:]?\s*([0-9A-Z]{15})", text)).group(1) if "GST" in text else ""
        data["seller_pan"] = (re.search(r"PAN(?: No)?\s*[:]?\s*([A-Z]{5}[0-9]{4}[A-Z]{1})", text)).group(1) if "PAN" in text else ""
        fssai = re.search(r"FSSAI License No\.\s*(\d+)", text)
        data["fssai_license"] = re.search(r"FSSAI License No\.\s*(\d+)", text).group(1) if fssai else "Not Available"
        state_code_match = re.search(r"State/UT Code\s*[:]?\s*(\d+)", text)
        data['billing_state_code'] = state_code_match.group(1) if state_code_match else "N/A"
        data['shipping_state_code'] = data['billing_state_code']
        
        pos_match = re.search(r"Place of supply\s*[:]?\s*(.*)", text, re.IGNORECASE)
        pod_match = re.search(r"Place of delivery\s*[:]?\s*(.*)", text, re.IGNORECASE)
        supply_val = pos_match.group(1).strip() if pos_match else ""
        delivery_val = pod_match.group(1).strip() if pod_match else ""

        # Flipkart
        if not supply_val:
            state_pattern = re.search(r",\s*IN-([A-Z]{2})", text)
            if state_pattern:
                supply_val = state_pattern.group(1).strip()
                delivery_val = supply_val
            else:
                supply_val = data.get("billing_state_code", "")
                delivery_val = supply_val

        data["place_of_supply"] = supply_val
        data["place_of_delivery"] = delivery_val

        words_match = re.search(r"Amount in Words\s*[:]?\s*(.*)", text, re.IGNORECASE)

        if words_match:
            data["amount_in_words"] = words_match.group(1).strip()
        else:
            fallback_words = re.search(r"([A-Za-z\s\-]+(?:only|Only))", text)
            data["amount_in_words"] = fallback_words.group(1).strip() if fallback_words else ""
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if any("TOTAL" in str(c).upper() for c in row if c):
                        data["total_amount"] = row[-1].replace('₹', '').strip()
                        data["total_tax"] = row[-2] if len(row) > 1 else ""

    return data

#  files
files = ["invoice1.pdf", "invoice2.pdf", "invoice3.pdf", "invvoice4.pdf"]
results = [extract_data(f) for f in files if os.path.exists(f)]

# Export Excel
df = pd.DataFrame(results).T
df.to_excel("Final_Output.xlsx")
print("Done! Check Final_Output.xlsx")