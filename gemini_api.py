import google.generativeai as genai
import os
import pandas as pd
import re
import json
import time
import openpyxl

client = genai.configure(api_key="AIzaSyCpujkcZAcvudbJlxGkHSmVXc5hePVxuSE")
# model = genai.GenerativeModel('gemini-2.0-flash')
model = genai.GenerativeModel('gemini-2.5-pro-preview-06-05')
# model = genai.GenerativeModel('gemini-1.5-flash')
# model = genai.GenerativeModel('gemini-1.5-pro')

# Upload file to Gemini
def upload_file_to_gemini(path, display_name=None):
    print(f"Uploading '{path}'")
    try:
        uploaded_file = genai.upload_file(path=path, display_name=display_name or os.path.basename(path))
        print(f"Upload complete. File URI: {uploaded_file.uri}")
        return uploaded_file
    except Exception as e:
        print(f"Error uploading '{path}': {e}")
        return None

def delete_file_from_gemini(file_name):
    try:
        genai.delete_file(file_name)
        print(f"Deleted uploaded file: {file_name}")
    except Exception as e:
        print(f"Error deleting file '{file_name}': {e}")

def delete_all_uploaded_files():
    print("\nDeleting all previously uploaded files...")
    try:
        files = genai.list_files()
        for f in files:
            delete_file_from_gemini(f.name)
    except Exception as e:
        print(f"Error during cleanup: {e}")

def parse_and_save_to_excel(response_text, output_file="output.xlsx"):
    print("\nParsing and appending to Excel")

    if "```json" in response_text:
        response_text = response_text.split("```json")[-1]
    if "```" in response_text:
        response_text = response_text.split("```")[0]
    response_text = response_text.strip()

    try:
        receipts = json.loads(response_text)
        if not isinstance(receipts, list):
            receipts = [receipts]

        all_dataframes = []
        for receipt in receipts:
            company = receipt.get("company_name", "")
            date = receipt.get("date", "")
            total_before_tax = receipt.get("total_before_tax", "")
            taxes = receipt.get("taxes", "")
            total_after_tax = receipt.get("total_after_tax", "")
            items = receipt.get("items", [])

            filtered_items = []
            for item in items:
                qty = item.get("quantity")
                try:
                    if qty is not None and float(qty) != 0:
                        filtered_items.append(item)
                except (ValueError, TypeError):
                    continue  

            if not filtered_items:
                continue  

            df = pd.DataFrame(filtered_items)
            df.insert(0, "Company Name", company)
            df.insert(1, "Date", date)
            df["Total Before Tax"] = total_before_tax
            df["Tax"] = taxes
            df["Total After Tax"] = total_after_tax
            all_dataframes.append(df)

        if not all_dataframes:
            print("No valid data to save.")
            return

        final_df = pd.concat(all_dataframes, ignore_index=True)

        if os.path.exists(output_file):
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                book = writer.book
                sheet = writer.sheets.get("Sheet1") or book.active
                startrow = sheet.max_row
                final_df.to_excel(writer, sheet_name=sheet.title, startrow=startrow, index=False, header=False)
                print(f"Data appended to existing file: {output_file}")
        else:
            final_df.to_excel(output_file, index=False)
            print(f"New Excel file created: {output_file}")

    except Exception as e:
        print(f"JSON error: {e}")

def process_single_document(delete_after=True):
    print("\nSingle Document Processing")
    # file_path = r"C:\Y3S1\Details Extraction from Receipts & Invoices\invoice.pdf"
    # file_path = r"C:\Y3S1\Details Extraction from Receipts & Invoices\scan receipt sample 1.pdf"
    file_path = r"C:\Y3S1\Details Extraction from Receipts & Invoices\scan receipt sample 2.pdf"
    # file_path = r"C:\Y3S1\Details Extraction from Receipts & Invoices\SGP Invoice 9926910986.PDF"
    # file_path = r"C:\Y3S1\Details Extraction from Receipts & Invoices\WhatsApp Image 2025-02-22 at 11.53.41.jpeg"

    if not os.path.exists(file_path):
        print(f"File not found at '{file_path}'.")
        return

    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension not in ['.pdf', '.jpeg', '.jpg', '.png', '.webp']:
        print(f"Unsupported file type '{file_extension}'.")
        return

    document_file = upload_file_to_gemini(file_path, display_name=f"User Document: {os.path.basename(file_path)}")
    if not document_file:
        print("Failed to upload document.")
        return
    time.sleep(1)

    prompt_parts = [
        f"The file '{os.path.basename(file_path)}' may contain one or more receipts or invoices.",
        "Your task is to extract detailed structured data from each receipt or invoice found in the document.",
        "",
        "For **each** receipt or invoice, extract the following fields:",
        "- `company_name`: The name of the business issuing the receipt or invoice.",
        "- `date`: The date the receipt or invoice was issued. Format it as `YYYY-MM-DD`.",
        "- `items`: A list of purchased items, each with the following:",
        "    - `description`: Name or description of the item",
        "    - `quantity`: Number of units",
        "    - `unit_price`: Price per unit",
        "    - `total_price`: Total price for this item (as written, do not calculate)",
        "- `taxes`: Any applicable taxes such as GST (if present)",
        "- `total_before_tax`: Subtotal before tax",
        "- `total_after_tax`: Total after including tax",
        "",
        "Return the results as a **JSON array** of receipt objects. Each object should follow this structure:",
        "",
        "```json",
        "{",
        '  "company_name": "string",',
        '  "date": "YYYY-MM-DD",',
        '  "total_before_tax": number_or_string,',
        '  "taxes": number_or_string,',
        '  "total_after_tax": number_or_string,',
        '  "items": [',
        "    {",
        '      "description": "string",',
        '      "quantity": number_or_string,',
        '      "unit_price": number_or_string,',
        '      "total_price": number_or_string',
        "    }, ...",
        "  ]",
        "}",
        "```",
        "",
        "Avoid performing calculations. Just extract the values as they appear in the document.",
        "Those without quanties and does not look like an item should not be rendered as an item. Like shipping charges, etc.",
        "",
        "After extraction, **review the document and your response again** to ensure:",
        "- There are no missing or misformatted values.",
        "If needed, revise and improve the output before returning the final result.",
        "",
        document_file
    ]

    print("\nSending to Gemini")
    try:
        response = model.generate_content(prompt_parts)
        print("\nGemini's analysis")
        print(response.text)
        parse_and_save_to_excel(response.text, output_file="gemini_output.xlsx")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if delete_after and document_file:
            delete_file_from_gemini(document_file.name)

if __name__ == "__main__":
    process_single_document(delete_after=True)
    # delete_all_uploaded_files()