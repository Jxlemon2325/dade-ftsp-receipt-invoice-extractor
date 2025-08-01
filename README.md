# dade-ftsp-receipt-invoice-extractor

This Python script uses Google’s Gemini AI API to extract structured data from receipt or invoice documents (PDF or image files). It uploads the document to Gemini, sends a detailed prompt to extract key fields, and saves the extracted data into an Excel file.

Features:
Uploads a local receipt/invoice file to Gemini AI.
Extracts company name, date, items purchased (description, quantity, unit price, total price), taxes, and totals.
Parses Gemini's JSON response and appends the data to gemini_output.xlsx.
Supports PDF and common image formats.

Requirements:
Python 3.x
google.generativeai Python SDK
pandas
openpyxl

Usage:
Set your Gemini API key in the script (client = genai.configure(api_key="apikey")). (optional to hide the api key)
Modify the file_path variable to point to your receipt/invoice file.

Note:
Error 500 occurs when the input(pdf or images) are too large. Reduce your input context or temporarily switch to another model (e.g. from Gemini 1.5 Pro to Gemini 1.5 Flash) and see if it works. Or wait a bit and retry your request. 

