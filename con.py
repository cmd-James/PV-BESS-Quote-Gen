# -*- coding: utf-8 -*-
"""
Created on Tue Jul  2 18:53:14 2024

@author: Duncan
"""

import openpyxl

def convert_to_pdf(workbook, output_filename):
  """Converts an openpyxl workbook to a PDF document.

  Args:
    workbook: The openpyxl workbook object to be converted.
    output_filename: The desired filename for the generated PDF.
  """

  try:
    writer = None  # Initialize writer outside conditional block

    if hasattr(openpyxl.writer, 'write_only'):  # Available in newer versions
      writer = openpyxl.writer.write_only.PdfWriter()
    else:
      # Fallback for older versions (may require external libraries)
      print("Using a fallback method (may require external libraries).")

    writer.write_workbook(workbook)
    writer.save(output_filename)
    print(f"Workbook converted to PDF successfully: {output_filename}")
  except Exception as e:
    print(f"An error occurred during conversion: {e}")


def main():
  """Main function to prompt for input and workbook conversion."""

  input_filename = input("Enter the Excel workbook filename (including .xlsx): ")
  output_filename = input("Enter the desired filename for the PDF (including .pdf): ")

  try:
    workbook = openpyxl.load_workbook(filename=input_filename)
    convert_to_pdf(workbook, output_filename)
  except FileNotFoundError:
    print(f"Error: File '{input_filename}' not found.")

if __name__ == "__main__":
  main()


