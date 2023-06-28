#!/usr/bin/env python
# coding: utf-8

# In[552]:


#Import the necessary libraries
import pandas as pd
import matplotlib.pyplot as plt
#! pip install reportlab
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
from reportlab.lib import colors
from datetime import datetime, timedelta
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.acroform import AcroForm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle


# In[553]:


#Define a function to get the previous month and year
def get_previous_month():
    now = datetime.now()
    previous_month = now.month - 1 if now.month > 1 else 12
    previous_year = now.year if now.month > 1 else now.year - 1
    return previous_month, previous_year


# In[554]:


# Set the page size to PowerPoint slide size (10 inches by 7 inches)
slide_size = landscape((10 * inch, 7 * inch))

#Create a PDF file and set the canvas
pdf = canvas.Canvas("governance_report.pdf", pagesize=slide_size)
width, height = slide_size


# In[555]:


#Create the Title Page (Page 1)

# Set the background color to blue
pdf.setFillColor(colors.navy)
pdf.rect(0, 0, width, height, fill=True)
# Set the text color to white
pdf.setFillColor(colors.white)

# Set the font and font size
pdf.setFont("Times-Bold", 20)
pdf.drawCentredString(140, height / 2 + 200, "Northwestern Mutual")

# Set the font and font size
pdf.setFont("Times-Bold", 20)
pdf.drawCentredString(190, height / 2 + 50, "TeamName Support Governance")
pdf.line(50, height / 2 + 43, 350, height / 2 + 43)

pdf.setFont("Times-Roman", 16)
previous_month, previous_year = get_previous_month()
pdf.drawString(50, (height / 2) + 20 , f"Previous Month: {previous_month} {previous_year}")
pdf.drawString(50, (height / 2) , f"Current Year: {datetime.now().year}")

pdf.setFont("Times-Roman", 10)
pdf.drawString(50, height / 2 - 200, "The Northwestern Mutal Life Insurance Company - Milwaukee, WI")


# In[556]:


#Create Page 2
pdf.showPage()
pdf.setFont("Times-Roman", 16)
pdf.setFillColor(colors.navy)
pdf.drawCentredString(140, 450, "Page 2: Incident Metrics")
pdf.line(50, 445, width-50, 445)

# draw a rectangle
pdf.rect(400, 300, 265, 140, stroke=1, fill=0)

# Define the position and size of the input box
input_box_x = 475  # X-coordinate of the input box
input_box_y = 370  # Y-coordinate of the input box
input_box_width = 125  # Width of the input box
input_box_height = 30  # Height of the input box

# Create the text field
pdf.setFont("Times-Roman", 14)
pdf.setFillColor(colors.black)
pdf.drawCentredString(530, 405, "CSAT%")
text_field = pdf.acroForm.textfield(
    name="percentage_input",
    x=input_box_x,
    y=input_box_y,
    width=input_box_width,
    height=input_box_height,
    fillColor=colors.white,
    textColor=colors.navy,
    borderWidth=2,
    borderColor=colors.black,
    forceBorder=True)


# Define the position and size of the input box
input_box_x1 = 475  # X-coordinate of the input box
input_box_y1 = 300  # Y-coordinate of the input box
input_box_width1 = 125  # Width of the input box
input_box_height1 = 30  # Height of the input box

# Create the text field
pdf.setFont("Times-Roman", 14)
pdf.setFillColor(colors.black)
pdf.drawCentredString(530, 335, "Association Rate")
text_field1 = pdf.acroForm.textfield(
    name="percentage_input",
    x=input_box_x1,
    y=input_box_y1,
    width=input_box_width1,
    height=input_box_height1,
    fillColor=colors.white,
    textColor=colors.navy,
    borderWidth=2,
    borderColor=colors.black,
    forceBorder=True)


# Add the text field to the PDF
text_field
text_field1


# In[557]:


import glob
from pathlib import Path

# Define the path to the directory containing the CSV files
directory_path = 'report-dataset-files/'

# Get a list of all CSV files in the directory
csv_files = Path(directory_path).glob('*.csv')

# Create an empty list to store the data from each CSV file
dataframes = []

# Iterate over each CSV file
for file in csv_files:
    # Read the CSV file into a pandas DataFrame
    df = pd.read_csv(file)
    df['file'] = file.stem
    # Append the DataFrame to the list
    dataframes.append(df)


# In[558]:


# Save and close the PDF
#pdf_form.updatePageFormFieldValues(pdf)
pdf.save()

