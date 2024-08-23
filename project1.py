from pdf2image import convert_from_path
import cv2
import numpy as np
import pytesseract
from PIL import Image
import pandas as pd
import matplotlib.pyplot as plt

# Convert PDF to images
pdf_path = 'Sample Problem.pdf'
images = convert_from_path(pdf_path)

# Image processing and text extraction
def process_page(image, box_num_start):
    boxes = segment_image(image)
    # print(boxes)
    # display_image(boxes[1])
    parsed_data = []
    row = 1
    box_num = box_num_start
    for row_boxes in group_boxes_row_wise(boxes):  
        row_data = []
        # print("Row nummber: ", row)
        row+=1
        for box in row_boxes:
            # print("Box in Row")
            if is_valid_box(box):
                # print("Box Number:", box_num)
                text = extract_text_from_box(box)
                # print(text)
                data = parse_text(text, box_num)
                row_data.append(data)
                box_num+=1
                # break
        parsed_data.extend(row_data)
        # break
    return parsed_data, box_num

def segment_image(image):
    image_cv = np.array(image)
    gray = cv2.cvtColor(image_cv, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)
    contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    boxes = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if w > 50 and h > 50:
            box = image_cv[y:y+h, x:x+w]
            boxes.append((x, y, w, h, box))  # Include coordinates for row-wise grouping
    return boxes

def group_boxes_row_wise(boxes):
    # Group boxes based on y-coordinate (rows) and sort by x-coordinate (columns)
    boxes.sort(key=lambda b: (b[1], b[0]))  # Sort by y (row), then by x (column)
    rows = []
    current_row = []
    previous_y = boxes[0][1]
    for box in boxes:
        x, y, w, h, _ = box
        if y - previous_y > h:  # New row detected
            rows.append(current_row)
            current_row = []
        current_row.append(box)
        previous_y = y
    if current_row:  # Add the last row
        rows.append(current_row)
    return rows

def is_valid_box(box):
    x, y, w, h, box_image = box
    box_image_pil = Image.fromarray(box_image)
    text = pytesseract.image_to_string(box_image_pil)
    # print(text)
    if "DELETED" in text.upper():
        return False
    return True

def extract_text_from_box(box):
    x, y, w, h, box_image = box
    box_image_pil = Image.fromarray(box_image)
    text = pytesseract.image_to_string(box_image_pil)
    return text.strip()





def parse_text(text, box_num):
    lines = text.split('\n')
    data = {
        'Part S.No': box_num,
        'Voter Full Name': '',
        'Relative\'s Name': '',
        'Relation Type': '',
        'Age': '',
        'Gender': '',
        'House No': '',
        'EPIC No': '',
    }
    
    next_line_for = None  # Keeps track of which field to fill with the next line
    
    for line in lines:
        line = line.strip()
        # print("Now line",line)

        if "Father's Name" in line or "Husband's Name" in line or "Others" in line:
            # print("Found, relative", line)
            if ':' in line:
                data['Relative\'s Name'] = line.split(':')[1].strip()
                if 'Father\'s Name' in line:
                    data['Relation Type'] = 'FTHR'
                elif 'Husband\'s Name:' in line:
                    data['Relation Type'] = 'HSBN'
                else:
                    data['Relation Type'] = 'OTHR'
            continue
        elif 'Name' in line:
            # print(line)
            parts = None
            if ':' in line:
                parts = line.split(':')
            elif '*' in line:
                parts = line.split('*')
            elif '=' in line:
                parts = line.split('=')
            elif '+' in line:
                parts = line.split('+')
            else:
                parts = line.split('Name')
            data['Voter Full Name'] = parts[1].strip()
            continue 
            
        if 'House' in line:
            parts = line.split(' ')
            next = False
            for part in parts:
                if ":" in part:
                    next = True
                if next == True:
                    if 'Photo' in part:
                        data['House No'] = '-'
                    else:
                        data['House No'] = part
            continue
                
        if line.startswith('TXK') or line.startswith('FRT'):
            data['EPIC No'] = line.strip()
            continue

        if 'Age' in line or 'Gender' in line:
            parts = None
            if '+' in line:
                parts = line.split('+')
            else:
                parts = line.split(':')
            for part in parts:
                if 'Male' in part or 'Female' in part:
                    data['Gender'] = part
                elif 'Gender' in part:
                    # print(part)
                    data['Age'] = part.split(' ')[1].strip()
            next_line_for = None


    return data

def save_to_excel(data, file_path):
    df = pd.DataFrame(data)
    df.to_excel(file_path, index=False)

all_data = []
next_box = 1
for image in images:
    page, b_n = process_page(image, next_box)
    next_box = b_n + 1
    all_data.extend(page)

save_to_excel(all_data, 'output.xlsx')
