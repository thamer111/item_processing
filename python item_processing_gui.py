import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Initialize global variable for the input file path
input_file_path = None

# Function to import the file
def import_file():
    global input_file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_label.config(text=f" {file_path} : تم استيراد ")
        input_file_path = file_path

# Function to execute the processing
def execute_processing():
    global input_file_path
    if not input_file_path:
        messagebox.showerror("خطأ", "لم يتم اختيار ملف")
        return

    try:
        # Load the provided files
        item_inventory_file = pd.read_excel('Item inventory file.xlsx', header=None, dtype=str)  # Read inventory file
        input_file = pd.read_excel(input_file_path, header=0, dtype=str)  # Read imported file

        # Debug: Print original data
        print("Original Data (Before Cleaning):")
        print(input_file.head(10))

        # Extract data for item Processing
        unique_items = input_file.iloc[:, [1, 2, 7]]
        unique_items.columns = ['رقم الصنف', 'اسم الصنف', 'الكمية']

        # Clean the data for unique items
        unique_items['رقم الصنف'] = unique_items['رقم الصنف'].str.strip()
        unique_items['الكمية'] = pd.to_numeric(unique_items['الكمية'], errors='coerce').fillna(0)

        # Debug: Check cleaned data
        print("Data After Cleaning (item Processing):")
        print(unique_items.head(10))

        # Read inventory data
        existing_item_numbers = (
            item_inventory_file.iloc[1:, 2]  # Column C
            .dropna()
            .str.strip()
            .tolist()
        )
        print("Existing Inventory Item Numbers:")
        print(existing_item_numbers)

        # Filter out existing items for item Processing
        unique_items = unique_items[~unique_items['رقم الصنف'].isin(existing_item_numbers)]

        # Debug: Check filtered data
        print("Data After Filtering (item Processing):")
        print(unique_items.head(10))

        # Group and summarize for item Processing
        unique_items = unique_items.groupby(['رقم الصنف', 'اسم الصنف'], as_index=False).agg({
            'الكمية': 'sum'
        })
        unique_items['رقم المجموعة'] = 1
        unique_items['الوحدة'] = 'حبة'
        unique_items['التخزين'] = 1
        unique_items['رقم الوظيفة'] = 1

        # Rearrange columns for item Processing
        unique_items = unique_items[[
            'رقم المجموعة', 'رقم الصنف', 'اسم الصنف', 'الوحدة', 'التخزين', 'رقم الوظيفة', 'الكمية'
        ]]

        # Save item Processing
        with pd.ExcelWriter('item Processing.xlsx', engine='xlsxwriter') as writer:
            unique_items.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            format_text = writer.book.add_format({'num_format': '@'})  # Text format for item number
            worksheet.set_column('B:B', None, format_text)

        # Extract data for Invoice Processing
        purchase_invoice_form = input_file.iloc[:, [1, 2, 3, 5, 8, 7]]
        purchase_invoice_form.columns = ['رقم الصنف', 'اسم الصنف', 'الوحدة', 'التعبئة', 'السعر', 'الكمية']

        # Clean the data for purchase invoice
        purchase_invoice_form['رقم الصنف'] = purchase_invoice_form['رقم الصنف'].str.strip()
        purchase_invoice_form['الكمية'] = pd.to_numeric(purchase_invoice_form['الكمية'], errors='coerce').fillna(0)
        purchase_invoice_form['السعر'] = pd.to_numeric(purchase_invoice_form['السعر'], errors='coerce').fillna(0)

        # Debug: Check cleaned data
        print("Data After Cleaning (Invoice Processing):")
        print(purchase_invoice_form.head(10))

        # Group and summarize for Invoice Processing
        purchase_invoice_form = purchase_invoice_form.groupby(
            ['رقم الصنف', 'اسم الصنف', 'الوحدة', 'التعبئة', 'السعر'],
            as_index=False
        ).agg({'الكمية': 'sum'})

        # Save Invoice Processing
        with pd.ExcelWriter('Invoice Processing.xlsx', engine='xlsxwriter') as writer:
            purchase_invoice_form.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            worksheet.set_column('B:B', None, format_text)

        # Create نموذج بيانات الأصناف
        item_data_model = unique_items.copy()
        item_data_model['الإسم الاجنبي'] = ''
        item_data_model['عبوة الصنف'] = 1
        item_data_model['وحدة رئيسية'] = 1
        item_data_model['نوع الصنف'] = ''
        item_data_model['التكلفة'] = ''
        item_data_model['وحدة البيع'] = ''
        item_data_model['الكسور'] = 1
        item_data_model['الفرعية'] = 1
        item_data_model['تحت الفرعية'] = ''
        item_data_model['التصنيف'] = ''
        item_data_model['وحدة شراء'] = 1
        item_data_model['وحدة جرد'] = 1
        item_data_model['الباركود'] = ''

        # Rearrange columns for نموذج بيانات الأصناف
        item_data_model = item_data_model[[
            'رقم المجموعة', 'رقم الصنف', 'اسم الصنف', 'الإسم الاجنبي', 'الوحدة', 'الباركود', 'عبوة الصنف',
            'وحدة رئيسية', 'نوع الصنف', 'التكلفة', 'وحدة البيع', 'الكسور', 'الفرعية',
            'تحت الفرعية', 'التصنيف', 'وحدة شراء', 'وحدة جرد'
        ]]

        # Save نموذج بيانات الأصناف
        with pd.ExcelWriter('نموذج بيانات الأصناف.xlsx', engine='xlsxwriter') as writer:
            item_data_model.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            worksheet.set_column('B:B', None, format_text)

        # Create نموذج إدخال فاتورة الشراء
        purchase_invoice_entry_form = purchase_invoice_form.copy()
        purchase_invoice_entry_form['العبوة'] = 1

        # Rearrange columns for نموذج إدخال فاتورة الشراء
        purchase_invoice_entry_form = purchase_invoice_entry_form[[
            'رقم الصنف', 'اسم الصنف', 'الوحدة', 'العبوة', 'الكمية', 'السعر'
        ]]

        # Save نموذج إدخال فاتورة الشراء
        with pd.ExcelWriter('نموذج إدخال فاتورة الشراء.xlsx', engine='xlsxwriter') as writer:
            purchase_invoice_entry_form.to_excel(writer, index=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            worksheet.set_column('B:B', None, format_text)

        messagebox.showinfo("نجاح", "تمت المعالجة بنجاح. تم حفظ الملفات.")

    except Exception as e:
        messagebox.showerror("خطأ", str(e))

# Create the main window
root = tk.Tk()
root.title("أداة معالجة الأصناف")
root.geometry("500x450")

# Fonts for styling
font_large = ("Arial", 14, "bold")
font_medium = ("Arial", 12)
font_small = ("Arial", 10)

# Title
title_label = tk.Label(root, text="أداة معالجة الأصناف", font=font_large)
title_label.pack(pady=10)

# Description
description_label = tk.Label(root, text="Excel هذا البرنامج يعالج بيانات الأصناف والمشتريات من ملفات\n"
                                        " : ينتج تقارير\n"
                                        " item Processing- \n"
                                        "Invoice Processing-\n"
                                        "نموذج بيانات الأصناف-\n"
                                        "نموذج إدخال فاتورة الشراء-",
                             font=font_small, justify="right")
description_label.pack(pady=5)

# Requirements
requirements_label = tk.Label(root, text=":متطلبات الملف المدخل\n"
                                         ": يجب أن يحتوي على الأعمدة التالية -\n"
                                         "  [رقم الصنف، اسم الصنف، الوحدة، التعبئة، السعر، الكمية]\n"
                                         ".xlsx يجب أن يكون الملف بصيغة -",
                               font=font_small, justify="right")
requirements_label.pack(pady=5)

# Import Button
import_button = tk.Button(root, text="استيراد ملف", font=font_medium, command=import_file)
import_button.pack(pady=10)

# File Status Label
file_label = tk.Label(root, text="لم يتم استيراد ملف", font=font_medium)
file_label.pack(pady=10)

# Execute Button
execute_button = tk.Button(root, text="تنفيذ", font=font_medium, command=execute_processing)
execute_button.pack(pady=10)

# Digital Signature
signature_label = tk.Label(root, text="Created by Thamer Ali © 2025", font=("Arial", 9, "italic"), fg="gray") 
signature_label.pack(side="bottom", pady=10)

if __name__ == "__main__":
    root.mainloop()
