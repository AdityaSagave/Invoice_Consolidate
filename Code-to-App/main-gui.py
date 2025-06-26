# Pandas Library for processing excel sheets
import pandas as pd

# Tkinter libraries for desktop app
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import os
from datetime import date

#REPORTlAB LIBRARIES FOR PDF INVOICES
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet

import tkinter.font as tkFont
import pygame

import datetime



selected_files = []

def refresh_file_list():
    for widget in file_list_frame.winfo_children():
        widget.destroy()

    for i, file_path in enumerate(selected_files):
        filename = os.path.basename(file_path)

        row_frame = ttk.Frame(file_list_frame, style="File.TFrame")
        row_frame.grid(row=i, column=0, sticky="ew", pady=4, padx=5)

        label = ttk.Label(row_frame, text=filename, anchor='w', style="File.TLabel")
        label.pack(side="left", fill="x", expand=True)

        remove_btn = ttk.Button(row_frame, text="Remove", style="Accent.TButton", command=lambda idx=i: remove_file(idx))
        remove_btn.pack(side="right", padx=5)

def add_file():
    file_paths = filedialog.askopenfilenames(
        title="Select Excel Invoice Files",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_paths:
        selected_files.extend(file_paths)
        refresh_file_list()

def remove_file(index):
    if 0 <= index < len(selected_files):
        selected_files.pop(index)
        refresh_file_list()

def clear_files():
    selected_files.clear()
    refresh_file_list()


def generate_invoice_pdf(df, output_path, invoice_month="May 2025"):
    doc = SimpleDocTemplate(output_path, pagesize=A4,
                            rightMargin=inch, leftMargin=inch,
                            topMargin=inch, bottomMargin=inch)
    
    styles = getSampleStyleSheet()
    story = []
    
    # Left content: company info
    left_paragraphs = [
        Paragraph("<b>Landauer Inc.</b>", styles['Normal']),
        Paragraph("2 Science Road", styles['Normal']),
        Paragraph("Glenwood, IL 60425", styles['Normal'])
    ]
    right_aligned_style = ParagraphStyle(
    name='RightAlign',
    parent=styles['Normal'],
    alignment=2,  # 2 = TA_RIGHT
    fontSize=10,
    )

    right_paragraphs = [
        Paragraph(f"Date: {date.today().strftime('%B %d, %Y')}", right_aligned_style),
        Paragraph(f"For the Month of: {invoice_month}", right_aligned_style),
    ]


    # Build the two-column table with left and right blocks
    header_table = Table(
        [[left_paragraphs, right_paragraphs]],
        colWidths=[3 * inch, 3.3 * inch],  # Adjust as needed
    )

    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
    ]))

    story.append(Paragraph("Landauer Australasia Pty Ltd", styles['Title']))
    story.append(Paragraph("Internal LDR Invoicing", styles['Title']))
    story.append(header_table)
    story.append(Spacer(1, 24))

    
    # Paragraph style for wrapped text inside table cells
    wrap_style = ParagraphStyle(
        name='WrapStyle',
        fontName='Helvetica',
        fontSize=8,
        leading=10,
        alignment=0,  # left align
    )
    
    # Columns you want to wrap and not wrap
    wrap_columns = ['AccountName', 'Badge Type']
    no_wrap_columns = ['Invoice Number']
    
    raw_data = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
    
    data = []
    for row_idx, row in enumerate(raw_data):
        new_row = []
        for col_idx, cell in enumerate(row):
            col_name = df.columns[col_idx] if row_idx > 0 else raw_data[0][col_idx]
            if row_idx == 0:
                # Header bold, no wrap needed
                new_row.append(Paragraph(f"<b>{cell}</b>", wrap_style))
            else:
                if col_name in no_wrap_columns:
                    new_row.append(cell)  # Plain text, no wrap
                else:
                    new_row.append(Paragraph(cell, wrap_style))  # Wrapped paragraph
        data.append(new_row)
    
    # Calculate column widths to fit page width nicely
    page_width, _ = A4
    available_width = page_width - 2 * inch  # margins
    num_cols = len(df.columns)
    col_width = available_width / num_cols
    col_widths = [col_width] * num_cols
    
    table = Table(data, colWidths=col_widths, repeatRows=1)
    
    table_style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),  # Important for wrapped paragraphs
    ])
    
    table.setStyle(table_style)
    story.append(table)
    doc.build(story)

def process_files():
    if not selected_files:
        messagebox.showwarning("No Files", "Please add at least one Excel file.")
        return

    try:
        all_grouped_data = []
        unique_account_names = set()
        unique_InvoiceNumbers = set()

        for file_path in selected_files:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)

                # Get columns up to 'Ext Price'
                cols_to_use = ['Account Number', 'AccountName', 'Frequency', 'Invoice Number', 'Invoice Date', 'Badge Type', 'Quantity', 'Unit Price', 'Ext Price']
                df_subset = df[cols_to_use]

                # Handle missing Frequency values
                df_subset['Frequency'] = df_subset['Frequency'].fillna('Missing')
                df_subset['Badge Type'] = df_subset['Badge Type'].fillna('Fees')

                # Convert to datetime in case of mixed types
                df_subset['Invoice Date'] = pd.to_datetime(df_subset['Invoice Date'], errors='coerce').dt.strftime('%d %B %Y')

                # Group with aggregation that includes non-numeric fields
                grouped = df_subset.groupby(['Account Number'], as_index=False).agg({
                    
                    
                    'Ext Price': 'sum',
                    'Invoice Date': 'first',
                    'Invoice Number': 'first'
                })

                # Merge back AccountName
                account_names = df[['Account Number', 'AccountName']].drop_duplicates()
                grouped = grouped.merge(account_names, on='Account Number', how='left')

                all_grouped_data.append(grouped)

        combined = pd.concat(all_grouped_data, ignore_index=True)
        # Final grouping with full aggregation
        final_grouped = combined.groupby(['Account Number', 'AccountName'], as_index=False).agg({
            
            'Invoice Number': 'first',
            'Invoice Date': 'first',
            
            
            'Ext Price': 'sum'
        })

        unique_account_names.update(df['AccountName'].dropna().unique())
        unique_InvoiceNumbers.update(df['Invoice Number'].dropna().unique())

        # Add the total row
        total_row = {
            'AccountName': f"TOTAL | Unique Accounts: {len(unique_account_names)}",
            'Account Number': '',
            
            'Invoice Number': '',
            'Invoice Date': '',
            
            
            'Ext Price': final_grouped['Ext Price'].sum()
        }

        # Append total row
        final_grouped_with_total = pd.concat([final_grouped, pd.DataFrame([total_row])], ignore_index=True)

        # Reorder columns: AccountName first
        desired_order = ['AccountName', 'Account Number', 'Invoice Number', 'Invoice Date', 'Ext Price']

        final_grouped_with_total = final_grouped_with_total[desired_order]
        # Round and format Ext Price with dollar sign
        final_grouped_with_total['Ext Price'] = final_grouped_with_total['Ext Price'].apply(
            lambda x: f"${x:,.2f}" if pd.notna(x) and str(x).strip() != '' else ''
        )

        output_pdf_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save summary as..."
        )
        if output_pdf_path:
            # Get the most common month-year combo from the 'Invoice Date' column
            invoice_dates = pd.to_datetime(final_grouped_with_total['Invoice Date'], errors='coerce')
            invoice_dates = invoice_dates.dropna()

            if not invoice_dates.empty:
                # Get most frequent month and year
                most_common_month_year = invoice_dates.dt.to_period('M').mode()[0]
                invoice_month_str = most_common_month_year.strftime('%B %Y')  # e.g., "May 2025"
            else:
                invoice_month_str = "Unknown"
            generate_invoice_pdf(final_grouped_with_total, output_pdf_path, invoice_month=invoice_month_str)
            messagebox.showinfo("Success", f"Processing complete!\nFile saved as:\n{output_pdf_path}")
        else:
            messagebox.showinfo("Cancelled", "Save operation cancelled.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

# ----- GUI Setup -----
root = tk.Tk()
root.title("Invoice Consolidator")
def center_window(root, width=820, height=839):
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    root.geometry(f"{width}x{height}+{x}+{y}")

# Replace these two lines:
# root.geometry("920x939")
# center_window(root) instead
center_window(root)
# root.minsize(910, 939)
root.configure(bg="#fef9f2")

# ---- Styling ----
style = ttk.Style()
style.theme_use("clam")

style.configure("TLabel", font=('Segoe UI', 10), background="#fef9f2")
style.configure("TFrame", background="#fef9f2")
style.configure("File.TFrame", background="#f7f3ea")
style.configure("File.TLabel", background="#f7f3ea", font=('Segoe UI', 10))

style.configure("TButton",
                font=('Segoe UI', 10),
                padding=6,
                relief="flat",
                background="#fef9f2")

style.map("TButton",
          background=[("active", "#dceeff")])

style.configure("Accent.TButton",
                background="#007acc",
                foreground="white",
                padding=6,
                font=('Segoe UI', 10, 'bold'))

style.map("Accent.TButton",
          background=[("active", "#005f9e")],
          foreground=[("disabled", "gray")])

# ---- Top Label ----
ttk.Label(root, text="Selected Excel Invoice Files:", font=('Segoe UI', 11, 'bold')).pack(anchor="w", padx=12, pady=(12, 0))

# ---- Scrollable File List Frame ----
canvas_frame = ttk.Frame(root)
canvas_frame.pack(fill='both', expand=True, padx=12, pady=6)

tk_canvas = tk.Canvas(canvas_frame, bg="#fef9f2", highlightthickness=0)
scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=tk_canvas.yview)
tk_canvas.configure(yscrollcommand=scrollbar.set)

scrollbar.pack(side="right", fill="y")
tk_canvas.pack(side="left", fill="both", expand=True)

file_list_frame = ttk.Frame(tk_canvas)
tk_canvas.create_window((0, 0), window=file_list_frame, anchor='nw')

def on_frame_configure(event):
    tk_canvas.configure(scrollregion=tk_canvas.bbox("all"))

file_list_frame.bind("<Configure>", on_frame_configure)

# ---- Buttons at Bottom ----
button_frame = ttk.Frame(root)
button_frame.pack(fill='x', pady=10, padx=12)

ttk.Button(button_frame, text="Add File", command=add_file, style="Accent.TButton").pack(side="left", padx=6)
ttk.Button(button_frame, text="Clear All", command=clear_files).pack(side="left", padx=6)
ttk.Button(button_frame, text="Process Files", command=process_files, style="Accent.TButton").pack(side="right", padx=6)

root.mainloop()
