import os
import tkinter as tk
from tkinter import ttk, messagebox, font, Canvas, Scrollbar, Frame, VERTICAL, S
from docxtpl import DocxTemplate
import pandas as pd
from docx2pdf import convert

CLIENTS_SHEET_NAME = 1

class InvoiceApp:

    def __init__(self, master):
        self.master = master
        master.title("Invoice App")

        self.label = tk.Label(master, text="Invoices - Select the one you want to generate:")
        self.label.pack(padx=10, pady=10)

        frame = Frame(master)
        frame.pack(padx=10, pady=10)

        self.scrollbar = Scrollbar(frame)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(frame, selectmode='extended', yscrollcommand=self.scrollbar.set)
        self.tree["columns"]=("date", "day_of_week", "name", "total", "notes")
        self.tree.column("#0", width=100, minwidth=50, stretch=tk.NO)
        self.tree.column("date", width=100, minwidth=50, stretch=tk.NO)
        self.tree.column("day_of_week", width=100, minwidth=50, stretch=tk.NO)
        self.tree.column("name", width=160, minwidth=100)
        self.tree.column("total", width=100, minwidth=50, stretch=tk.NO)
        self.tree.column("notes", width=400, minwidth=200)

        self.tree.heading("#0",text="Invoice Number",anchor=tk.W)
        self.tree.heading("date", text="Date",anchor=tk.W)
        self.tree.heading("day_of_week", text="Day of Week",anchor=tk.W)
        self.tree.heading("name", text="Name",anchor=tk.W)
        self.tree.heading("total", text="Total",anchor=tk.W)
        self.tree.heading("notes", text="Notes",anchor=tk.W)

        self.scrollbar.config(command=self.tree.yview)

        excel_path = "invoice_data.xlsx"
        self.df_invoice = pd.read_excel(excel_path).set_index('invoice_number')
        self.df_client = pd.read_excel(excel_path, sheet_name=CLIENTS_SHEET_NAME).set_index('id')

        self.populate_treeview()
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH)

        self.submit_button = tk.Button(master, text="Generate Invoice", command=self.main)
        self.submit_button.pack(side=tk.BOTTOM, padx=10, pady=10)

    def populate_treeview(self):
        for invoice_number in self.df_invoice.index:
            invoice_data = self.df_invoice.loc[invoice_number]
            total = round(invoice_data["amunt"] * invoice_data["price_per"], 2)
            self.tree.insert('', 'end', text=str(invoice_number), values=(invoice_data['date'].strftime('%d-%m-%Y'), invoice_data['date'].day_name(), invoice_data['name'], total, invoice_data['notes']))

    def main(self):
        try:
            selected_invoices = [self.tree.item(item)['text'] for item in self.tree.selection()]
            if not os.path.exists('invoices'):
                os.makedirs('invoices')
            for invoice_number in selected_invoices:
                invoice_data = self.get_invoice(invoice_number)
                client_data = self.get_client(invoice_data["client_id"])

                context = self.create_context(invoice_data, client_data)

                doc = DocxTemplate('invoice_template.docx')
                doc.render(context)

                filename = f'invoices/{invoice_data.name}.docx'
                doc.save(filename)
                convert(filename)
            messagebox.showinfo("Success", "Invoice(s) generated successfully!")
        except ValueError as e:
            messagebox.showerror("Error", str(e))

    def get_invoice(self, invoice_number: str) -> pd.Series:
        if invoice_number not in self.df_invoice.index:
            raise ValueError(f"Invoice number {invoice_number} not found.")
        return self.df_invoice.loc[invoice_number]

    def get_client(self, client_id: str) -> pd.Series:
        if client_id not in self.df_client.index:
            raise ValueError(f"Client id {client_id} not found.")
        return self.df_client.loc[client_id]

    def create_context(self, invoice_data: pd.Series, client_data: pd.Series) -> dict:
        formatted_date = invoice_data["date"].date().strftime('%d-%m-%Y')
        price_per = round(invoice_data["price_per"], 2)
        price_total = round(invoice_data["amunt"] * invoice_data["price_per"], 2)
        return {
            #Invoice data - Customer Data
            "name": client_data["name"],
            "identification": client_data["identification"],
            "address": client_data["addreess"],
            "postal_code": client_data["postal_code"],
            "city": client_data["city"],
            "country": client_data["country"],
            #Invoice data - Invoice Data
            "invoice_number": invoice_data.name,
            "date": formatted_date,
            "notes": invoice_data["notes"],
            "item_description": invoice_data["item_description"],
            "amunt": invoice_data["amunt"],
            "price_per": price_per,
            "currency": invoice_data["currency"],
            "price_total": price_total
        }

root = tk.Tk()
my_gui = InvoiceApp(root)
root.mainloop()