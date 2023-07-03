from docxtpl import DocxTemplate
import pandas as pd
from docx2pdf import convert

CLIENTS_SHEET_NAME = 1

def main():
    try:
        excel_path = "invoice_data.xlsx"
        df_invoice = pd.read_excel(excel_path).set_index('invoice_number')
        df_client = pd.read_excel(excel_path, sheet_name=CLIENTS_SHEET_NAME).set_index('id')

        invoice_number = choose_invoice(df_invoice)
        invoice_data = get_invoice(df_invoice, invoice_number)
        client_data = get_client(df_client, invoice_data["client_id"])

        context = create_context(invoice_data, client_data)

        doc = DocxTemplate('Invoice_template.docx')
        doc.render(context)

        filename = f'invoices/{invoice_data.name}.docx'
        doc.save(filename)
        convert(filename)
    except ValueError as e:
        print(e)


def choose_invoice(df_invoice: pd.DataFrame) -> str:
    print(df_invoice[['date','name']].tail(5))
    return input("Enter invoice number: ").strip()

def get_invoice(df_invoice: pd.DataFrame, invoice_number: str) -> pd.Series:
    if invoice_number not in df_invoice.index:
        raise ValueError(f"Invoice number {invoice_number} not found.")
    return df_invoice.loc[invoice_number]

def get_client(df_client: pd.DataFrame, client_id: str) -> pd.Series:
    if client_id not in df_client.index:
        raise ValueError(f"Client id {client_id} not found.")
    return df_client.loc[client_id]

def create_context(invoice_data: pd.Series, client_data: pd.Series) -> dict:
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

if __name__ == "__main__":
    main()
