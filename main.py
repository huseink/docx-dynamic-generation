from fastapi import FastAPI
from docxtpl import DocxTemplate, InlineImage
from pydantic import BaseModel
from typing import List
from io import BytesIO
from fastapi.responses import StreamingResponse
from docx.shared import Mm
import requests

app = FastAPI()


class Company(BaseModel):
    name: str
    logo: str | None = None
    address_line_1: str | None = None
    address_line_2: str | None = None
    address_line_3: str | None = None
    phone_number: str | None = None
    registered_id: str


class BankInformation(BaseModel):
    name: str
    bank_name: str
    bank_address: str
    bank_swift_code: str
    account_number: str
    iban: str


class Item(BaseModel):
    description: str
    amount: int


class VatInformation(BaseModel):
    description: str
    percentage: int
    amount: int


class InvoiceContext(BaseModel):
    company: Company
    billed_company: Company
    beneficiary: BankInformation
    sender: BankInformation
    items: List[Item]
    vat_info: VatInformation | None = None

    invoice_no: str
    invoice_date: str
    total_amount: int = 0


def get_image_from_url(image_url: str):
    """Fetches an image from a provided URL and returns it as a BytesIO object."""
    response = requests.get(image_url)
    image = BytesIO(response.content)
    return image


def process_logo(template, logo_url: str):
    """Process the logo URL and return the InlineImage if a valid URL is provided."""
    if logo_url:
        image = get_image_from_url(logo_url)
        return InlineImage(template, image, width=Mm(21))
    return None


def process_total_amount(items: List[Item], vat_info: VatInformation | None) -> int:
    """
    Calculates the total amount for the invoice by summing the amounts of all items
    and adding the VAT (if provided).
    """
    total = sum(item.amount for item in items)

    if vat_info:
        # Calculate the VAT amount
        vat_amount = (total * vat_info.percentage) / 100
        total += vat_amount

    return total


@app.post("/")
async def create_invoice(context: InvoiceContext):
    # Load the template
    template = DocxTemplate("invoice_tpl.docx")

    context.company.logo = process_logo(template, context.company.logo)
    context.total_amount = process_total_amount(
        context.items, context.vat_info)

    # Render the template with the context
    template.render(context)

    # Save the document into a BytesIO buffer
    result = BytesIO()
    template.save(result)

    # Rewind the buffer to the beginning
    result.seek(0)

    # Return the result as a StreamingResponse
    return StreamingResponse(result,
                             media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                             headers={"Content-Disposition": "attachment; filename=invoice.docx"})
