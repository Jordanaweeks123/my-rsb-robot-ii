from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML reciept as a PDF File.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates a ZIP archive of the receipts and the images.
    """
    # browser.configure(slowmo=1000)
    open_robot_order_website()
    orders = get_orders()
    for order in orders:
        close_annoying_modal()
        fill_the_form(order)

def open_robot_order_website():
    """Mavigate to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """Downloads CSV orders and converts to table"""
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite=True)
    table = Tables()
    return table.read_table_from_csv("orders.csv", header=True).to_list()

def close_annoying_modal():
    """Closes modal upon opening page"""
    page = browser.page()
    page.click(selector="button:text('OK')")

def fill_the_form(order):
    """Fills form using information from order row"""
    page = browser.page()
    page.select_option(selector="#head", value=order.get('Head'))
    page.check(selector=f"#id-body-{order.get('Body')}")
    page.fill(selector="//div[label[text()='3. Legs:']]/input", value=order.get('Legs'))
    page.fill(selector="#address", value=order.get('Address'))
    page.click(selector="#order")
    # Check if any errors occured, retry order if found
    while page.locator(".alert.alert-danger").is_visible():
        page.click(selector="#order")

    order_number = page.locator(selector="//div[@id='receipt']/p[1]").text_content()
    receipt_pdf = store_receipt_as_pdf(order_number)
    screenshot = screenshot_robot(order_number)
    embed_screenshot_to_receipt(screenshot, receipt_pdf)
    archive_receipts()

    page.click(selector="#order-another")

def store_receipt_as_pdf(order_number):
    """Gets HTML receipt element and stores as PDF"""
    page = browser.page()
    page_html = page.inner_html(selector="//div[@id='receipt']")
    pdf = PDF()
    pdf.html_to_pdf(page_html, f'output/receipts/{order_number}.pdf')
    # Return file path
    return f'output/receipts/{order_number}.pdf'

def screenshot_robot(order_number):
    """Takes screenshot of order image"""
    page = browser.page()
    page.locator(selector="//div[@id='robot-preview']").screenshot(path=f'output/screenshots/{order_number}.png')
    # Return file path
    return f'output/screenshots/{order_number}.png'

def embed_screenshot_to_receipt(screenshot, receipt_pdf):
    """Embeds screenshot of order image to PDF"""
    pdf = PDF()
    pdf.add_files_to_pdf(files=[receipt_pdf, screenshot], target_document=receipt_pdf)

def archive_receipts():
    """Zips PDF Receipts"""
    lib = Archive()
    lib.archive_folder_with_zip('output/receipts', 'order_archive.zip')

    
    