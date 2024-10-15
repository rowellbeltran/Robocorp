from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
from datetime import datetime
import os

curr_work_dir = "output/receipts_" + datetime.today().strftime('%Y%m%d_%H%M%S') + "/"

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=500,
    )

    os.mkdir(curr_work_dir)

    open_sparebin_website()
    submit_orders(get_orders())    
    archive_receipts()

   
def open_sparebin_website():
    """Open the sparebin website"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")    

def close_annoying_modal():
    """Close popup message"""
    page = browser.page()
    
    if page.is_visible("//div[@class='modal-content']"):
        page.click("button:text('OK')")

def get_orders():
    """Download and read orders"""
    orderURL = "https://robotsparebinindustries.com/orders.csv"
    orderFilePath = "C:/Users/rowel/Downloads/order.csv"

    http =  HTTP()
    http.download(url=orderURL, overwrite=True, target_file = orderFilePath)

    table = Tables()
    return table.read_table_from_csv(path=orderFilePath, dialect="excel", header=True)

def submit_orders(orders):
    """Submit all downloaded orders"""
    for order in orders:
        try:
            close_annoying_modal()
            fill_the_form(order)
            receipt_file = store_receipt_as_pdf(order["Order number"])
            sc_file = take_screenshot(order["Order number"])
            embed_screenshot_to_receipt(sc_file, receipt_file)
            order_another_bot() 
        except:
            print("Error processing order: " + order["Order number"])
        

def fill_the_form(order):
    """Submit an order"""
    close_annoying_modal()
    page = browser.page()
    page.select_option("#head", str(order["Head"]))
    page.click("#id-body-" + str(order["Body"]))
    page.fill("//div/input[@placeholder='Enter the part number for the legs']", str(order["Legs"]))
    page.fill("#address", str(order["Address"]))
    page.click("#order")

def order_another_bot():
    """Click on order another robot"""
    page = browser.page()
    if page.is_visible("#order-another"):
        page.click("#order-another")

def store_receipt_as_pdf(ord_num):
    page = browser.page()
    receipts_html = page.locator("//*[@id='receipt']").inner_html()
    pdf_filepath = curr_work_dir + "receipt_" + ord_num + ".pdf"

    pdf = PDF()
    pdf.html_to_pdf(receipts_html, pdf_filepath)    
    return(pdf_filepath)
    

def take_screenshot(ord_num):
    page = browser.page()
    sc_filepath = curr_work_dir + "sc_" + ord_num + ".png"

    page.screenshot(path = sc_filepath)
    return(sc_filepath)

def embed_screenshot_to_receipt(screenshot, pdf_file):
    pdf = PDF()
    list_of_files = [
        screenshot
    ]

    pdf.add_files_to_pdf(
        files=list_of_files,
        target_document=pdf_file,
        append=True
    )

def archive_receipts():
    archive = Archive()
    archive.archive_folder_with_zip(curr_work_dir, "output/receipts_" + datetime.today().strftime('%Y%m%d_%H%M%S') + ".zip", include="*receipt*")