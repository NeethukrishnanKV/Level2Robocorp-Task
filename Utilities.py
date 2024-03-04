
from robocorp import browser
from RPA.Browser.Selenium import Selenium
from RPA.HTTP import HTTP
from RPA.Excel import Files
from RPA.PDF import PDF
from RPA.Tables import Tables
from RPA.Archive import Archive
import time

def open_robot_order_website():
    """Open the site RobotSpareBin Industries Inc."""
    browser.configure(
        headless=False,
        slowmo=100,
    )
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def get_orders():
    """Get Orders robots from RobotSpareBin Industries Inc."""    
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv",target_file="Input/orders.csv", overwrite=True)
    tables=Tables()
    orders = tables.read_table_from_csv("Input/orders.csv", header=True)
    return orders
def close_annoying_modal():
    """Close the popup"""
    page = browser.page()
    page.click("button:text('OK')")
def fill_the_form(orders):
    """Fill the form wih orders list and create the robot,get PDF and screenshot of the order"""
    page = browser.page()
    time.sleep(5)
    for row in orders:
        close_annoying_modal()
        page.select_option("#head",index=int(row["Head"]))
        num=row["Body"]
        page.click("#id-body-"+num)
        leg=row["Legs"]
        page.fill('input[placeholder="Enter the part number for the legs"]',leg)
        page.fill("#address",row["Address"])
        page.click("#order")
        time.sleep(1) 
        result = page.is_visible("#receipt")
        while(result == False):     
            page.dblclick("#order")
            time.sleep(1)
            result = page.is_visible("#receipt")
        receiptPDF=store_receipt_as_pdf(row["Order number"])
        time.sleep(1)
        screenshotrobot=screenshot_robot(row["Order number"])
        time.sleep(6)
        embed_screenshot_to_receipt(screenshotrobot,receiptPDF)
        page.click("#order-another")
        time.sleep(3)
def store_receipt_as_pdf(order_number):
    """Store receiot in output folder"""
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()

    pdf = PDF()
    receiptPDF="output/receipts/receipt"+order_number+".pdf"
    pdf.html_to_pdf(receipt_html, receiptPDF)    
    return receiptPDF
def screenshot_robot(order_number):
    """Store screenshot in output folder"""
    page = browser.page()
    screenshotrobot="output/Screenshots/robot"+order_number+".png"
    x1=float(750)
    y1=float(70)
    height1=float(1100)
    width1=float(500)
    clipdict = dict(x=x1,y=y1,width=width1,height=height1)
    page.screenshot(path=screenshotrobot, clip=clipdict) 
    return screenshotrobot
def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Merge PDF and screenshot in one file"""
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot,source_path=pdf_file,output_path=pdf_file) 
def archive_receipts():
    """Zip the output PDF folder and keep it in Output folder"""
    zipfile=Archive()
    zipfile.archive_folder_with_zip("output/receipts","output/RobotReceipts.zip")             