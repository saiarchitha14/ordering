from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables
import time
from RPA.PDF import PDF 
from RPA.Archive import Archive

page = browser.page()

@task
def order_robots_from_RobotSpareBin():
    browser.configure(slowmo=100)
    open_robots_order_website()
    get_orders()
    open_form()
    close_annoying_modal()
    fill_the_form()
    archive_receipts()
    # store_receipt_as_pdf()
    # screenshot_robot()
    # embed_screenshot_to_receipt()


def open_the_intranet_browser():
    browser.goto("https://robotsparebinindustries.com/")

def open_robots_order_website():
    browser.goto("https://robotsparebinindustries.com/")
    
def get_orders():
    http=HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    
    

def open_form():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def fill_the_form():
    # page=browser.page()
    tables = Tables()
    orders = tables.read_table_from_csv("C:/Users/LENOVO/Desktop/ARCHITHA/Robocorp Sema4/Level 2/orders.csv", header=True)
    for row in orders:
        print(row)
        #fill_the_form(row)
        steps(row)
        # order()
        # error_msg=page.query_selector(selector='//div[@role="alert"]')
        # print(error_msg)
        # if error_msg:
        #     steps(row)
    
def close_annoying_modal():
#    page=browser.page()
    close_popup = page.query_selector('//button[text() = "OK"]')
    if close_popup:
       page.click('//button[text() = "OK"]')
    else:
       pass

def select_function(head):
    if head=='1':
        return "Roll-a-thor head"
    elif head=='2':
        return "Peanut crusher head"
    elif head=='3':
        return "D.A.V.E head"
    elif head=='4':
        return "Andy Roid head"
    elif head=='5':
        return "Spanner mate head"
    elif head=='6':
        return "Drillbit 2000 head"
    

def select_button(body):
    return f'//input[@id="id-body-{body}"]'
    # if body==1:
    #     return '//input[@id="id-body-1"]'
    # elif body==2:
    #     return '//input[@id="id-body-2"]'


def steps(row):
    # page.click(f'//*[@id="head"]')
    page.select_option(f'//*[@id="head"]',select_function(row["Head"]))
    page.click(selector=select_button(row["Body"]))
    page.fill('//input[@type="number"]', str(row["Legs"]))
    page.fill('//input[@type="text"]', str(row["Address"]))
    page.click('//button[@id="preview"]')
    time.sleep(2)
    # page.click('//button[@id="order"]')
    while True:  
        page.click("css=#order")
        next_order_button = page.query_selector("css=#order-another")
        if next_order_button:
            store_receipt_as_pdf(row["Order number"])
            screenshot = screenshot_robot(row["Order number"])
            embed_screenshot_to_receipt(screenshot, f"output/receipts/{row['Order number']}.pdf")
            page.click("css=#order-another")
            print(f"Order {row['Order number']} successful")
            close_annoying_modal()
            break  
        else:
            print("Order failed, trying again")

def order():
    page.click('//button[@type="submit"]')
    close_annoying_modal()

def store_receipt_as_pdf(order_number):
    page=browser.page()
    order_results_html=page.locator('//div[@id="receipt"]').inner_html()
    print(order_results_html)
    pdf=PDF()
    pdf_path=f"output/receipts/{order_number}.pdf".format(order_number)
    pdf.html_to_pdf(order_results_html,pdf_path)
    return pdf_path

def screenshot_robot(order_number):
    screenshot_path="output/screenshots/{order_number}.png"
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)
    return screenshot_path

def embed_screenshot_to_receipt(screenshot_path, pdf_path):
    pdf=PDF()
    pdf.add_watermark_image_to_pdf(image_path=screenshot_path, source_path=pdf_path, output_path=pdf_path)

def order_another_robot():
    page.click(f'//button[@id="order-another"]')

def archive_receipts():
    lib = Archive()
    lib.archive_folder_with_zip("./output/receipts", "./output/receipts.zip")