from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from RPA.Tables import Tables
from RPA.Archive import Archive
from RPA.FileSystem import FileSystem

tables = Tables()
excel = Files()

@task
def main_task():
    """ Main Function """
    fs = FileSystem()
    fs.create_directory("output")
    orders = get_orders()
    browser.goto("https://robotsparebinindustries.com/#/robot-order")    
    close_annoying_modal()
    fill_and_submit(orders)    
    zip_it()

def get_orders():
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    table = tables.read_table_from_csv("orders.csv")
    return table
    #excel.create_workbook("orders.xlsx")
    #excel.append_rows_to_worksheet(table)
    #excel.save_workbook()
    #excel.close_workbook()

def close_annoying_modal():
    page = browser.page()
    try:                            #chatgpt
        page.click("text = OK")
    except:                         #chatgpt
        pass

def fill_and_submit(worksheet):
    #excel.open_workbook("orders.xlsx")
    #worksheet = excel.read_worksheet_as_table(header = True)
    #excel.close_workbook()
    for row in worksheet:
        fill_up(row)
        get_preview_and_ss(str(row['Order number']))
        submit()
        create_pdf(str(row['Order number']))
        order_next()
        close_annoying_modal()

def fill_up(row):
    page = browser.page()
    page.select_option("#head", value = str(row["Head"]))
    page.click(f'input[name = "body"][value="{row["Body"]}"]')
    page.fill("input[type='number']", str(row["Legs"]))
    page.fill("#address", row["Address"])


def get_preview_and_ss(name):
    page =  browser.page()
    page.click("text = Preview")
    page.wait_for_selector("#robot-preview-image")
    page.locator("#robot-preview-image").screenshot(path = "output/"+name+".png")

def submit():
    page = browser.page()
    while True:
        page.click("text = ORDER")
        try:
            page.wait_for_selector("#order-completion")
            break
        except:
            print("Order failed, retrying...")

def create_pdf(name):
    page = browser.page()
    receipt = page.locator("#order-completion").inner_html()
    pdf = PDF()
    image_path = "output/" + name + ".png"
    pdf_path = "output/" + name + ".pdf"
    pdf.html_to_pdf(receipt, pdf_path)
    pdf.add_files_to_pdf(files=[pdf_path, image_path], target_document=pdf_path)

def order_next():
    page = browser.page()
    page.click("text = ORDER ANOTHER ROBOT")

def zip_it():
    """ Creates a ZIP Folder """
    archive = Archive()
    archive.archive_folder_with_zip(folder="output", archive_name="output/receipts.zip")
    for i in range(1, 21):
        archive.add_to_archive(files=[f"output/{i}.pdf"], archive_name="output/receipts.zip")
    #pdf_files = [f"output/{i}.pdf" for i in range(1, 21)]
    #archive.add_to_archive(files=pdf_files, archive_name="output/receipts.zip")   