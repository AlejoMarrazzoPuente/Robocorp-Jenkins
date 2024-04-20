from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.PDF import PDF
from sqlite3 import Cursor, Connection
import sqlite3


class SqliteDB():
    DATABASE: str
    CURSOR: Cursor
    CONNECTION: Connection
    TABLE_NAME: str

    def __init__(self, db_name: str) -> None:
        self.set_database(db_name)
        self.create_db()
        self.set_cursor()

    def create_db(self) -> None:
        """Create or connect to a database"""
        self.CONNECTION = sqlite3.connect(self.DATABASE)

    def set_database(self, db: str) -> None:
        """Set a database name to connect or create it"""
        self.DATABASE = db

    def set_cursor(self) -> None:
        """Set a cursor. 
        WARNING: You should have set a connection/create a db previously"""
        if self.CONNECTION == None:
            raise Exception("System Error - You should have set a connection previously")
        self.CURSOR = self.CONNECTION.cursor()
    
    def create_table(self, table_name: str ,fields: str) -> None:
        """Create table based on the fields given by parameter."""
        if fields != "" and table_name != "":
            self.TABLE_NAME = table_name
            self.CURSOR.execute("CREATE TABLE "+table_name+"(" + fields + ")")

    def insert_row(self, row: str) -> None:
        """Insert a row into the database"""
        self.CURSOR.execute("INSERT INTO " + self.TABLE_NAME + " VALUES " + row)
        self.CONNECTION.commit()

    def mark_as_completed(self, first_name: str) -> None:
        self.CURSOR.execute("UPDATE " + self.TABLE_NAME + " SET queued = 0, status = 'completed' WHERE first_name = '" + first_name + "'")
        self.CONNECTION.commit()

    def get_next_item(self):
        return self.CURSOR.execute("SELECT * FROM "+self.TABLE_NAME+" WHERE queued = 1 LIMIT 1")


### SE INSTANCIAN CLASES GLOBALES ###
sql = SqliteDB("WQ.db")
sql.TABLE_NAME = "item"
### FIN INSTANCIACION DE CLASES GLOBALES ###



### INICIO - GENERAR DB SOLAMENTE LA PRIMERA VEZ ###
@task
def generar_db():
    sql.create_table(table_name="item", fields="first_name, last_name, sales, sales_target, queued, status")
### FIN - GENERAR DB SOLAMENTE LA PRIMERA VEZ ### 



### INICIO - MAIN PROCESS ###
@task
def robot_spare_bin_python():
    """Insert the sales data for the week and export it as PDF"""
    browser.configure(
        slowmo=100,
    )
    open_the_intranet_website("https://robotsparebinindustries.com/")
    log_in()
    populate_queue()
    item: tuple = get_next_item()
    while(item != None):
        fill_and_submit_sales_form(item[0],item[1], item[2], item[3])
        sql.mark_as_completed(item[0])
        item = get_next_item()
    collect_results()
    export_as_pdf()
    log_out()
### FIN - MAIN PROCESS ###



### INICIO - PAGINAS UTILIZADAS EN EL MAIN PROCESS ###
def open_the_intranet_website(url: str) -> None:
    """Navigates to the given URL"""
    browser.goto(url)

def log_in() -> None:
    page = browser.page()
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('log in')")

def download_excel_file() -> None:
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)

def populate_queue() -> None:
    download_excel_file()
    excel = Files()
    excel.open_workbook("SalesData.xlsx", )
    worksheet = excel.read_worksheet_as_table("data", header=True)
    excel.close_workbook()
    for row in worksheet:
        item = "('" + row["First Name"] + "', '" + str(row["Last Name"]).replace("'", "") + "', '" + str(row["Sales"]) + "', '" + str(row["Sales Target"]) + "', 1, 'new')"
        print(item)
        sql.insert_row(item)

def get_next_item() -> tuple:
    return sql.get_next_item().fetchone()

def fill_and_submit_sales_form(first_name, last_name, sales, sales_target) -> None:
    """Fills in the sales data and click the 'Submit' button"""
    page = browser.page()
    page.fill("#firstname", first_name)
    page.fill("#lastname", last_name)
    page.select_option("#salestarget", sales_target)
    page.fill("#salesresult", sales)
    page.click("text=Submit")

def collect_results() -> None:
    """Take a screenshot of the page"""
    page = browser.page()
    page.screenshot(path="output/sales_summary.png")

def export_as_pdf() -> None:
    """Export the data to a pdf file"""
    page = browser.page()
    sales_results_html = page.locator("#sales-results").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")

def log_out() -> None:
    """Presses the 'Log out' button"""
    page = browser.page()
    page.click("#logout")
### FIN - PAGINAS UTILIZADAS EN EL MAIN PROCESS ###


  
