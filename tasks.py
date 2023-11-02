from robocorp import browser, vault
from robocorp.tasks import task
import time
import os
import re
from playwright.sync_api import Page, expect
from RPA.Excel.Files import Files as Excel


@task
def solve_challenge():
    """Solve the Automation challenge"""
    #Configure Browser
    browser.configure(
        browser_engine="chromium",
        screenshot="only-on-failure",
        headless=False,
    )

    #Open Automation Challenge Site
    page = browser.goto("https://www.theautomationchallenge.com/")


    #Download the Excel input file
    full_file_path = download_excel_file(page)

    #Read all rows in the excel file
    excel = Excel()
    excel.open_workbook(full_file_path)
    rows = excel.read_worksheet_as_table(
        excel.get_active_worksheet(), header=True)


    #Get credentials from Vault
    user_id, password = get_credentials('rpa_challenge_creds')

    #Login using credentials
    login(page, user_id, password)

    #Start the Challenge
    page.get_by_role("button", name="Start").click()

    #For each row in the excel, submit values
    for row in rows:
        fill_and_submit_form(page, row=row)

    page.close()


def download_excel_file(page):
    with page.expect_download() as download_info:
        page.get_by_text(re.compile(
            "^.*Download Excel Spreadsheet.*$")).click()
    download = download_info.value
    file_name = download.suggested_filename
    destination_folder_path = "./data/"
    full_file_path = os.path.join(destination_folder_path, file_name)
    download.save_as(full_file_path)
    return full_file_path


def get_credentials(credential_name):
    secret = vault.get_secret(credential_name)
    user_id = secret['user_id']
    password = secret['password']
    return user_id, password


def login(page, user_id, password):
    # Click 'sign-up or login' button
    page.get_by_text(re.compile("^.*SIGN UP OR LOGIN.*$")).click()
    page.get_by_text("OR LOGIN", exact=True).click()
    expect(page.get_by_role("textbox", name="Email")).to_be_visible()
    page.get_by_role("textbox", name="Email").fill(user_id)
    page.get_by_role("textbox", name="Password").fill(password)
    page.get_by_text(re.compile("^.*LOG IN.*$")).click()
    time.sleep(3)


def fill_and_submit_form(page, row):
    field_mapping = {"Company Name": 'company_name',
                     "Address": 'company_address',
                     "EIN": 'employer_identification_number',
                     "Sector": 'sector',
                     "Automation Tool": 'automation_tool',
                     "Annual Saving": 'annual_automation_saving',
                     "Date": 'date_of_first_project'}
    for key in field_mapping.keys():
        if page.get_by_text(re.compile("^.*Get through this reCAPTCHA to continue.*$")).is_visible():
            if page.get_by_role("button", name="presentation").is_visible():
                page.get_by_role("button", name="presentation").click()
                time.sleep(1)
        page.locator("div[class=\"bubble-element Group\"]").filter(has_text=re.compile(
            f"^{key}$")).get_by_role("textbox").fill(str(row[field_mapping.get(key)]), force=True)

    page.get_by_role("button", name="Submit").click()
