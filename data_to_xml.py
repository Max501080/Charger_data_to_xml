from playwright.sync_api import sync_playwright
import time
import os
from datetime import datetime, timedelta
import shutil

now = datetime.now()
target_weekday = 6 #6 = sunday(0 or 7 would be monday, 1 would be tuesday, 2 would be wednesday, and so on)
days_since_target = ((now.weekday() - target_weekday) % 7)+7 #now.weekday returns a number 0 through 6, 0 being monday and six being sunday 
                                                             #the +7 at the end makes this return the target day of last week so when doing it
                                                             #on monday, it wont return yesterday, instead it returns the sunday before yesterday
first_day = now - timedelta(days=days_since_target)

last_day = first_day + timedelta(days=6)

first_day = first_day.replace(hour=0, minute=0, second=0, microsecond=0)  #set the time of the first day to 12:00 AM
last_day = last_day.replace(hour=11, minute=59, second = 0, microsecond=0)  #set the time of the last day to 11:59 PM

paths = []
url = '' #overview page where we start the download process from
login = '' #login page

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        page.goto(login)
        page.fill('input[name="identifier"]', '')  #where your un needs to go
        page.click('button[type="submit"]')
        page.fill('input[name="pd"]', '') # where your pd needs to go
        page.click('button[type="submit"]')
        page.wait_for_url('') # home page
        context.storage_state(path="auth.json") # save login cookies to maintain login

        for path in paths:
            page.goto(url)
            page.wait_for_selector(path)
            page.click(path)
            name = page.locator('h4.ng-star-inserted').inner_text()
            page.get_by_text("Connector 1").wait_for()
            page.get_by_text("Connector 1").click()
            time.sleep(1)
            page.fill('input[placeholder="From"]', first_day.strftime("%m/%d/%Y %H:%M")+" AM")
            page.fill('input[placeholder="To"]', last_day.strftime("%m/%d/%Y %H:%M")+" PM")
            page.click('button[aria-label="Open calendar"]')
            time.sleep(.5)
            page.mouse.click(0,0)

            try:
                page.get_by_text("Delivered Energy (kWh)").wait_for(timeout=2000)
            except:
                continue
            page.get_by_text("Delivered Energy (kWh)").click()

            page.click('path.highcharts-button-symbol')

            page.get_by_text('Download XLS').wait_for()
            with page.expect_download() as d:
                page.get_by_text('Download XLS').click()

            filename =  name +' '+ first_day.strftime('%m_%d')+'-'+ last_day.strftime('%m_%d') + '.xls'
            downloads_dir = os.path.expanduser("~/Downloads")
            foldername = first_day.strftime('%m_%d')+'-'+ last_day.strftime('%m_%d')
            filepath = os.path.join(downloads_dir, os.path.join(foldername, filename))
            download = d.value
            download.save_as(filepath)

        time.sleep(1)
        browser.close()
    folder_to_zip = os.path.join(downloads_dir, foldername)
    output_zip_path = os.path.join(downloads_dir, foldername) 
    shutil.make_archive(output_zip_path, 'zip', folder_to_zip)
    time.sleep(10)
if __name__ == "__main__":
    main()
    