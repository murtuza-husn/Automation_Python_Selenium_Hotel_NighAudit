import msvcrt
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta, time
import re
from openpyxl import Workbook
from datetime import datetime
import os

# Path to save file on Desktop
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
url = "https://www.choiceadvantage.com/choicehotels/sign_in.jsp"
driver_path = r"C:\Program Files (x86)\chromedriver-win64\chromedriver.exe"
inHouseList_URL = "https://www.choiceadvantage.com/choicehotels/ViewInHouseList.init"
todaysCheckedOutGuest_URL = "https://www.choiceadvantage.com/choicehotels/ViewCheckedOutList.init"

def get_headless_driver():
    chrome_options = Options()
    #chrome_options.add_argument("--headless=new")  # Use 'new' headless mode
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920x1080")

    try:
        # Automatically download the correct driver
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("Driver loaded successfully!")
        return driver
    except Exception as e:
        print("An error occurred:", e)
        return None

def get_password(prompt="Enter password: "):
    print(prompt, end="", flush=True)
    password = ""
    while True:
        ch = msvcrt.getch()
        if ch in {b'\r', b'\n'}:  # Enter key pressed
            print("")  # Move to next line
            break
        elif ch == b'\x08':  # Backspace
            if len(password) > 0:
                password = password[:-1]
                print("\b \b", end="", flush=True)
        else:
            password += ch.decode("utf-8")
            print("*", end="", flush=True)
    return password



def get_credentials(driver):

    userName = input("Enter Username : ")
    password = pwd = get_password()

    try:
        driver.get(url)
        print(driver.title)
        username_field = driver.find_element(By.NAME, "j_username")
        password_field = driver.find_element(By.NAME, "j_password")
        username_field.send_keys(userName)
        password_field.send_keys(password)
        login_button = driver.find_element(By.ID, "greenButton")
        login_button.click()
        WebDriverWait(driver, 25).until(
            EC.presence_of_element_located((By.ID, "infoNews"))
        )
        print("\nLogin Successful!")
    except Exception as e:
        raise Exception("The Credentials entered are not valid. Please try again.") from e

def todaysCheckedOutGuest(driver):
    driver.get(todaysCheckedOutGuest_URL)
    checkout_count = 0
    tbody = driver.find_element("id", "checkedOutList")
    rows = tbody.find_elements("tag name", "tr")
    checkedout_room_numbers = []
    for row in rows:
        cells = row.find_elements("tag name", "td")
        if len(cells) >= 4:
            room_text = cells[3].text.strip()
            if room_text.isdigit():  # make sure it's a number
                checkedout_room_numbers.append(int(room_text))

    checkedout_room_numbers.sort()
    print("Today's Checked OUt Guest List :")
    for room in checkedout_room_numbers:
        print(room)
        checkout_count += 1
    print(f"The total no of Checked Out guests : {checkout_count}")
    return checkedout_room_numbers

#Function for Check-In Rooms
def inHouseList(driver):
    driver.get(inHouseList_URL)
    print("\n")
    print(f"Now Printing the Checked Out Guest List")
    now = datetime.now()
    current_time = now.time()

    # Define the window: 12:00 AM to 6:00 AM
    midnight = time(0, 0)
    six_am = time(6, 0)

    # Determine filter date
    if midnight <= current_time < six_am:
        filter_date = (now - timedelta(days=1)).date()
    else:
        filter_date = now.date()

    # Format filter_date string according to your table date format
    # Here I assume arrival dates are in "YYYY-MM-DD" format in the table cells
    filter_date_str = filter_date.strftime("%m/%d/%Y")
    print(f"Filtering Date : {filter_date_str}")
    # Locate the table body
    tbody = driver.find_element("id", "inHouseList")
    rows = tbody.find_elements("tag name", "tr")

    checkin_room_numbers = []
    arrival_count = 0
    # or better explicit wait here
    for row in rows:
        cells = row.find_elements("tag name", "td")
        if not cells:
            continue
        arrival_date_text = cells[9].text.strip()
        arrival_date_text = datetime.strptime(arrival_date_text, "%m/%d/%Y").strftime("%m/%d/%Y")
        try:
            if arrival_date_text == filter_date_str:
                room_text = cells[5].text.strip()
                if room_text.isdigit():
                    checkin_room_numbers.append(int(room_text))
        except Exception as e:
            print("No Matching entries found")
    # Sort the room numbers ascending
    checkin_room_numbers.sort()
    print("Printing the values of checkin room numbers:")
    print(checkin_room_numbers)

    print("Filtered and sorted Room Numbers for arrival date", filter_date_str)
    for room in checkin_room_numbers:
        print(room)
        arrival_count += 1
    print(f"The Number of Arrivals for today {filter_date} is : {arrival_count}")
    return checkin_room_numbers

def vacant_list(checkin_room_numbers, checkedout_room_numbers):
    print("The Rooms that are Vacant - Not Checked In")
    vacant_list = []
    for room in checkedout_room_numbers:
        if room not in checkin_room_numbers:
            vacant_list.append(room)
            print(room)
    return vacant_list

rooms_data = {}
def GuestTracking(driver, checkin_room_numbers, ws):
    error_list = []
    driver.get(inHouseList_URL)
    WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.ID, "inHouseList"))
    )

    for in_room in checkin_room_numbers:
        print(f"Processing room {in_room}")

        try:
            # Refresh table each loop to avoid stale elements
            tbody = driver.find_element(By.ID, "inHouseList")
            rows = tbody.find_elements(By.TAG_NAME, "tr")

            room_found = False

            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                if not cells:
                    continue

                room_num = cells[5].text.strip()
                if room_num.isdigit() and int(room_num) == in_room:
                    room_found = True
                    name = cells[1].text.strip()
                    print(f"Opening details for {name} in room {room_num}")

                    # Open guest details
                    element_name = driver.find_element(By.LINK_TEXT, name)
                    element_name.click()
                    WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.ID, "guestFolioEnabled"))
                    )
                    print(f"Guest Name is : {name}")

                    estimated_total_cost = driver.find_element(By.ID, "estimated_total_cost").text
                    match = re.search(r"\d+(?:\.\d+)?", estimated_total_cost)
                    if match:
                        print(f"Estimated Total Cost: {match.group()}")
                    else:
                        print("Could not find estimated total cost:", estimated_total_cost)

                    # Check rate plan
                    plan = driver.find_element(By.ID, "ratePlan").text.strip()
                    print(f"Guest Plan: {plan}")
                    plans_to_skip = ["SRD", "LCITY"]

                    if plan not in plans_to_skip:
                        driver.find_element(By.ID, "guestFolioEnabled").click()
                        WebDriverWait(driver, 3).until(
                            EC.presence_of_element_located((By.ID, "button_12"))
                        )
                        balance = driver.find_element(By.ID, "guestFolioBalance").text
                        match = re.search(r"\d+(?:\.\d+)?", balance)
                        if match:
                            print(f"Balance: {match.group()}")
                        else:
                            print("Could not find Balance:", balance)

                        # Click 'View Estimated Cost'
                        driver.find_element(By.ID, "button_12").click()
                        table = WebDriverWait(driver, 3).until(
                            EC.presence_of_element_located((By.XPATH, "//table"))
                        )
                        #Authorization and Card
                        auth_rows = driver.find_elements(By.XPATH, "//table/tbody/tr")

                        cards = []
                        existing_auths = []

                        # Extract Card and Existing Auth values for each row
                        for authrow in auth_rows:
                            try:
                                card = authrow.find_element(By.XPATH, "./td[4]/p/em").text.strip()  # 4th td = Card
                                existing_auth = authrow.find_element(By.XPATH,
                                                                 "./td[8]/p").text.strip()  # 8th td = Existing Auth
                                if existing_auth:  # only consider rows with values
                                    existing_auth_value = float(existing_auth)
                                    cards.append(card)
                                    existing_auths.append(existing_auth_value)
                            except Exception as e:
                                # Skip rows without proper data
                                continue

                        # Apply the 3-case logic
                        selected_card = None
                        selected_existing_auth = None

                        if len(existing_auths) == 1:
                            # Case 1: only one row with a value
                            selected_card = cards[0]
                            selected_existing_auth = existing_auths[0]
                        elif 25.00 in existing_auths:
                            # Case 2: one of the rows has 25.00
                            index = existing_auths.index(25.00)
                            selected_card = cards[index]
                            selected_existing_auth = existing_auths[index]
                        else:
                            # Case 3: pick the row with the highest Existing Auth value
                            max_value = max(existing_auths)
                            index = existing_auths.index(max_value)
                            selected_card = cards[index]
                            selected_existing_auth = existing_auths[index]

                        print(f"Selected Card: {selected_card}")
                        print(f"Selected Existing Auth: {selected_existing_auth}")
                        # file_name = f"{timestamp}.xlsx"
                        # file_path = os.path.join(desktop_path, file_name)
                        # wb = Workbook()
                        # ws = wb.active
                        row_data = [
                            in_room,  # In Room
                            name,  # Name
                            selected_card,  # Selected Card
                            selected_existing_auth,  # Selected Existing Card
                            balance,  # Payment
                            "",  # Empty column
                            "",  # Empty column
                            estimated_total_cost,  # Estimated Total Cost
                            "COVERED"  # Status
                        ]

                        ws.append(row_data)
                        # Save the workbook


                    # After finishing this room, go back to inHouseList page
                    driver.get(inHouseList_URL)
                    WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.ID, "inHouseList"))
                    )
                    break  # Exit rows loop after finding the room

            if not room_found:
                print(f"Room {in_room} not found in the current in-house list.")

        except Exception as e:
            error_list.append(in_room)
    return error_list

# def repeat(error_list, max_tries =3):
#     if len(error_list) != 0:
#         GuestTracking(driver, error_list)
def retry_guest_tracking(driver, checkin_room_numbers, max_retries=20):
    """
    Repeatedly call GuestTracking on the remaining rooms until no errors or max_retries reached.
    """
    error_list = []
    remaining_rooms = checkin_room_numbers.copy()
    attempt = 1

    file_name = "october.xlsx"
    file_path = os.path.join(desktop_path, file_name)

    # Create workbook and sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Room Payments"

    headers = ["Room No", "Name", "Plan", "Card Type", "Authorization", "Payment", "Deposit/Refund", "Balance", "Est. Total Cost", "Room Covered"]
    ws.append(headers)


    while remaining_rooms and attempt <= max_retries:
        print(f"\nAttempt #{attempt} for rooms: {remaining_rooms}")
        error_list = GuestTracking(driver, remaining_rooms, ws)

        if not error_list:
            print("All rooms processed successfully!")
            break

        print(f"Rooms that failed in attempt #{attempt}: {error_list}")
        remaining_rooms = error_list
        attempt += 1

    if error_list:
        print(f"\nThe following rooms could not be processed after {max_retries} attempts: {error_list}")
    else:
        print(f"\nAll rooms processed successfully after {attempt} retries!")

    wb.save(file_path)
    print(f"Excel file saved at: {file_path}")

def main():
    driver = get_headless_driver()
    if driver is None:
        print("Failed to load driver, exiting.")
        return
    try:
        get_credentials(driver)
        checkin_rooms = inHouseList(driver)
        checkedout_rooms = todaysCheckedOutGuest(driver)
        vacant_rooms = vacant_list(checkin_rooms, checkedout_rooms)
        retry_guest_tracking(driver, checkin_rooms)

    finally:
        try:
            driver.quit()
        except Exception as e:
            print(f"Error quitting driver: {e}")


if __name__ == '__main__':
    main()
