import pyautogui
import time
import pyperclip
import pandas as pd
from get_PB_Window import check_power_broker, locate_PB_Window

# Function to click on the "Policies" button within Power Broker
def click_policies_button():
    check_power_broker()
    locate_PB_Window()

    policies_button_x = 495
    policies_button_y = 422

    pyautogui.click(x=policies_button_x, y=policies_button_y)
    print("Clicked the Policies button.")
    time.sleep(1)  # Increased wait time for loading

# Function to click on the "PolicySearch" button within Power Broker
def click_policy_search_button():
    check_power_broker()
    locate_PB_Window()

    PolicySearch_button_x = 622
    PolicySearch_button_y = 324

    pyautogui.click(x=PolicySearch_button_x, y=PolicySearch_button_y)
    print("Clicked the PolicySearch button.")
    time.sleep(1)  # Increased wait time for search dialog

# Function to search for a policy number and check for identical match
def search_and_check_policy(policy_number):
    pyperclip.copy(policy_number)

    click_policy_search_button()
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(1)  # Increased wait time for search results

    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    search_result = pyperclip.paste().strip()

    if search_result == str(policy_number).strip():
        print(f"Policy {policy_number} found: Identical")
        return 1
    else:
        print(f"Policy {policy_number} not found: Not identical")
        return 0

# Load policies and add Match Flag
def load_and_check_policies(file_path):
    df = pd.read_excel(file_path, sheet_name=0)
    match_flags = []

    for policy_number in df['Policy#']:
        result_flag = search_and_check_policy(policy_number)
        match_flags.append(result_flag)

    df['Match Flag'] = match_flags
    return df

def double_click_the_policy():
    x_policy_coordinate = 619
    y_policy_coordinate = 633

    pyautogui.click(x=x_policy_coordinate, y=y_policy_coordinate)
    pyautogui.click(x=x_policy_coordinate, y=y_policy_coordinate)
    time.sleep(2.5)  # Slightly longer for double-click response

    print("Entered the Policy Page")

def vehicle_information():
    vehicle_x = 775
    vehicle_y = 268
    pyautogui.click(x=vehicle_x, y=vehicle_y)
    time.sleep(0.5)  # Allow time for vehicle information to load

    star_info_x = 802
    star_info_y = 414
    pyautogui.click(x=star_info_x, y=star_info_y)

    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')

def driver_information():
    driver_x = 
    driver_y = 
    pyautogui.click(x = driver_x, y = driver_y)
    time.sleep(0.5)

    Age_x = 
    Age_y = 
    pyautogui.click(x = Age_x, y = Age_y)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.3)

    Gender_x = 
    Gender_y = 
    pyautogui.click(x = Gender_x, y = Gender_y)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')   

def return_to_homepage():
    return_x = 1427
    return_y = 209
    pyautogui.click(x = return_x, y = return_y)
    time.sleep(2.5)

def retrieve_star_info_and_update_limited(file_path, limit):
    # Load the Excel file
    df = pd.read_excel(file_path, sheet_name=0)
    star_data = []

    # Apply limit only if it's set
    if limit is None:
        rows_to_process = df
    else:
        rows_to_process = df.head(limit)

    for index, row in rows_to_process.iterrows():
        policy_number = row['Policy#']

        if pd.notna(policy_number):
            print(f"Processing Policy: {policy_number}")
            
            # Search and check the policy in Power Broker
            result_flag = search_and_check_policy(policy_number)

            if result_flag == 1:
                double_click_the_policy()
                vehicle_information()

                star_info = pyperclip.paste().strip()
                print(f"Star Info for Policy {policy_number}: {star_info}")
            else:
                star_info = "Not Found"
                print(f"Policy {policy_number} not found.")

            star_data.append(star_info)
            return_to_homepage()
            time.sleep(2)  # Allow UI to reset before next interaction
        else:
            star_data.append("N/A")

    # Update only the rows processed
    df.loc[:len(rows_to_process) - 1, 'Star'] = star_data

    # Save the updated file
    updated_file_path = file_path.replace(".xlsx", "_updated_limited.xlsx")
    df.to_excel(updated_file_path711638232
                , index=False, sheet_name="Sheet2")
    print(f"Updated file saved as: {updated_file_path}")


if __name__ == "__main__":
    # file_path = './Intact_Loss_Report.xlsx'
    file_path = r'C:\Users\Vincent.Zhong\Documents\GitHub\Label_Star\Book10.xlsx'

    retrieve_star_info_and_update_limited(file_path, limit=None)