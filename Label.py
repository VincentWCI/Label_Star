import pyautogui
import time
import pyperclip
import pandas as pd
from pyautogui import FailSafeException
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
    driver_x = 974
    driver_y = 269
    pyautogui.click(x = driver_x, y = driver_y)
    time.sleep(0.5)

    Age_x = 873
    Age_y = 370
    pyautogui.click(x = Age_x, y = Age_y)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.3)
    Age = pyperclip.paste().strip()

    Gender_x = 714
    Gender_y = 390
    pyautogui.click(x = Gender_x, y = Gender_y)
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')   
    Gender = pyperclip.paste().strip()

    return Age, Gender

def return_to_homepage():
    return_x = 1427
    return_y = 209
    pyautogui.click(x = return_x, y = return_y)
    time.sleep(2.5)

# def retrieve_star_info_and_update_limited(file_path, limit):
#     # Load the Excel file
#     df = pd.read_excel(file_path, sheet_name=0)
#     star_data = []

#     # Apply limit only if it's set
#     if limit is None:
#         rows_to_process = df
#     else:
#         rows_to_process = df.head(limit)

#     for index, row in rows_to_process.iterrows():
#         policy_number = row['Policy#']

#         if pd.notna(policy_number):
#             print(f"Processing Policy: {policy_number}")
            
#             # Search and check the policy in Power Broker
#             result_flag = search_and_check_policy(policy_number)

#             if result_flag == 1:
#                 double_click_the_policy()
#                 vehicle_information()

#                 star_info = pyperclip.paste().strip()
#                 print(f"Star Info for Policy {policy_number}: {star_info}")
#             else:
#                 star_info = "Not Found"
#                 print(f"Policy {policy_number} not found.")

#             star_data.append(star_info)
#             return_to_homepage()
#             time.sleep(2)  # Allow UI to reset before next interaction
#         else:
#             star_data.append("N/A")

#     # Update only the rows processed
#     df.loc[:len(rows_to_process) - 1, 'Star'] = star_data

#     # Save the updated file
#     updated_file_path = file_path.replace(".xlsx", "_updated_limited.xlsx")
#     df.to_excel(updated_file_path, index=False, sheet_name="info_extracted")
#     print(f"Updated file saved as: {updated_file_path}")


# def retrieve_star_info_and_update_limited(file_path, limit):
#     df = pd.read_excel(file_path, sheet_name=0)
#     star_data = []
#     age_data = []
#     gender_data = []

#     rows_to_process = df if limit is None else df.head(limit)

#     for index, row in rows_to_process.iterrows():
#         policy_number = row['Policy#']

#         if pd.notna(policy_number):
#             print(f"Processing Policy: {policy_number}")

#             result_flag = search_and_check_policy(policy_number)

#             if result_flag == 1:
#                 double_click_the_policy()
#                 vehicle_information()
#                 star_info = pyperclip.paste().strip()
#                 print(f"Star Info for Policy {policy_number}: {star_info}")

#                 # Now get driver information
#                 age, gender = driver_information()
#                 print(f"Driver Age: {age}, Gender: {gender}")
#             else:
#                 star_info = "Not Found"
#                 age = "N/A"
#                 gender = "N/A"
#                 print(f"Policy {policy_number} not found.")

#             star_data.append(star_info)
#             age_data.append(age)
#             gender_data.append(gender)

#             return_to_homepage()
#             time.sleep(2)
#         else:
#             star_data.append("N/A")
#             age_data.append("N/A")
#             gender_data.append("N/A")

#     # Add the new data to the DataFrame
#     df.loc[:len(rows_to_process) - 1, 'Star'] = star_data
#     df.loc[:len(rows_to_process) - 1, 'Driver Age'] = age_data
#     df.loc[:len(rows_to_process) - 1, 'Driver Gender'] = gender_data

#     updated_file_path = file_path.replace(".xlsx", "_updated_limited.xlsx")
#     df.to_excel(updated_file_path, index=False, sheet_name="info_extracted")
#     print(f"Updated file saved as: {updated_file_path}")

def retrieve_star_info_and_update_limited(file_path, limit=None):
    df = pd.read_excel(file_path, sheet_name=0)

    # Ensure the columns exist
    for col in ['Star', 'Driver Age', 'Driver Gender']:
        if col not in df.columns:
            df[col] = None

    rows_to_process = df if limit is None else df.head(limit)

    for index, row in rows_to_process.iterrows():
        policy_number = row['Policy#']

        # Skip if already processed
        if pd.notna(row.get('Star')) and pd.notna(row.get('Driver Age')) and pd.notna(row.get('Driver Gender')):
            print(f"Skipping Policy {policy_number}: already processed.")
            continue

        try:
            if pd.notna(policy_number):
                print(f"Processing Policy: {policy_number}")

                result_flag = search_and_check_policy(policy_number)

                if result_flag == 1:
                    double_click_the_policy()
                    vehicle_information()
                    star_info = pyperclip.paste().strip()

                    age, gender = driver_information()
                    print(f"→ Star: {star_info}, Age: {age}, Gender: {gender}")
                else:
                    star_info = "Not Found"
                    age = "N/A"
                    gender = "N/A"
                    print(f"Policy {policy_number} not found.")

                # Save data to DataFrame
                df.at[index, 'Star'] = star_info
                df.at[index, 'Driver Age'] = age
                df.at[index, 'Driver Gender'] = gender

                # Save after each entry
                temp_file_path = file_path.replace(".xlsx", "_in_progress.xlsx")
                df.to_excel(temp_file_path, index=False)
                print(f"Progress saved to: {temp_file_path}")

                return_to_homepage()
                time.sleep(2)
        except FailSafeException:
            print(f"❗ Fail-safe triggered while processing policy: {policy_number}. Exiting safely.")
            break
        except Exception as e:
            print(f"❗ Error processing policy {policy_number}: {e}")
            continue  # Skip to next row

    # Save final version
    final_file_path = file_path.replace(".xlsx", "_completed.xlsx")
    df.to_excel(final_file_path, index=False)
    print(f"✅ Final file saved as: {final_file_path}")

def resume_star_driver_processing(file_path, limit=None):
    import pandas as pd

    df = pd.read_excel(file_path)

    # Ensure necessary columns exist
    for col in ['Star', 'Driver Age', 'Driver Gender']:
        if col not in df.columns:
            df[col] = None

    rows_to_process = df if limit is None else df.head(limit)

    for index, row in rows_to_process.iterrows():
        policy_number = row['Policy#']

        # Skip rows that are fully processed
        if all(pd.notna(row.get(col)) and row.get(col) != 'Not Found' for col in ['Star', 'Driver Age', 'Driver Gender']):
            continue

        try:
            if pd.notna(policy_number):
                print(f"▶️ Resuming from Policy: {policy_number} (row {index+2})")

                result_flag = search_and_check_policy(policy_number)

                if result_flag == 1:
                    double_click_the_policy()
                    vehicle_information()
                    star_info = pyperclip.paste().strip()

                    age, gender = driver_information()
                    print(f"→ Star: {star_info}, Age: {age}, Gender: {gender}")
                else:
                    star_info = "Not Found"
                    age = "N/A"
                    gender = "N/A"
                    print(f"⚠️ Policy {policy_number} not found.")

                # Save the result
                df.at[index, 'Star'] = star_info
                df.at[index, 'Driver Age'] = age
                df.at[index, 'Driver Gender'] = gender

                # Save after each policy
                temp_file_path = file_path.replace(".xlsx", "_in_progress.xlsx")
                df.to_excel(temp_file_path, index=False)
                print(f"💾 Progress saved to: {temp_file_path}")

                return_to_homepage()
                time.sleep(2)

        except Exception as e:
            print(f"❗ Error on policy {policy_number}: {e}")
            continue

    final_file_path = file_path.replace(".xlsx", "_completed.xlsx")
    df.to_excel(final_file_path, index=False)
    print(f"✅ Final version saved as: {final_file_path}")


if __name__ == "__main__":
    # file_path = './Intact_Loss_Report.xlsx'
    file_path = r'C:\Users\Vincent.Zhong\Documents\GitHub\Label_Star\Book12.xlsx'

    retrieve_star_info_and_update_limited(file_path, limit=None)
    # resume_star_driver_processing('./Book11_in_progress.xlsx')


