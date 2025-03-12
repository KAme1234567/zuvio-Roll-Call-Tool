from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import threading
import keyboard  # 監聽 C 鍵退出

# 設定 CSV 檔案的路徑
csv_file = "Account_Information.csv"

# 檢查檔案是否存在，若不存在則創建空的 CSV 文件
try:
    df = pd.read_csv(csv_file, header=None, names=["Name", "Email", "Password"])
except FileNotFoundError:
    df = pd.DataFrame(columns=["Name", "Email", "Password"])
    df.to_csv(csv_file, index=False, header=True)

# 顯示帳號資料（最多顯示5筆）
def show_accounts():
    if df.empty:
        print("尚未有帳號")
    else:
        for i, row in df.head(5).iterrows():
            print(f"{i+1}. {row['Name']} - {row['Email']}")

# 新增帳號
def add_account():
    name = input("請輸入帳號名稱: ")
    email = input("請輸入帳號Email: ")
    password = input("請輸入帳號密碼: ")
    df.loc[len(df)] = [name, email, password]
    df.to_csv(csv_file, index=False, header=True)
    print("帳號新增成功！")

# 刪除帳號
def delete_account():
    show_accounts()
    try:
        account_index = int(input("請輸入要刪除的帳號編號（1-5）: ")) - 1
        if 0 <= account_index < len(df):
            df.drop(account_index, inplace=True)
            df.reset_index(drop=True, inplace=True)
            df.to_csv(csv_file, index=False, header=True)
            print("帳號已刪除！")
        else:
            print("無效的帳號編號！")
    except ValueError:
        print("無效的輸入！")

# 點名自動化
def auto_rollcall(driver):
    global exit_flag  # 用於監聽 `C` 鍵退出
    exit_flag = False

    # 監聽 `C` 鍵來退出
    def listen_for_exit():
        global exit_flag
        while True:
            if keyboard.is_pressed("c"):
                exit_flag = True
                print("\n手動退出點名等待...")
                break

    # 啟動監聽 `C` 鍵的線程
    exit_thread = threading.Thread(target=listen_for_exit, daemon=True)
    exit_thread.start()

    while not exit_flag:
        try:
            # 檢查按鈕是否存在
            button = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CLASS_NAME, "i-r-f-b-button"))
            )
            if "點名" in button.text:
                button.click()
                print("已成功點名！")
                break  # 點名成功後結束迴圈
        except:
            print("尚未到點名時間，重新整理頁面...")
            driver.refresh()  # 重新整理頁面
            time.sleep(4)  # 每 4 秒檢查一次

# 自動登入
def auto_login(username, password):
    driver_path = "chromedriver-win64\\chromedriver.exe"  # 修改你的 chromedriver 路徑
    options = webdriver.ChromeOptions()
    service = Service(driver_path)
    driver = webdriver.Chrome(service=service, options=options)

    login_url = "https://irs.zuvio.com.tw/student5/irs/index"
    try:
        driver.get(login_url)

        # 填入帳號密碼
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "email"))).send_keys(username)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "password"))).send_keys(password)

        # 點擊登入
        driver.find_element(By.ID, "login-btn").click()
        print("自動登入成功！")
        time.sleep(3)  # 等待頁面載入

        # 獲取所有課程
        course_elements = driver.find_elements(By.CLASS_NAME, "i-m-p-c-a-c-l-c-b-t-course-name")
        course_ids = driver.find_elements(By.CLASS_NAME, "i-m-p-c-a-c-l-course-box")

        courses = []
        for idx, course_element in enumerate(course_elements):
            course_name = course_element.text
            course_id = course_ids[idx].get_attribute("data-course-id")
            courses.append((course_name, course_id))

        # 顯示所有課程
        if courses:
            print("\n所有可選課程：")
            for i, (course_name, course_id) in enumerate(courses):
                print(f"{i+1}. {course_name}")

            # 讓使用者選擇課程
            try:
                choice = int(input(f"請輸入1~{len(courses)}選擇課程："))
                if 1 <= choice <= len(courses):
                    selected_course_name, selected_course_id = courses[choice-1]
                    print(f"您選擇的課程是: {selected_course_name}")

                    # 進入點名
                    rollcall_url = f"https://irs.zuvio.com.tw/student5/irs/rollcall/{selected_course_id}"
                    driver.get(rollcall_url)
                    print(f"正在進入點名頁面: {rollcall_url}")

                    # 開始點名
                    auto_rollcall(driver)
                else:
                    print("無效選擇！")
            except ValueError:
                print("請輸入數字！")
        else:
            print("目前沒有可選課程")
    finally:
        driver.quit()  # 關閉瀏覽器

# 主程式
def main():
    while True:
        print("\n選擇操作:")
        print("1. 顯示帳號資料")
        print("2. 登入選擇帳號")
        print("D. 刪除帳號")
        print("S. 新增帳號")
        print("E. 結束程式")

        choice = input("請選擇操作: ").strip().upper()

        if choice == "1":
            show_accounts()
        elif choice == "2":
            if df.empty:
                print("尚未有帳號！")
            else:
                show_accounts()
                try:
                    account_index = int(input("請選擇登入帳號編號 (1-5): ")) - 1
                    if 0 <= account_index < len(df):
                        selected_account = df.iloc[account_index]
                        print(f"登入帳號: {selected_account['Name']} ({selected_account['Email']})")
                        auto_login(selected_account['Email'], selected_account['Password'])
                    else:
                        print("無效帳號編號！")
                except ValueError:
                    print("無效輸入！")
        elif choice == "D":
            delete_account()
        elif choice == "S":
            add_account()
        elif choice == "E":
            print("結束程式")
            break
        else:
            print("無效選擇！")

if __name__ == "__main__":
    main()
