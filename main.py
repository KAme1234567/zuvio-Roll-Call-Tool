from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd

# 設定CSV檔案的路徑
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
        # 顯示最多前5個帳號
        for i, row in df.head(5).iterrows():
            print(f"{i+1}. {row['Name']} - {row['Email']}")

# 新增帳號
def add_account():
    name = input("請輸入帳號名稱: ")
    email = input("請輸入帳號Email: ")
    password = input("請輸入帳號密碼: ")

    # 新帳號加入 DataFrame
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

# 開啟瀏覽器進行自動登入
def auto_login(username, password):
    driver_path = "chromedriver-win64\\chromedriver.exe"  # 修改為你的 chromedriver 路徑
    options = webdriver.ChromeOptions()
    service = Service(driver_path)
    driver = webdriver.Chrome(service=service, options=options)

    login_url = "https://irs.zuvio.com.tw/student5/irs/index"  # 替換成你的網站登入頁面
    try:
        # 打開登入頁面
        driver.get(login_url)

        # 等待並填入帳號密碼
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "email"))).send_keys(username)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "password"))).send_keys(password)

        # 點擊登入按鈕
        driver.find_element(By.ID, "login-btn").click()  # 使用 id 找到登入按鈕並點擊
        print("自動登入成功！")

        # 確保課程頁面已經完全加載
        time.sleep(3)  # 等待 3 秒

        # 抓取所有課程名稱與對應的課程 ID
        course_elements = driver.find_elements(By.CLASS_NAME, "i-m-p-c-a-c-l-c-b-t-course-name")
        course_ids = driver.find_elements(By.CLASS_NAME, "i-m-p-c-a-c-l-course-box")

        courses = []
        for idx, course_element in enumerate(course_elements):
            course_name = course_element.text
            course_id = course_ids[idx].get_attribute("data-course-id")  # 取得 data-course-id
            courses.append((course_name, course_id))

        # 顯示所有課程名稱
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
                    # 進入點名頁面
                    rollcall_url = f"https://irs.zuvio.com.tw/student5/irs/rollcall/{selected_course_id}"
                    print(f"正在進入點名頁面: {rollcall_url}")
                    driver.get(rollcall_url)  # 跳轉到對應的點名頁面
                    while True:
                        try:
                            # 檢查按鈕是否存在
                            button = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located((By.CLASS_NAME, "i-r-f-b-button"))
                            )
                            if button.text == "點名":
                                button.click()
                                print("已成功點名！")
                                break  # 點名成功後結束迴圈
                        except:
                            print("尚未到點名時間，重新整理頁面...")
                            driver.refresh()  # 重新整理頁面
                            time.sleep(4)  # 每 10 秒檢查一次

                else:
                    print("無效的選擇，請重新選擇！")
            except ValueError:
                print("請輸入正確的數字！")
        else:
            print("目前沒有可選課程")
    finally:
        driver.quit()  # 關閉瀏覽器

# 主程式循環
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
                    account_index = int(input("請選擇要登入的帳號編號 (1-5): ")) - 1
                    if 0 <= account_index < len(df):
                        selected_account = df.iloc[account_index]
                        print(f"正在登入帳號: {selected_account['Name']} ({selected_account['Email']})")
                        # 自動登入
                        auto_login(selected_account['Email'], selected_account['Password'])
                    else:
                        print("無效的帳號編號！")
                except ValueError:
                    print("無效的輸入！")
        elif choice == "D":
            delete_account()
        elif choice == "S":
            add_account()
        elif choice == "E":
            print("結束程式")
            break
        else:
            print("無效的選項，請重新選擇！")

if __name__ == "__main__":
    main()
