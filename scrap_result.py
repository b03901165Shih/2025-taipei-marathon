import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

# --------------------------------------------------
# Selenium / 瀏覽器設定
# --------------------------------------------------

BASE_URL = "https://www.bravelog.tw/contest/rank/2025122101"

chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument(
    "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
)


def setup_driver() -> webdriver.Chrome:
    """建立並回傳一個已設定好的 Chrome WebDriver。"""
    driver = webdriver.Chrome(options=chrome_options)
    driver.implicitly_wait(5)
    return driver


# --------------------------------------------------
# 賽事類型與分組處理
# --------------------------------------------------

# 賽事類型：全馬(MA) 和 半馬(HM)
RACE_TYPES = [
    ("半馬", "HM"),
    ("全馬", "MA"),
]

# 預設分組列表（如果無法動態獲取時使用）
DEFAULT_GROUP_NAMES = [
    "男國際選手",
    "女國際選手",
    "女50-59歲",
    "女30-39歲",
    "女40-49歲",
    "女20-29歲",
    "女60歲+",
    "男40-49歲",
    "男30-39歲",
    "男20-29歲",
    "男50-59歲",
    "男60歲+",
    "女19歲-",
    "女視障選手",
    "男19歲-",
    "男視障選手",
]


def get_available_groups(driver: webdriver.Chrome) -> list:
    """
    從當前頁面上動態獲取所有可用的分組選項。
    回傳分組名稱的列表。
    """
    try:
        wait = WebDriverWait(driver, 10)
        
        # 先等所有 nice-select 都出現
        selects = wait.until(
            EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "div.nice-select")
            )
        )
        
        groups = []
        
        # 掃描所有下拉選單，找到分組選單（通常會包含「年齡分組」或類似的選項）
        for select_root in selects:
            try:
                # 確保打開選單（使用 JavaScript 點擊避免元素攔截）
                cls = select_root.get_attribute("class") or ""
                if "open" not in cls:
                    driver.execute_script("arguments[0].click();", select_root)
                    time.sleep(0.5)  # 簡單等待選單展開
                
                # 獲取這個選單中的所有選項
                options = select_root.find_elements(By.CSS_SELECTOR, "li.option")
                
                # 如果沒有選項，跳過這個選單
                if not options:
                    if "open" in (select_root.get_attribute("class") or ""):
                        driver.execute_script("arguments[0].click();", select_root)
                    continue
                
                # 檢查這個選單是否看起來像分組選單（排除賽事類型選單）
                option_texts = []
                is_race_type_menu = False
                
                for opt in options:
                    text = opt.text.strip()
                    data_value = opt.get_attribute("data-value") or ""
                    
                    # 如果 data-value 是 4 位數字（1670/1671），這是賽事類型選單
                    if data_value.isdigit() and len(data_value) == 4:
                        is_race_type_menu = True
                        break
                    
                    # 如果文字是 MA 或 HM，這也是賽事類型選單
                    if text in ("MA", "HM", "項目"):
                        is_race_type_menu = True
                        break
                    
                    # 跳過空選項和「年齡分組」標題（但保留其他分組）
                    if not text or text == "年齡分組":
                        continue
                    
                    option_texts.append(text)
                
                # 如果是賽事類型選單，跳過
                if is_race_type_menu:
                    if "open" in (select_root.get_attribute("class") or ""):
                        driver.execute_script("arguments[0].click();", select_root)
                    continue
                
                # 如果找到多個分組選項（通常分組選單會有 5+ 個選項），就使用這個
                if len(option_texts) >= 5:
                    groups = option_texts
                    # 關閉選單
                    if "open" in (select_root.get_attribute("class") or ""):
                        driver.execute_script("arguments[0].click();", select_root)
                    break
                
                # 關閉選單
                if "open" in (select_root.get_attribute("class") or ""):
                    driver.execute_script("arguments[0].click();", select_root)
                    
            except Exception:
                continue
        
        if groups:
            print(f"✅ 動態獲取到 {len(groups)} 個分組：{groups[:5]}..." if len(groups) > 5 else f"✅ 動態獲取到 {len(groups)} 個分組：{groups}")
            return groups
        else:
            print("⚠️ 無法動態獲取分組，使用預設分組列表")
            return DEFAULT_GROUP_NAMES
            
    except Exception as e:
        print(f"⚠️ 獲取分組時發生錯誤: {e}，使用預設分組列表")
        return DEFAULT_GROUP_NAMES


def switch_race_type(driver: webdriver.Chrome, race_type_value: str) -> bool:
    """
    切換賽事類型（全馬/半馬）。
    race_type_value: "MA" 或 "HM"
    依照實際 HTML：
        <div class="nice-select chosen-select">
            ...
            <ul class="list">
                <li data-value="1670" class="option selected">MA</li>
                <li data-value="1671" class="option">HM</li>
            </ul>
        </div>
    回傳是否成功切換。
    """
    try:
        wait = WebDriverWait(driver, 10)

        # 只鎖定賽事項目的 nice-select.chosen-select
        selects = wait.until(
            EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "div.nice-select.chosen-select")
            )
        )

        for select_root in selects:
            try:
                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center'});", select_root
                )

                # 確保打開選單（使用 JavaScript 點擊避免元素攔截）
                cls = select_root.get_attribute("class") or ""
                if "open" not in cls:
                    # 使用 JavaScript 直接點擊，繞過元素攔截問題
                    driver.execute_script("arguments[0].click();", select_root)
                    time.sleep(0.3)  # 增加等待時間確保選單完全打開

                # 先檢查這個選單是否包含 MA 和 HM（確認是賽事類型選單）
                all_options = select_root.find_elements(By.CSS_SELECTOR, "li.option")
                has_ma = False
                has_hm = False
                for opt in all_options:
                    text = opt.text.strip()
                    if text == "MA":
                        has_ma = True
                    if text == "HM":
                        has_hm = True
                
                # 如果不是賽事類型選單（沒有 MA 和 HM），跳過
                if not (has_ma and has_hm):
                    if "open" in (select_root.get_attribute("class") or ""):
                        # 使用 JavaScript 點擊關閉選單
                        driver.execute_script("arguments[0].click();", select_root)
                    continue

                # 在這個 select 範圍內找對應的 option（文字是 MA / HM）
                options = select_root.find_elements(
                    By.XPATH,
                    f".//li[@class='option' and normalize-space(text())='{race_type_value}']",
                )

                if not options:
                    # 沒有這個選項，關掉選單換下一個 select
                    if "open" in (select_root.get_attribute("class") or ""):
                        # 使用 JavaScript 點擊關閉選單
                        driver.execute_script("arguments[0].click();", select_root)
                    continue

                option = options[0]
                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center'});", option
                )
                # 使用 JavaScript 直接點擊，繞過元素攔截問題
                driver.execute_script("arguments[0].click();", option)

                # 簡單等待頁面載入
                time.sleep(2)
                return True
            except Exception as e:
                # 這個 select 失敗就試下一個
                continue

        print(f"⚠️ 找不到賽事類型「{race_type_value}」的選項")
        return False
    except Exception as e:
        print(f"⚠️ 切換賽事類型「{race_type_value}」時發生錯誤: {e}")
        return False


def click_category_tab(driver: webdriver.Chrome, category_name: str) -> bool:
    """
    在頁面上的 nice-select 下拉選單中選取指定分組。
    回傳是否成功點擊。
    """
    try:
        wait = WebDriverWait(driver, 10)

        # 先等所有 nice-select 都出現，再一個一個掃描，看哪一個裡面有目標 data-value
        selects = wait.until(
            EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "div.nice-select")
            )
        )

        for select_root in selects:
            try:
                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center'});", select_root
                )

                # 確保打開選單（使用 JavaScript 點擊避免元素攔截）
                cls = select_root.get_attribute("class") or ""
                if "open" not in cls:
                    driver.execute_script("arguments[0].click();", select_root)
                    time.sleep(0.3)

                # 僅在這個 select 範圍內找對應的 option
                options = select_root.find_elements(
                    By.XPATH,
                    f".//li[@class='option' and @data-value='{category_name}']",
                )
                if not options:
                    # 沒有這個分組，關掉選單換下一個 select
                    if "open" in (select_root.get_attribute("class") or ""):
                        driver.execute_script("arguments[0].click();", select_root)
                    continue

                option = options[0]
                driver.execute_script(
                    "arguments[0].scrollIntoView({block: 'center'});", option
                )
                # 使用 JavaScript 直接點擊，繞過元素攔截問題
                driver.execute_script("arguments[0].click();", option)

                # 簡單等待頁面載入
                time.sleep(2)
                return True
            except Exception:
                # 這個 select 失敗就試下一個
                continue

        print(f"⚠️ 頁面上的所有下拉選單都找不到分組「{category_name}」")
        return False
    except Exception as e:
        print(f"⚠️ 找不到或無法點擊分組「{category_name}」: {e}")
        return False


# --------------------------------------------------
# 成績表解析
# --------------------------------------------------

def parse_time_to_timedelta(time_str: str):
    """將 hh:mm:ss 或 mm:ss 轉成 pandas Timedelta，錯誤則回傳 NaT。"""
    if not time_str or time_str in ("N/A", "-", "--"):
        return pd.NaT
    t = time_str.strip()
    try:
        parts = t.split(":")
        if len(parts) == 2:
            # mm:ss -> 0:mm:ss
            parts = ["0"] + parts
        if len(parts) != 3:
            return pd.NaT
        h, m, s = map(int, parts)
        return pd.to_timedelta(f"{h:02d}:{m:02d}:{s:02d}")
    except Exception:
        return pd.NaT


def scrape_current_table(driver: webdriver.Chrome, category_name: str, race_type_name: str = ""):
    """
    在當前已顯示該分組的頁面上，解析成績卡片列表。
    依照你提供的 HTML 結構，成績每一筆大致為：
    <div class="fl-wrap list-single-main-item_content">
        <div class="list-item">
            <div class="list-user-info">
                <div class="name">姓名</div>
                <div class="detail-info">
                    <span>背號</span>
                    <span>MA/HM 等賽別</span>
                    <span>分組名稱</span>
                </div>
            </div>
            <div class="time"><span>完賽時間</span></div>
        </div>
    race_type_name: 賽事類型名稱（"全馬" 或 "半馬"），用於標記資料來源
    """
    wait = WebDriverWait(driver, 10)

    # 等待至少一個成績卡片出現
    try:
        wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "div.fl-wrap.list-single-main-item_content")
            )
        )
    except Exception:
        print(f"⚠️ 分組「{category_name}」找不到成績區塊")
        return []

    soup = BeautifulSoup(driver.page_source, "html.parser")
    cards = soup.select("div.fl-wrap.list-single-main-item_content")

    results = []
    for card in cards:
        # 姓名
        name_el = card.select_one(".list-user-info .name")
        name = name_el.get_text(strip=True) if name_el else ""

        # 背號、賽別、分組
        spans = card.select(".list-user-info .detail-info span")
        bib = spans[0].get_text(strip=True) if len(spans) >= 1 else ""
        race_type = spans[1].get_text(strip=True) if len(spans) >= 2 else ""  # MA / HM
        group_text = spans[2].get_text(strip=True) if len(spans) >= 3 else category_name

        # 完賽時間
        time_el = card.select_one(".time span")
        finish_time = time_el.get_text(strip=True) if time_el else ""

        # 沒有名字或背號就略過（通常是異常卡片）
        if not name and not bib:
            continue

        results.append(
            {
                "姓名": name,
                "背號": bib,
                "賽別": race_type,
                "賽事類型": race_type_name,  # 全馬/半馬
                "分組": group_text,
                "完賽時間": finish_time,
                "來源分組標籤": category_name,
            }
        )

    print(f"「{category_name}」解析到 {len(results)} 筆")
    return results


def scrape_category(driver: webdriver.Chrome, category_name: str, race_type_name: str = ""):
    """切換到指定分組並抓取該分組「所有頁數」的成績。"""
    print(f"=== 處理分組：{category_name} ===")
    ok = click_category_tab(driver, category_name)
    if not ok:
        return []

    all_results = []

    page_count = 0
    max_pages = 10000  # 安全上限，避免無限循環
    
    while page_count < max_pages:
        page_count += 1
        
        # 先抓目前頁面的所有卡片
        page_results = scrape_current_table(driver, category_name, race_type_name)
        all_results.extend(page_results)

        # 嘗試找到分頁區塊（每次都重新獲取，因為頁面更新後元素可能失效）
        try:
            pagination = driver.find_element(By.ID, "pagination")
        except Exception:
            # 沒有分頁區塊，表示只有一頁
            print(f"分組「{category_name}」沒有分頁區塊，結束")
            break

        try:
            current_page = int(pagination.get_attribute("data-page") or "1")
            total_pages = int(pagination.get_attribute("data-total") or "1")
        except Exception as e:
            # 取不到 page / total 就不要勉強翻頁
            print(f"⚠️ 分組「{category_name}」無法讀取頁數資訊: {e}")
            break

        # 顯示目前頁數資訊
        print(f"分組「{category_name}」目前在第 {current_page} / {total_pages} 頁（本頁爬取 {len(page_results)} 筆）")

        # 已經是最後一頁了，就結束這個分組
        if current_page >= total_pages:
            print(f"✅ 分組「{category_name}」已到最後一頁")
            break

        # 找「下一頁」按鈕（右箭頭），且不能是 disabled
        try:
            # 重新獲取 pagination 元素（頁面可能已更新）
            pagination = driver.find_element(By.ID, "pagination")
            next_btn = pagination.find_element(
                By.CSS_SELECTOR, "li.nextposts-link:not(.disabled) a.page-link"
            )
        except Exception as e:
            # 找不到可用的下一頁按鈕，就停在這一頁
            print(f"⚠️ 分組「{category_name}」找不到可用的下一頁按鈕: {e}")
            break

        # 滾動並點擊下一頁
        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", next_btn
            )
            next_btn.click()
            
            # 簡單等待頁面載入
            time.sleep(2)
        except Exception as e:
            print(f"⚠️ 分組「{category_name}」翻頁時發生錯誤: {e}")
            break

    print(f"=== 分組「{category_name}」累計 {len(all_results)} 筆 ===")
    return all_results


# --------------------------------------------------
# 主流程
# --------------------------------------------------

def main():
    driver = setup_driver()

    try:
        print("開啟成績頁面…")
        driver.get(BASE_URL)
        time.sleep(3)

        all_results = []

        # 循環處理每種賽事類型（全馬、半馬）
        for race_type_name, race_type_value in RACE_TYPES:
            print(f"\n{'='*50}")
            print(f"開始處理賽事類型：{race_type_name} ({race_type_value})")
            print(f"{'='*50}\n")
            
            # 切換到對應的賽事類型
            if not switch_race_type(driver, race_type_value):
                print(f"⚠️ 無法切換到「{race_type_name}」，跳過此賽事類型")
                continue
            
            # 簡單等待頁面載入
            time.sleep(3)
            
            # 動態獲取當前賽事類型下可用的分組列表
            available_groups = get_available_groups(driver)
            
            if not available_groups:
                print(f"⚠️ 「{race_type_name}」沒有可用分組，跳過")
                continue
            
            # 爬取這個賽事類型下的所有分組
            for cat in available_groups:
                data = scrape_category(driver, cat, race_type_name)
                all_results.extend(data)
                # 避免太頻繁操作
                time.sleep(1)

        df = pd.DataFrame(all_results)
        if df.empty:
            print("⚠️ 最後沒有抓到任何成績資料，請檢查 selector 或頁面結構。")
            return

        # 轉換完賽時間為 Timedelta 並排序
        df["完賽時間_td"] = df["完賽時間"].apply(parse_time_to_timedelta)
        df = df.sort_values(["完賽時間_td", "分組", "姓名"], na_position="last")

        # 加一個整體排名欄位
        df["總排名"] = range(1, len(df) + 1)

        # 儲存到 Excel
        output_file = "2025_台北馬拉松_完整成績.xlsx"
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="完整成績", index=False)

            # 分組統計（只在確定欄位存在時進行）
            if "分組" in df.columns:
                group_stats = (
                    df.groupby("分組")
                    .agg(
                        完賽人數=("姓名", "count"),
                        最快時間=("完賽時間_td", "min"),
                        最慢時間=("完賽時間_td", "max"),
                    )
                )
                group_stats.to_excel(writer, sheet_name="分組統計")
            
            # 按賽事類型統計
            if "賽事類型" in df.columns:
                race_type_stats = (
                    df.groupby("賽事類型")
                    .agg(
                        完賽人數=("姓名", "count"),
                        最快時間=("完賽時間_td", "min"),
                        最慢時間=("完賽時間_td", "max"),
                    )
                )
                race_type_stats.to_excel(writer, sheet_name="賽事類型統計")
                
                # 按賽事類型+分組統計
                race_group_stats = (
                    df.groupby(["賽事類型", "分組"])
                    .agg(
                        完賽人數=("姓名", "count"),
                        最快時間=("完賽時間_td", "min"),
                        最慢時間=("完賽時間_td", "max"),
                    )
                )
                race_group_stats.to_excel(writer, sheet_name="賽事類型_分組統計")

        print(f"✅ 完成！共爬取 {len(df)} 筆成績，已儲存至 {output_file}")
        if "分組" in df.columns:
            print("\n各分組筆數：")
            print(df["分組"].value_counts())
        if "賽事類型" in df.columns:
            print("\n各賽事類型筆數：")
            print(df["賽事類型"].value_counts())

    except Exception as e:
        print(f"❌ 執行過程發生錯誤: {e}")

    finally:
        driver.quit()


if __name__ == "__main__":
    # 需要套件：pip install selenium beautifulsoup4 pandas openpyxl
    main()
