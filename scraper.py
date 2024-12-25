import os
import datetime
import json
import shutil
from time import sleep
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
import pandas
from slip import Slip

FILE_NAME_SETTING = 'settings.json'
FILE_NAME_CSV = 'SalarySlip.csv'
FOLDER_NAME_BACKUP = 'backup'

file_name_backup = 'backup_' + datetime.date.today().strftime('%Y%m%d') + '.csv'
file_path_backup = FOLDER_NAME_BACKUP + '//' +file_name_backup

class SalarySlipScraper():
    
    def __init__(self):
        """
        インスタンス生成時にMicrosoftEdgeを起動し、Web明細に接続する
        """
        
        self.slips = {}
        
        print("------------------------------------------------------")
        print("  SalarySlipScraper起動")
        print("------------------------------------------------------")
        print("  設定ファイル読込")
        print("    読込完了:" + FILE_NAME_SETTING)
        
        with open(FILE_NAME_SETTING) as f:
            self.setting = json.load(f)
        
        print("------------------------------------------------------")
        print("  MicrosoftEdge起動")
        print("    環境によっては1分程度時間がかかる場合があります")
        
        self.driver = webdriver.Edge()
        
        print("  Web明細に接続")
        print("    URL:" + self.setting['url'])
        
        self.driver.get(self.setting['url'])
    
    def __del__(self):
        """
        インスタンス破棄時にMicrosoftEdgeを終了する
        """
        
        print("------------------------------------------------------")
        print("  MicrosoftEdge終了")
        
        self.driver.quit()
        
        print("  Enterを押して終了")
        input("    - Press Enter Key -")
    
    def try_auto_log_in(self):
        """
        設定ファイルに記載されたIDとPASSでログインを施行する
        """
        
        print("------------------------------------------------------")
        print("  設定ファイルのIDとPASSでログイン試行")
        
        # ログインIDを入力
        log_in_id = self.driver.find_element(By.ID, 'HID_USER_ID')
        log_in_id.send_keys(self.setting['id'])
        
        # パスワードを入力
        log_in_pass = self.driver.find_element(By.ID, 'PASSWORD')
        log_in_pass.send_keys(self.setting['pass'])
        
        # ログインボタンをクリック
        log_in_btn = self.driver.find_element(By.XPATH, "/html/body/form/div[1]/div[2]/center/div/div[1]/div/div[3]")
        log_in_btn.click()
    
    def prompt_manual_log_in(self):
        """
        手動でのログインを促す
        """
        
        print("------------------------------------------------------")
        print("  Web明細にログインしたらEnterを押してください")
        input("    - Press Enter Key - ")
    
    def displays_log_in_page(self):
        """
        ログイン画面が表示されている場合はTrue、そうでなければFalseを返す
        """
        
        try:
            
            # IDを入力する欄が取得できたらログイン画面が表示されているとする
            WebDriverWait(self.driver, 5).until(expected_conditions.presence_of_element_located((By.ID, 'HID_USER_ID')))
            
            print("------------------------------------------------------")
            print("  Web明細への接続に成功")
            
            return True
            
        except TimeoutException:
            
            print("------------------------------------------------------")
            print("  Web明細への接続に失敗")
            print("  ウェブサイトのサービス提供状態もしくは")
            print("  " + FILE_NAME_SETTING + "のURLを確認してください")
            
            return False
    
    def completes_log_in(self):
        """
        ログイン完了している場合はTrue、そうでなければFalseを返す
        """
        
        try:
            
            # 給与明細リストが書かれたテーブルを取得できたらログインに成功している
            WebDriverWait(self.driver, 5).until(expected_conditions.presence_of_element_located((By.ID, 'SEL_FIGURE_DATE')))
            
            print("------------------------------------------------------")
            print("  ログイン状態検出成功")
            
            return True
            
        except TimeoutException:
            
            print("------------------------------------------------------")
            print("  ログイン状態検出失敗")
            
            return False
    
    def close_note(self):
        """
        ログイン時にお知らせが出ている場合は閉じる
        """
        
        try:
            
            # Xボタンをクリック(存在しない場合もある？)
            close_btn = self.driver.find_element(By.XPATH, '//*[@id="popupDialog"]/div[1]/a/span')
            close_btn.click()
            
        finally:
            
            pass
    
    def scrape_slips(self):
        """
        給与明細の取得処理
        """
        
        def scrape_slip():
            
            while True:
                
                # ページの更新待ち
                sleep(0.8)
                
                # 表示されている明細を取得
                html = self.driver.page_source
                
                slip = Slip()
                slip.load_from_html(html)
                
                self.slips[slip.key] = slip
                
                print("    取得完了：" + slip.key)
                
                # ドロップダウンを選択して次の月の明細へ移動
                drop_down = self.driver.find_element(By.ID, 'SEL_FIGURE_DATE')
                select = Select(drop_down)
                
                # 一番下の明細が選択中の場合
                if select.first_selected_option == select.options[-1]:
                    
                    # リストの全明細取得完了によりループから出る
                    break
                    
                # 明細がまだある場合
                else:
                    
                    # ドロップダウンリストを下に送る
                    element = self.driver.find_element(By.XPATH, '//*[@id="tablet_menu"]/div[2]')
                    webdriver.ActionChains(self.driver).click(element).perform()
                    sleep(0.4)
                    webdriver.ActionChains(self.driver).send_keys(Keys.ARROW_DOWN).perform()
                    sleep(0.2)
                    webdriver.ActionChains(self.driver).send_keys(Keys.ENTER).perform()
        
        print("------------------------------------------------------")
        print("  明細取得開始")
        
        # 給与明細を選択(ログイン時点で選択済みだが念の為)
        kyuyo_btn = self.driver.find_element(By.XPATH, '//*[@id="tablet_menu"]/ul/li[1]')
        kyuyo_btn.click()
        
        scrape_slip()
        
        # 賞与明細を選択
        shoyo_btn = self.driver.find_element(By.XPATH, '//*[@id="tablet_menu"]/ul/li[2]')
        shoyo_btn.click()
        
        scrape_slip()
        
        print("  明細取得完了")
    
    def import_csv(self):
        """
        出力済みcsvがある場合は取り込む
        """
        
        print("------------------------------------------------------")
        print("  出力済み明細のcsvを取得")
        
        if os.path.exists(FILE_NAME_CSV):
            
            data_frame = pandas.read_csv(FILE_NAME_CSV, encoding='utf-8_sig')
            
            # 各行を1行ずつデータフレームとして処理
            for index, row in data_frame.iterrows():
                # rowを1行のデータフレームとして扱う
                single_row_df = pandas.DataFrame([row])
                
                slip = Slip()
                slip.load_from_data_frame(single_row_df)
                
                self.slips[slip.key] = slip
                
            print("    取得完了:" + FILE_NAME_CSV)
            
            # 取得したcsvは上書き出力されるため、バックアップフォルダにコピーを作成
            if not os.path.exists(FOLDER_NAME_BACKUP):
                os.mkdir(FOLDER_NAME_BACKUP)
            shutil.copy(FILE_NAME_CSV, file_path_backup)
            print("    複製完了:" + file_name_backup)
            
        else:
            
            print("    取得失敗:" + FILE_NAME_CSV)
    
    def export_csv(self):
        """
        csv出力する
        """
        
        print("------------------------------------------------------")
        print("  取得した明細をcsv出力")
        
        data_frames = []
        for slip in self.slips.values():
            data_frames.append(slip.get_data_frame())
        
        # 全明細のデータフレームを統合する
        concatenated_data_frame = pandas.concat(data_frames)
        
        # データフレームをcsv出力する
        concatenated_data_frame.to_csv(FILE_NAME_CSV, index = False, encoding='utf-8_sig')
        
        print("    出力完了:" + FILE_NAME_CSV )