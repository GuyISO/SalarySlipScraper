from bs4 import BeautifulSoup
from pandas import DataFrame

class Slip():
    
    def __init__(self):
        
        self.key = ""
        self.contents = {}
    
    def load_from_data_frame(self, data_frame):
        self.contents = data_frame.iloc[0].to_dict()
        self.key =self.contents['明細キー']
        
    def load_from_html(self, html):
        
        # BeautifulSoupを使用してhtmlをパースする
        soup = BeautifulSoup(html, "html.parser")
        
        def add_content(key, value, changes_sign):
            """
            指定したkeyとvalueをcontentsに追加する
            """
            
            # キーか値いずれかが空白の場合は辞書に追加しない
            if key == '':
                return
            if value =='':
                return
            
            # valueの符号の反転処理
            if changes_sign:
                if value[0] == '-':
                    value = value[1:]
                else:
                    value = '-' + value
            
            # 辞書に追加
            self.contents[key] = value
        
        def get_text(tr_no, td_no):
            """
            給与明細リストが書かれたテーブルの指定位置から文字列を取得する
            """
            
            ui_link = soup.find('a', {'class': 'ui-link'})
            table = ui_link.find('table')
            tbody = table.find('tbody')
            trs = tbody.find_all('tr')
            tr = trs[tr_no]
            td = tr.find_all('td')[td_no]
            # 文字列が記入されている場合のみtd内にspanがありそこに文字列が記述されている
            if len(td.find_all('span')) == 0:
                return ""
            span = td.find('span')
            text = span.text
            
            return text
        
        # 「yyyy年mm月  XX明細書」のような、明細の名称が記載された見出しを取得
        ui_content = soup.find('div', {'class': 'ui-content'})
        table = ui_content.find('table')
        tbody = table.find('tbody')
        tr = tbody.find('tr')
        td = tr.find_all('td')[1]
        center = td.find('center')
        text = center.text
        
        # 文字列に含まれる改行とタブと空白を削除して明細キーとして扱う
        text = text.replace("\n", "")
        text = text.replace("\t", "")
        text = text.replace("\xa0", "")
        
        self.contents["明細キー"] = text
        self.key = text
        
        # 明細分類
        add_content(get_text(2, 2), get_text(2, 1), False)
        
        # 年
        add_content(get_text(2, 5), get_text(2, 4), False)
        
        # 月
        add_content(get_text(2, 7), get_text(2, 6), False)
        
        # No
        add_content(get_text(2, 8), get_text(2, 9), False)
        
        # 支給総額
        add_content(get_text(33, 1), get_text(33, 2), False)
        
        # 引去項目計
        add_content(get_text(45, 1), get_text(45, 2), True)
        
        # その他計
        add_content(get_text(55, 5), get_text(55, 6), False)
        
        # 勤怠
        offset = 1
        for tr_no in range(19, 24):
            for td_no in range(1, 10, 3):
                add_content('勤怠_' + get_text(tr_no, td_no + offset), get_text(tr_no, td_no + offset + 1), False)
            offset = 0
        
        # 支給項目
        offset = 1
        for tr_no in range(25, 32):
            for td_no in range(1, 5, 2):
                add_content('支給項目_' + get_text(tr_no, td_no + offset), get_text(tr_no, td_no + offset + 1), False)
            offset = 0
        
        # 引去項目
        offset = 1
        for tr_no in range(34, 44):
            for td_no in range(1, 5, 2):
                add_content('引去項目_' + get_text(tr_no, td_no + offset), get_text( tr_no, td_no + offset + 1), True)
            offset = 0
        
        # 現物給与等雇保対象額
        offset = 1
        for tr_no in range(46, 56):
            add_content('現物給与等雇保対象額_' + get_text(tr_no, 1 + offset), get_text(tr_no, 2 + offset), False)
            offset = 0
        
        # 現物給与等課税対象額
        offset = 2
        for tr_no in range(46, 56):
            add_content('現物給与等課税対象額_' + get_text(tr_no, 3 + offset), get_text(tr_no, 4 + offset), False)
            offset = 0
        
        # その他
        offset = 3
        for tr_no in range(46, 54):
            add_content('その他_' + get_text(tr_no, 5 + offset), get_text(tr_no, 6 + offset), False)
            offset = 0
        
        # 振込先
        for tr_no in range(58, 60):
            add_content('振込先_' + get_text(tr_no, 2), get_text(tr_no, 6), False)
        
        # 欄外
        for td_no in range(1, 4):
            text = ""
            for tr_no in range(3, 12):
                text = text + get_text(tr_no, td_no)
            text = text.replace("\n", "")
            text = text.replace("　", "")
            text = text.replace("\t", "")
            text = text.replace("\xa0", "")
            add_content('欄外' + str(td_no), text, False)
    
    def get_data_frame(self):
        
        # 作成した辞書から1行のデータフレームを作成する
        # 辞書のキーが列名になり、アイテムがindex=0に追加される、indexを指定しないとエラーが出る
        data_frame = DataFrame(self.contents, index=[0])
        
        return data_frame