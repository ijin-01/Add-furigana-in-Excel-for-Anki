import re, jaconv, sys, os, csv, builtins, pandas as pd, platform, subprocess

from fugashi import Tagger
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *

from openpyxl import *

# 후리가나 붙이기 ------------------------------------------------------------------
CJK_Unified_Ideographs = r'[\u4E00-\u9FFF]'

def add_furigana_with_fugashi(text, exclude_text='', kana_mode='hiragana'):
    """
    text: 후리가나를 붙이려는 대상 텍스트 (B 문자열)
    exclude_text: 후리가나를 붙이지 않으려는 한자를 포함하는 텍스트 (A 문자열)
    """

    # 1. A 문자열에서 CJK 범위의 한자를 추출하여 set에 저장
    excluded_kanji_set = set(ch for ch in exclude_text if re.search(CJK_Unified_Ideographs, ch))

    tagger = Tagger()
    tokens = tagger(text)

    result = []
    for token in tokens:
        surface = token.surface
        # 2. 현재 토큰이 A 문자열의 한자를 하나라도 포함하면 후리가나 생략
        #    (any()로 한 글자라도 포함하면 True)
        if any(ch in excluded_kanji_set for ch in surface):
            result.append(surface)
        else:
            # 3. 그 외 일반적인 경우에만 후리가나를 부착
            if bool(re.compile(CJK_Unified_Ideographs).search(surface)):
                kana = token.feature.kana
                if kana:
                    if kana_mode == 'katakana':
                        result.append(f" {surface}[{kana}]")
                    elif kana_mode == 'hiragana':
                        result.append(f" {surface}[{jaconv.kata2hira(kana)}]")
                else:
                    # 형태소 분석 결과 kana 정보가 없는 경우는 그대로 출력
                    result.append(surface)
            else:
                # CJK 범위 이외(히라가나, 가타카나, 알파벳 등)는 그냥 이어붙임
                result.append(surface)

    # 앞쪽에 공백이 있을 수 있으므로 정리
    msg = "".join(result)
    if msg.startswith(' '):
        msg = msg[1:]
    return msg
'''
japanese_text = "詐欺に遭い憤った被害者達が会社を相手に抗議活動を行った。あの人が言うと、褒め言葉も嫌味に聞こえる。"
print(add_furigana_with_fugashi(japanese_text))
'''

# UI 만들기 ------------------------------------------------------------------------
def column_to_number(column_name):
    column_number = 0
    for i, char in enumerate(reversed(column_name.upper())):
        column_number += (ord(char) - ord('A') + 1) * (26 ** i)
    return column_number

def number_to_column(column_number):
    """숫자를 엑셀 열 이름으로 변환"""
    column_name = []
    while column_number > 0:
        column_number -= 1  # 1을 빼서 0부터 시작하도록 조정
        column_name.append(chr(column_number % 26 + ord('A')))
        column_number //= 26
    return ''.join(reversed(column_name))

def check_file_is_open(file_path):
    if platform.system() == "Windows":
        try:
            with builtins.open(file_path, 'a'):
                return False
        except IOError:
            return True
    else:
        try:
            result = subprocess.run(['lsof', file_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            return bool(result.stdout)
        except FileNotFoundError:
            #print("lsof 명령어를 사용할 수 없습니다.")
            return False

class AutoResizingLineEdit(QLineEdit):
    def __init__(self):
        super().__init__()
        self.min_width = 100
        self.setFixedWidth(self.min_width)
        self.textChanged.connect(self.adjust_size)  # 텍스트 변경 시 크기 조정

    def adjust_size(self):
        # 텍스트의 폭 계산
        font_metrics = QFontMetrics(self.font())
        text_width = font_metrics.horizontalAdvance(self.text()) + 20  # 추가 여백
        self.setFixedWidth(max(self.min_width, text_width))  # 최소 폭 100 설정
        
class Thread(QThread):
    fault_message = ''
    fault_signal = pyqtSignal(int)
    continue_signal = pyqtSignal()

    def __init__(self, parent, filepath, columns):
        super().__init__()
        self.parent = parent
        self.filepath = filepath
        self.columns = columns
        
    def get_multiple_columns_with_rows(self, columns):
        result = {}

        try:
            if self.filepath.endswith('.xlsx') or self.filepath.endswith('.xlsm'):
                self.workbook = load_workbook(self.filepath)
                self.sheet = self.workbook.active

                for column_letter in columns:
                    result[column_letter] = []
                    for cell in self.sheet[column_letter]:
                        if cell.value is not None:
                            result[column_letter].append((cell.row, str(cell.value)))
                return result
            
            elif self.filepath.endswith('.csv'):
                columns_list = []
                for x in columns:
                    columns_list.append(column_to_number(x))

                column_headers = []
                for x in range(max(columns_list)):
                    column_headers.append(number_to_column(x))

                with builtins.open(self.filepath, 'r', encoding='utf-8') as file:
                    csv_reader = csv.DictReader(file, fieldnames=column_headers)
                    
                    # CSV 데이터를 리스트로 저장
                    data = list(csv_reader)

                    # 열별 데이터 출력
                    for x in columns:
                        column_data = []
                        for y in range(len(data)):
                            value = data[y].get(number_to_column(column_to_number(x)-1))
                            if value is not None and value.strip() != '':
                                column_data.append((y+1, value.replace('\ufeff','')))

                        result[x] = column_data
                return result

            else:
                self.fault_signal.emit(4)
                self.fault_message = 'ファイルの形式が間違っています。'
        except:
            self.fault_signal.emit(2)
            self.fault_message = 'ファイルを開けません。\nファイルの経路や名前を確認してください。'

    def run(self):
        self.input_columns_array = [item.strip() for item in self.columns.split(",")]

        self.output_columns_array = [
            number_to_column(column_to_number(item.strip()) + 1)
            for item in self.columns.split(",")
        ]
        Existing_data = self.get_multiple_columns_with_rows(self.output_columns_array)

        Existed_column_list = []
        for x in self.output_columns_array:
            if Existing_data[x]:
                for y in Existing_data[x]:
                    if y[1] != '':
                        if x not in Existed_column_list:
                            Existed_column_list.append(x)

        if Existed_column_list:
            # 열이 겹치는 경우 시그널로 메시지 전송
            self.fault_signal.emit(1)
            self.fault_message = '出力しようとする'+'、'.join(Existed_column_list) + '列に既にデータがあります。進めますか？'
        else:
            self.continue_process()

    def continue_process(self):
        if self.filepath.endswith('.xlsx') or self.filepath.endswith('.xlsm'):
            for x in self.input_columns_array:
                for y in self.get_multiple_columns_with_rows(x)[x]:
                    self.sheet[number_to_column(column_to_number(x)+1)+str(y[0])] = add_furigana_with_fugashi(y[1], kana_mode=self.parent.kana_mode)
                self.workbook.save(self.filepath)

        elif self.filepath.endswith('.csv'):
            df = pd.read_csv(self.filepath, encoding='utf-8', header=None, dtype=str)

            columns_list = []
            for x in self.output_columns_array:
                columns_list.append(column_to_number(x))
            required_cols = max(columns_list)
            
            while len(df.columns) <= required_cols:
                df[f'new_col_{len(df.columns)}'] = None

            for x in self.input_columns_array:
                for y in self.get_multiple_columns_with_rows(x)[x]:
                    df.iloc[y[0]-1, column_to_number(x)] = add_furigana_with_fugashi(y[1],kana_mode=self.parent.kana_mode)
            df.to_csv(self.filepath, encoding='utf-8-sig', index=False, header=None)
        else:
            self.fault_signal.emit(4)
            self.fault_message = 'ファイルの形式が間違っています。'
            
        self.fault_message = '完了しました。'
        self.fault_signal.emit(3)
        
class MainWindow(QWidget):
    window_width = 450
    window_height = 200
    kana_mode = 'hiragana'

    def __init__(self):
        super(MainWindow, self).__init__()
        self.columns_array = []
        self.response = None
        
        self.alert_start = 0
        self.alert_columns_is_null = 0
        self.msg_columns_is_null = '単語がある列を入力してください。'
        self.alert_columns_contains_null = 0
        self.msg_columns_contains_null = '入力した列に空白があります。'
        self.alert_column_out_of_range = 0
        self.msg_column_out_of_range = '列の範囲が外れました。範囲はAからXFDまでです。'
        self.alert_column_overlap = 0
        self.msg_column_overlap = '選択した列と出力する列が重なります。'
        
        self.initUI()

    def check_contains_strange(self, array):
        for item in array:
            if item is None or item.strip() == '':
                return 1
        
        for x in range(len(array)):
            if not column_to_number('A') <= column_to_number(array[x]) <= column_to_number('XFD'):
                return 2
            elif len(array) != 1:
                for y in range(len(array)):
                    if abs(column_to_number(array[x]) - column_to_number(array[y])) == 1:
                        return 3
        return 0
    
    def check_columns_text(self, text):
        if text == '':
            self.alert_columns_is_null = 1
            if self.alert_start == 1:
                self.label_alert.setText(self.msg_columns_is_null)
        else:
            self.alert_start = 1
            self.alert_columns_is_null = 0
            self.columns_array = [item.strip() for item in text.split(',')]

            alert = self.check_contains_strange(self.columns_array)
            if alert == 1:
                self.label_alert.setText(self.msg_columns_contains_null)
                self.alert_columns_contains_null = 1
            elif alert == 2:
                self.label_alert.setText(self.msg_column_out_of_range)
                self.alert_column_out_of_range = 1
            elif alert == 3:
                self.label_alert.setText(self.msg_column_overlap)
                self.alert_column_overlap = 1
            else:
                self.label_alert.setText('')
                self.alert_columns_contains_null = 0
                self.alert_column_out_of_range = 0
                self.alert_column_overlap = 0
        
    def SelctFilePath(self):
        filepath = QFileDialog.getOpenFileName(self, 'ファイル選択', '', 'Excel Files (*.xlsx *.xlsm *.csv)')
        self.qle_file_path.setText(filepath[0])

    def Start(self):
        old_alert_start = self.alert_start
        self.alert_start = 1
        if old_alert_start == 0 and self.alert_start == 1:
            self.label_alert.setText(self.msg_columns_is_null)

        filepath = self.qle_file_path.text()
        
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle('警告')
        msg_box.addButton(QMessageBox.StandardButton.Yes).setText('はい')

        if self.label_alert.text() != '':
            msg_box.setText(self.label_alert.text())
            msg_box.exec()
        elif filepath == '':
            msg_box.setText('ファイルの位置を確認してください。')
            msg_box.exec()
            
        elif check_file_is_open(filepath):
            msg_box.setText('ファイルが開けています。閉じてください。')
            msg_box.exec()

        else:
            if os.path.exists(filepath):
                self.th = Thread(self, filepath, self.column_input.text())
                self.th.fault_signal.connect(self.show_message_box)
                self.th.continue_signal.connect(self.th.continue_process)
                self.th.start()
            else:
                msg_box.setText('ファイルを開けません。\nファイルの経路や名前を確認してください。')
                msg_box.exec()


    def show_message_box(self, signal):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle('警告')
        
        if signal == 1:
            msg_box.addButton(QMessageBox.StandardButton.No).setText('いいえ')
        elif signal in (2, 4):
            self.th.terminate()
        
        msg_box.setText(self.th.fault_message)
        msg_box.addButton(QMessageBox.StandardButton.Yes).setText('はい')

        response = msg_box.exec()
        if response == QMessageBox.StandardButton.Yes and signal == 1:  # signal이 1이면서 경고창에서 YES를 클릭한 경우
            self.th.continue_signal.emit()
        elif response == QMessageBox.StandardButton.No and signal == 1:
            self.th.terminate()

    def get_selected_value(self):
        if self.hiragana_btn.isChecked():
            self.kana_mode = 'hiragana'
        elif self.katakana_btn.isChecked():
            self.kana_mode = 'katakana'

    def initUI(self):
        self.setWindowTitle('Add furigana in Excel for Anki')
        self.setMinimumWidth(self.window_width)
        self.setMinimumHeight(self.window_height)
        self.resize(self.window_width, self.window_height)
        
        label_column_input = QLabel('単語がある列のアルファベットを入力してください。\n'
                                    'フリガナは単語がある列の右の列に出力されます。\n'
                                    'コンマ「,」を書いて多重入力もできます。例）A,C,E',self)
        self.hiragana_btn = QRadioButton('平仮名出力', self)
        self.hiragana_btn.setChecked(True)
        self.hiragana_btn.clicked.connect(self.get_selected_value)
        self.katakana_btn = QRadioButton('片仮名出力', self)
        self.katakana_btn.clicked.connect(self.get_selected_value)

        self.column_input = AutoResizingLineEdit()
        self.label_alert = QLabel('', self)
        self.label_alert.setStyleSheet('color: red;')
        
        self.column_input.setText('b,d')
        regex = QRegularExpression('^[a-zA-Z,]*$')  # 영문자와 쉼표만 허용
        validator = QRegularExpressionValidator(regex, self.column_input)
        self.column_input.setValidator(validator)
        
        self.check_columns_text(self.column_input.text())
        self.column_input.textChanged.connect(self.check_columns_text)
        
        label_file_path = QLabel('ファイル位置', self)
        self.qle_file_path = QLineEdit(self)
        btn_file_path_select = QPushButton('...', self)
        btn_file_path_select.clicked.connect(self.SelctFilePath)

        label_start = QLabel('始める前にエクセルを閉じてください。',self)
        btn_start = QPushButton('始め', self)
        btn_start.clicked.connect(self.Start)

        kana_btn_layout = QVBoxLayout()
        kana_btn_layout.addWidget(self.hiragana_btn)
        kana_btn_layout.addWidget(self.katakana_btn)

        label_n_btn_layout = QHBoxLayout()
        label_n_btn_layout.addWidget(label_column_input)
        label_n_btn_layout.addStretch(1)
        label_n_btn_layout.addLayout(kana_btn_layout)

        file_path_layout = QHBoxLayout()
        file_path_layout.addWidget(label_file_path)
        file_path_layout.addWidget(self.qle_file_path)
        file_path_layout.addWidget(btn_file_path_select)

        start_btn_layout = QHBoxLayout()
        start_btn_layout.addStretch(1)
        start_btn_layout.addWidget(label_start)
        start_btn_layout.addWidget(btn_start)

        vbox = QVBoxLayout()
        vbox.addStretch(1)
        vbox.addLayout(label_n_btn_layout)
        vbox.addWidget(self.column_input)
        vbox.addWidget(self.label_alert)
        vbox.addLayout(file_path_layout)
        vbox.addLayout(start_btn_layout)
        vbox.addStretch(1)

        self.setLayout(vbox)

        self.column_input.setFocus()

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(resource_path('app.ico')))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
