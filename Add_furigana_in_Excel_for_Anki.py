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

#-----------------------------------------------------------------------------------
def column_to_number(column_name):
    # 엑셀 열 이름을 숫자로 변환
    column_number = 0
    for i, char in enumerate(reversed(column_name.upper())):
        column_number += (ord(char) - ord('A') + 1) * (26 ** i)
    return column_number

def number_to_column(column_number):
    # 숫자를 엑셀 열 이름으로 변환
    column_name = []
    while column_number > 0:
        column_number -= 1  # 1을 빼서 0부터 시작하도록 조정
        column_name.append(chr(column_number % 26 + ord('A')))
        column_number //= 26
    return ''.join(reversed(column_name))

def check_file_is_open(file_path):
    # 파일이 열려있는지 확인하는 함수
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

def parse_mixed_input(input_str):
    # 튜플 패턴 찾기 (괄호 안의 내용 추출)
    tuple_pattern = re.compile(r'\([^()]*\)')

    def safe_eval(match):
        # 괄호 제거
        if match.group(0)[0]==',':
            content = match.group(0)[2:-1]
        else:
            content = match.group(0)[1:-1]

        items = [item.strip() for item in content.split(',') if item.strip()]
        
        # 요소가 하나만 있을 경우 튜플로 반환되도록 처리
        if len(items) == 1:
            return (items[0],)
        return tuple(items)

    # 튜플을 파싱하여 리스트로 저장
    tuples = []
    # finditer를 사용하여 input_str에서 패턴에 일치하는 모든 match 객체 검색
    for match in tuple_pattern.finditer(input_str):
        # 매칭된 결과(match)를 safe_eval 함수로 처리
        evaluated_value = safe_eval(match)
        
        # 결과를 리스트에 추가
        tuples.append(evaluated_value)
    
    # 문자열에서 튜플 패턴 제거 후 남은 요소 처리
    remaining_str = tuple_pattern.sub('tuple', input_str)
    
    remaining_elements = []  # 빈 리스트 초기화
    # 남은 문자열을 리스트로 변환, 빈 문자열 처리
    if remaining_str:
        split_items = remaining_str.split(',')  # 쉼표로 문자열 나누기
        for item in split_items:  # stripped_items 리스트의 각 요소를 순회
            if item is None or item.strip() == '':
                remaining_elements.append('')
            elif item != 'tuple':
                remaining_elements.append(item)
                
    else:
        remaining_elements = []

    return tuples, remaining_elements if remaining_elements else []

# UI 만들기 ------------------------------------------------------------------------
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
        self.input_columns_array = []
        self.tuples, self.lists = parse_mixed_input(self.columns)
        
        # 리스트 처리
        for item in self.lists:
            self.input_columns_array.append(item)
        
        # 튜플 처리
        for item in self.tuples:
            for element in item:
                self.input_columns_array.append(element)
        
        self.output_columns_array = [number_to_column(column_to_number(item) + 1) for item in self.input_columns_array]
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
        # XLSX, XLSM 파일 처리
        if self.filepath.endswith('.xlsx') or self.filepath.endswith('.xlsm'):
            if self.lists:
                for x in self.lists:
                    for y in self.get_multiple_columns_with_rows(x)[x]:
                        self.sheet[number_to_column(column_to_number(x)+1)+str(y[0])] = add_furigana_with_fugashi(text=y[1], kana_mode=self.parent.kana_mode)
                self.workbook.save(self.filepath)

            if self.tuples:
                for x in self.tuples:
                    words = [item for item in self.get_multiple_columns_with_rows(x[0])[x[0]]]
                    words_number = [item[0] for item in self.get_multiple_columns_with_rows(x[0])[x[0]]]
                    sentences = [item for item in self.get_multiple_columns_with_rows(x[1])[x[1]]]

                    # 문장 처리                    
                    for y in sentences:
                        if y[0] in words_number:
                            for z in words:
                                if z[0] == y[0]:
                                    self.sheet[number_to_column(column_to_number(x[1])+1)+str(y[0])] = add_furigana_with_fugashi(text=y[1],exclude_text=z[1], kana_mode=self.parent.kana_mode)
                        else:
                            self.sheet[number_to_column(column_to_number(x[1])+1)+str(y[0])] = add_furigana_with_fugashi(text=y[1], kana_mode=self.parent.kana_mode)
                    
                    # 단어 처리
                    for y in words:
                        self.sheet[number_to_column(column_to_number(x[0])+1)+str(y[0])] = add_furigana_with_fugashi(text=y[1], kana_mode=self.parent.kana_mode)
            
                self.workbook.save(self.filepath)

        # CSV 파일 처리
        elif self.filepath.endswith('.csv'):
            df = pd.read_csv(self.filepath, encoding='utf-8', header=None, dtype=str)

            required_cols = max([column_to_number(x) for x in self.output_columns_array])
            while len(df.columns) <= required_cols:
                df[f'new_col_{len(df.columns)}'] = None

            if self.lists:
                for x in self.lists:
                    for y in self.get_multiple_columns_with_rows(x)[x]:
                        df.iloc[y[0]-1, column_to_number(x)] = add_furigana_with_fugashi(text=y[1],kana_mode=self.parent.kana_mode)
                df.to_csv(self.filepath, encoding='utf-8-sig', index=False, header=None)

            if self.tuples:
                for x in self.tuples:
                    words = [item for item in self.get_multiple_columns_with_rows(x[0])[x[0]]]
                    words_number = [item[0] for item in self.get_multiple_columns_with_rows(x[0])[x[0]]]
                    sentences = [item for item in self.get_multiple_columns_with_rows(x[1])[x[1]]]

                    # 문장 처리
                    for y in sentences:
                        if y[0] in words_number:
                            for z in words:
                                if z[0] == y[0]:
                                    df.iloc[y[0]-1, column_to_number(x[1])] = add_furigana_with_fugashi(text=y[1], exclude_text=z[1],kana_mode=self.parent.kana_mode)
                        else:
                            df.iloc[y[0]-1, column_to_number(x[1])] = add_furigana_with_fugashi(text=y[1],kana_mode=self.parent.kana_mode)
                    
                    #단어 처리
                    for y in words:
                        df.iloc[y[0]-1, column_to_number(x[0])] = add_furigana_with_fugashi(text=y[1],kana_mode=self.parent.kana_mode)
                df.to_csv(self.filepath, encoding='utf-8-sig', index=False, header=None)

        else:
            self.fault_signal.emit(4)
            self.fault_message = 'ファイルの形式が間違っています。'
            
        self.fault_message = '完了しました。'
        self.fault_signal.emit(3)

class AutoLineEdit(QLineEdit):
    def __init__(self):
        super().__init__()
        '''
        self.min_width = 200
        self.setFixedWidth(self.min_width)
        self.textChanged.connect(self.adjust_size)  # 텍스트 변경 시 크기 조정
        '''

    def adjust_size(self):
        # 텍스트의 폭 계산
        font_metrics = QFontMetrics(self.font())
        text_width = font_metrics.horizontalAdvance(self.text()) + 20  # 추가 여백
        self.setFixedWidth(max(self.min_width, text_width))  # 최소 폭 100 설정

    def keyPressEvent(self, event):
        current_text = self.text()
        cursor_pos = self.cursorPosition()
        key_input = event.text()

        # 괄호 개수 확인 함수
        def count_brackets(text):
            open_count = text.count("(")
            close_count = text.count(")")
            return open_count, close_count

        open_brackets, close_brackets = count_brackets(current_text)

        # "(" 입력 처리 (닫히지 않은 괄호가 있다면 입력 불가)
        if key_input == "(":
            # 닫히지 않은 괄호가 있는 경우 "(" 입력 방지
            if open_brackets > close_brackets:
                return

            # 커서가 닫힌 괄호 앞에 위치하는지 확인
            if cursor_pos < len(current_text) and current_text[cursor_pos] == ")":
                return  # 닫힌 괄호 앞에서는 추가 입력 방지
            
            # 커서 직전 문자가 ,가 아닌 경우 ",(" 자동 삽입
            if cursor_pos > 0 and current_text[cursor_pos - 1] == "(":
                pass
            elif cursor_pos > 0 and current_text[cursor_pos - 1] != ",":
                new_text = current_text[:cursor_pos] + ",(" + current_text[cursor_pos:]
                self.setText(new_text)
                # 커서는 새로 삽입된 2글자 뒤로 이동
                self.setCursorPosition(cursor_pos + 2)
                return

        # ")" 입력 처리 (모든 괄호가 닫혔다면 입력 불가)
        if key_input == ")":
            # 모든 괄호가 이미 닫혀 있는 경우 ")" 입력 방지
            if open_brackets == close_brackets:
                return
            
            # 커서가 이미 닫힌 괄호 바로 앞에 위치한 경우는 통과
            if cursor_pos < len(current_text) and current_text[cursor_pos] == ")":
                return
            
            '''
            # 커서 직후 문자가 ,가 아닌 경우 ")," 자동 삽입
            if cursor_pos == len(current_text) or current_text[cursor_pos] != ",":
                new_text = current_text[:cursor_pos] + ")," + current_text[cursor_pos:]
                self.setText(new_text)
                # 커서는 새로 삽입된 2글자 뒤로 이동
                self.setCursorPosition(cursor_pos + 2)
                return
            '''
            
        # 알파벳 입력 처리
        if key_input.isalpha():
            # 직전 글자가 ")" 일 때 "," 입력
            if cursor_pos > 0 and current_text[cursor_pos - 1] == ")":
                new_text = current_text[:cursor_pos] + "," + key_input + current_text[cursor_pos:]
                self.setText(new_text)
                self.setCursorPosition(cursor_pos + 2)
                return

            
            # 네번째 알파벳 입력방지 !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! 수정 필요
            '''
            if cursor_pos > 2 and all(ch.isalpha() for ch in current_text[cursor_pos - 3:cursor_pos]):
                new_text = current_text[:cursor_pos-1] + key_input + "," + current_text[cursor_pos-1:]
                self.setText(new_text)
                self.setCursorPosition(cursor_pos + 2)
                return
            '''
            
            # 만약 커서 오른쪽에 '('가 있으면 자동으로 ',' 추가 등
            if cursor_pos < len(current_text) and current_text[cursor_pos] == "(":
                new_text = (
                    current_text[:cursor_pos]
                    + key_input
                    + ","
                    + current_text[cursor_pos:]
                )
                self.setText(new_text)
                # 여기서는 알파벳 이후 +1 (콤마는 커서 뒤에 붙어서?)
                self.setCursorPosition(cursor_pos + 1)
                return
        
        '''
        # Backspace 입력 처리
        if event.key() == Qt.Key.Key_Backspace:
            # 커서의 왼쪽 두번째 문자가 "," 이면 직전 문자와 "," 지우기
            if cursor_pos > 2 and current_text[cursor_pos -2] == "," and current_text[cursor_pos - 1] != ',':
                new_text = current_text[:cursor_pos-2]+current_text[cursor_pos:]
                self.setText(new_text)
                self.setCursorPosition(cursor_pos - 2)
                return
        '''

        # 나머지 키 입력은 기본 동작 그대로 진행
        super().keyPressEvent(event)

class MainWindow(QWidget):
    window_width = 500
    window_height = 400
    kana_mode = 'hiragana'

    def __init__(self):
        super(MainWindow, self).__init__()
        self.list_columns = []
        self.tuple_colums = []
        self.response = None
        
        self.alert_start = 0
        self.alert_columns_is_null = 0
        self.msg_columns_is_null = '単語がある列を入力してください。'
        self.alert_columns_contains_null = 0
        self.msg_columns_contains_null = '入力した列に空白があります。'
        self.alert_column_out_of_range = 0
        self.msg_column_out_of_range = '列の範囲が外れました。範囲はAからXFD列までです。'
        self.alert_column_overlap = 0
        self.msg_column_overlap = '選択した列と出力する列が重なります。'
        self.alert_bracket_is_not_close = 0
        self.msg_bracket_is_not_close = '括弧が閉じていません。'
        self.alert_bracket_overflow = 0
        self.msg_bracket_overflow = '括弧の中に二つ以内の列を入力してください。'
        
        self.initUI()

    def check_contains_strange(self, tuples=[], lists=[]):
        all_columns = []

        # 리스트 처리
        for item in lists:
            if item is None or item.strip() == '':
                return 1
            else:
                all_columns.append(item)
        
        # 튜플 처리
        for item in tuples:
            for element in item:
                if element is None or element.strip() == '':
                    return 1
                else:
                    all_columns.append(element)
            try:
                if item[2]:
                    return 5
            except:
                pass
            
        
        # 공통 처리
        for x in range(len(all_columns)):
            # 괄호가 닫히지 않았는지 확인
            if all_columns[x][0] == '(' or all_columns[x][len(all_columns[x])-1] == ')':
                return 4

            # 열 입력 범위 제한
            elif not column_to_number('A') <= column_to_number(all_columns[x]) <= column_to_number('XFD'):
                return 2
            
            # 열 겹침 방지
            elif len(all_columns) != 1:
                for y in range(len(all_columns)):
                    if abs(column_to_number(all_columns[x])-column_to_number(all_columns[y])) == 1:
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
            #self.columns_array = [item.strip() for item in text.split(',')]
            self.tuple_colums, self.list_columns = parse_mixed_input(text)

            alert = self.check_contains_strange(self.tuple_colums, self.list_columns)
            if alert == 1:
                self.label_alert.setText(self.msg_columns_contains_null)
                self.alert_columns_contains_null = 1
            elif alert == 2:
                self.label_alert.setText(self.msg_column_out_of_range)
                self.alert_column_out_of_range = 1
            elif alert == 3:
                self.label_alert.setText(self.msg_column_overlap)
                self.alert_column_overlap = 1
            elif alert == 4:
                self.label_alert.setText(self.msg_bracket_is_not_close)
                self.alert_bracket_is_not_close = 1
            elif alert == 5:
                self.label_alert.setText(self.msg_bracket_overflow)
                self.alert_bracket_overflow = 1
            else:
                self.label_alert.setText('')
                self.alert_columns_contains_null = 0
                self.alert_column_out_of_range = 0
                self.alert_column_overlap = 0
                self.alert_bracket_is_not_close = 0
                self.alert_bracket_overflow = 0
        
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
        elif signal == 3:
            msg_box.setWindowTitle('報告')
        
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
        '''
        self.setMinimumWidth(self.window_width)
        self.setMinimumHeight(self.window_height)
        self.resize(self.window_width, self.window_height)
        '''
       
        '''
        # 1.0버전 설명
        label_column_input = QLabel(
            '単語がある列のアルファベットを入力してください。\n'
            'フリガナは単語がある列の右の列に出力されます。\n'
            'コンマ「,」を書いて多重入力もできます。例）A,C,E'
            ,self)
        '''
        # 1.1버전 설명
        label_column_input = QLabel(
            '<b>使い方</b><br>'
            '<br>'
            '文字がある列のアルファベットを入力してください。<br>'
            'フリガナが付けられた文字は文字がある列の右の列に出力されます。<br>'
            '<br>'
            'フリガナを付けたい文章の中で特定の単語にフリガナを付けたくない場合には括弧の中に（単語列，文章列）の様に入力してください。<br>'
            '<br>'
            'コンマ「,」を書いて多重入力もできます。<br>'
            '例）\"A,C,(E,G),(I,K),M\"<br>'
            '<br>'
            '上の様に入力した場合Ｂ、Ｄ、Ｆ、Ｊ、Ｎ列にフリガナが付けられた文字が、Ｈ列にはＥ列に含まれた単語以外のＧ列の文章にフリガナが付けられて出力します。(I,K)も同じく作動します。<br>'
            ,self)
        label_column_input.setWordWrap(True)

        self.hiragana_btn = QRadioButton('平仮名出力', self)
        self.hiragana_btn.setChecked(True)
        self.hiragana_btn.clicked.connect(self.get_selected_value)
        self.katakana_btn = QRadioButton('片仮名出力', self)
        self.katakana_btn.clicked.connect(self.get_selected_value)

        self.column_input = AutoLineEdit()
        self.label_alert = QLabel('', self)
        self.label_alert.setStyleSheet('color: red;')
        
        #self.column_input.setText('b,d')
        regex = QRegularExpression('^[a-zA-Z,()]*$')  # 영문자와 쉼표, 괄호만 허용
        validator = QRegularExpressionValidator(regex, self.column_input)
        self.column_input.setValidator(validator)
        
        self.check_columns_text(self.column_input.text())
        self.column_input.textChanged.connect(self.check_columns_text)
        
        self.qle_file_path = QLineEdit(self)
        btn_file_path_select = QPushButton('...', self)
        btn_file_path_select.clicked.connect(self.SelctFilePath)

        label_start = QLabel('始める前にエクセルを閉じてください。',self)
        btn_start = QPushButton('始め', self)
        btn_start.clicked.connect(self.Start)

        kana_btn_layout = QVBoxLayout()
        kana_btn_layout.addWidget(self.hiragana_btn)
        kana_btn_layout.addWidget(self.katakana_btn)
        kana_btn_layout.addStretch(1)

        label_n_btn_layout = QHBoxLayout()
        label_n_btn_layout.addWidget(label_column_input)
        label_n_btn_layout.addStretch(1)
        label_n_btn_layout.addLayout(kana_btn_layout)

        column_input_layout = QHBoxLayout()
        column_input_layout.addWidget(QLabel('列を入力', self))
        column_input_layout.addWidget(self.column_input)

        file_path_layout = QHBoxLayout()
        file_path_layout.addWidget(QLabel('ファイル位置', self))
        file_path_layout.addWidget(self.qle_file_path)
        file_path_layout.addWidget(btn_file_path_select)

        start_btn_layout = QHBoxLayout()
        start_btn_layout.addStretch(1)
        start_btn_layout.addWidget(label_start)
        start_btn_layout.addWidget(btn_start)

        vbox = QVBoxLayout()
        vbox.addStretch(1)
        vbox.addLayout(label_n_btn_layout)
        vbox.addLayout(column_input_layout)
        vbox.addWidget(self.label_alert)
        vbox.addLayout(file_path_layout)
        vbox.addLayout(start_btn_layout)
        vbox.addStretch(1)

        self.setLayout(vbox)

        self.column_input.setFocus()

def resource_path(relative_path):
    ''' Get absolute path to resource, works for dev and for PyInstaller '''
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(resource_path('app.ico')))

    QFontDatabase.addApplicationFont(resource_path('UDDigiKyokashoN-R.ttc'))
    app.setFont(QFont('UD Digi Kyokasho N-R', 10))
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
