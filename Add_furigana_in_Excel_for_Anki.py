import re, jaconv, sys, os, csv, builtins, pandas as pd, platform, subprocess

from fugashi import Tagger
from PyQt6.QtGui import *
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *

from openpyxl import *

# 후리가나 붙이기 ------------------------------------------------------------------
def process_japanese_text(text, exclude_text='', kana_mode='hiragana'):
    CJK_Unified_Ideographs = r'[\u4E00-\u9FFF]+'

    def split_into_blocks(word: str):
        """
        주어진 문자열(word)을 '연속된 한자' 블록(K)과
        '연속된 그 외 문자(주로 히라가나 등)' 블록(H)으로 나누어 리스트로 반환.
        
        예: 
            "問題視する" -> [("K", "問題"), ("K", "視"), ("H", "する")]
            "ご飯" -> [("H", "ご"), ("K", "飯")]
            "忘れ去る" -> [("K", "忘"), ("H", "れ"), ("K", "去"), ("H", "る")]
        """
        blocks = []
        if not word:
            return blocks
        
        # 현재 블록의 종류(K or H), 내용
        current_type = None
        current_buf = []
        
        def flush_buffer():
            """누적된 버퍼를 blocks 리스트에 추가하고 비움"""
            nonlocal current_type, current_buf, blocks
            if current_buf:
                blocks.append((current_type, ''.join(current_buf)))
                current_buf = []
        
        for ch in word:
            if re.compile(CJK_Unified_Ideographs).match(ch):
                # 현재 문자가 한자
                if current_type == 'K':  # 이전도 한자 블록
                    current_buf.append(ch)
                else:
                    # 블록 타입이 바뀌므로 flush 후 새 블록 시작
                    flush_buffer()
                    current_type = 'K'
                    current_buf.append(ch)
            else:
                # 현재 문자는 한자가 아님(히라가나, 가타카나, 알파벳, 기타 등)
                if current_type == 'H':
                    current_buf.append(ch)
                else:
                    flush_buffer()
                    current_type = 'H'
                    current_buf.append(ch)
        
        # 마지막 누적 블록 flush
        flush_buffer()
        return blocks


    def align_word_with_furigana(word: str, reading: str) -> str:
        """
        word(실제 표기)와 reading(전체 후리가나)을 받아
        다음 예시처럼 한자 블록마다 후리가나를 할당하여 변환:
        
        1) "ご飯" + "ごはん"            -> "ご 飯[はん]"
        2) "忘れ去る" + "わすれさる"    -> "忘[わす]れ 去[さ]る"
        3) "問題視する" + "もんだいしする" -> "問題[もんだい] 視[し]する"
        
        구현 아이디어(단순/예시용):
        - word를 '연속된 한자(K) / 그 외(H)' 블록 리스트로 분할
        - reading에서 각 블록에 대응하는 후리가나를 조금씩 소진하면서 할당
            * 한자(K) 블록은 그 다음 블록(특히 H 블록)이 reading 상에 등장하기 직전까지를 통째로 할당
            * H 블록은 가능하면 reading에서도 동일하게 소진(예: 'ご' ↔ 'ご')
        - 블록 사이에서 K→H, H→K 등으로 전환될 때 적절히 공백 삽입
        """
        # 가타카나 모드일 경우 히라가나로 변환
        if kana_mode == 'katakana':
            reading = jaconv.kata2hira(reading)

        # 1) 단어를 블록 리스트로 분할
        blocks = split_into_blocks(word)
        
        # 결과 문자열을 쌓을 리스트
        result = []
        # reading 소비 인덱스
        r_idx = 0
        
        # 다음 블록의 문자열이 reading 내에 있는지 찾는 헬퍼 함수
        def find_next_block_in_reading(next_block_str):
            """reading[r_idx:]에서 next_block_str이 등장하는 첫 위치를 찾는다.
            없으면 -1 반환"""
            if not next_block_str:
                return -1
            return reading.find(next_block_str, r_idx)
        
        prev_type = None
        
        for i, (btype, btext) in enumerate(blocks):
            if btype == 'H':
                # H(히라가나/기타) 블록
                # 가능하면 reading에서도 btext가 일치하면 소비
                length = len(btext)
                # reading[r_idx : r_idx+length]와 btext가 같으면 그대로 소비
                if reading[r_idx:r_idx+length] == btext:
                    # 그대로 사용
                    result.append(btext)
                    r_idx += length
                else:
                    # 일치하지 않으면 그냥 원문 출력 (후리가나 소비는 없음)
                    result.append(btext)
                
                prev_type = 'H'
            
            else:
                # K(한자) 블록
                # 다음 블록이 있다면 그 블록의 텍스트가 reading 상 어느 위치에 나오는지를 확인
                if (i + 1) < len(blocks):
                    next_btype, next_btext = blocks[i+1]
                else:
                    next_btype, next_btext = None, ''
                
                # next_btext가 혹시 reading에서 r_idx 이후에 등장한다면
                # 그 위치를 찾아서 그 직전까지를 이 한자 블록의 후리가나로 할당
                pos_next = -1
                if next_btype == 'H' and next_btext:
                    # 다음 블록이 H라면, reading 상에 exact match로 등장할 수 있으니 찾아봄
                    pos_next = find_next_block_in_reading(next_btext)
                
                if pos_next >= 0:
                    # next_btext가 r_idx 이후에 있다면, 그 직전까지를 한자 블록 후리가나로 할당
                    allocated = reading[r_idx:pos_next]
                    r_idx = pos_next
                else:
                    # 없다면 남은 reading 전부 할당
                    allocated = reading[r_idx:]
                    r_idx = len(reading)
                
                # 앞 블록이 H였으면 한자 블록 앞에 공백 삽입 (질문 예시 규칙)
                if prev_type == 'H' and len(result) > 0 and result[-1] != ' ':
                    result.append(' ')
                
                # "한자블록[후리가나]" 형태로 변환
                if allocated:
                    if kana_mode == 'katakana':
                        result.append(f"{btext}[{jaconv.hira2kata(allocated)}]")
                    else:
                        result.append(f"{btext}[{allocated}]")
                else:
                    # 후리가나가 아예 없으면 그냥 한자 블록만 출력
                    result.append(btext)
                
                prev_type = 'K'
        
        return ''.join(result)


    # 이 함수는 "문장 내 여러 '단어[후리가나]' 패턴"을 찾아 변환해 주는 예시
    # 정규식: ([^\s\[\]]+) => 공백/대괄호 제외 1글자 이상
    #        \[([ぁ-んァ-ン]+)\] => 대괄호 안 히라가나 또는 가타카나 1글자 이상
    WORD_READING_PATTERN = re.compile(r'([^\s\[\]]+)\[([ぁ-んァ-ン]+)\]')

    def convert_text(text: str) -> str:
        """문장 전체에서 '단어[후리가나]' 형태를 찾아 변환"""
        def repl_func(m: re.Match) -> str:
            word = m.group(1)    # 대괄호 앞 실제 단어
            reading = m.group(2) # 대괄호 안 히라가나

            return align_word_with_furigana(word, reading)
        
        return WORD_READING_PATTERN.sub(repl_func, text)

    def add_furigana_with_fugashi(_text, _exclude_text='', _kana_mode='hiragana'):
        excluded_kanji_set = set(ch for ch in exclude_text if re.search(CJK_Unified_Ideographs, ch))
        tagger = Tagger()
        tokens = tagger(text)

        result = []
        for token in tokens:
            surface = token.surface
            # exclude_text에 있는 한자가 하나라도 포함되어 있으면 후리가나 생략
            if any(ch in excluded_kanji_set for ch in surface):
                result.append(surface)
            else:
                # 그 외 일반적인 경우만 후리가나 부착
                if bool(re.compile(CJK_Unified_Ideographs).search(surface)):
                    kana = token.feature.kana
                    if kana:
                        if kana_mode == 'katakana':
                            result.append(f" {surface}[{kana}]")
                        else:  # 기본: 히라가나
                            result.append(f" {surface}[{jaconv.kata2hira(kana)}]")
                    else:
                        result.append(surface)
                else:
                    # 한자 이외(히라가나, 가타카나, 알파벳 등)는 그대로 이어붙임
                    result.append(surface)

        msg = "".join(result)
        if msg.startswith(' '):
            msg = msg[1:]
        return msg

    return convert_text(add_furigana_with_fugashi(text, exclude_text, kana_mode))
'''
japanese_text = "詐欺に遭い憤った被害者達が会社を相手に抗議活動を行った。あの人が言うと、褒め言葉も嫌味に聞こえる。"
print(process_japanese_text(japanese_text))
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

        if Existed_column_list and self.parent.overWrite_mode == True:
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
                        item = str(self.sheet[number_to_column(column_to_number(x)+1)+str(y[0])].value)

                        if item != None and item != 'None' and item.strip() != '' and self.parent.overWrite_mode == False:
                            self.sheet[number_to_column(column_to_number(x)+1)+str(y[0])] = item
                        else:
                            self.sheet[number_to_column(column_to_number(x)+1)+str(y[0])] = process_japanese_text(text=y[1], kana_mode=self.parent.kana_mode)
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
                                    item = str(self.sheet[number_to_column(column_to_number(x[1])+1)+str(y[0])].value)

                                    if item != None and item != 'None' and item.strip() != '' and self.parent.overWrite_mode == False:
                                        self.sheet[number_to_column(column_to_number(x[1])+1)+str(y[0])] = item
                                    else:
                                        self.sheet[number_to_column(column_to_number(x[1])+1)+str(y[0])] = process_japanese_text(text=y[1],exclude_text=z[1], kana_mode=self.parent.kana_mode)
                        else:
                            item = str(self.sheet[number_to_column(column_to_number(x[1])+1)+str(y[0])].value)
                            
                            if item != None and item != 'None' and item.strip() != '' and self.parent.overWrite_mode == False:
                                self.sheet[number_to_column(column_to_number(x[1])+1)+str(y[0])] = item
                            else:
                                self.sheet[number_to_column(column_to_number(x[1])+1)+str(y[0])] = process_japanese_text(text=y[1], kana_mode=self.parent.kana_mode)
                    
                    # 단어 처리
                    for y in words:
                        item = str(self.sheet[number_to_column(column_to_number(x[0])+1)+str(y[0])].value)

                        if item != None and item != 'None' and item.strip() != '' and self.parent.overWrite_mode == False:
                            self.sheet[number_to_column(column_to_number(x[0])+1)+str(y[0])] = item
                        else:
                            self.sheet[number_to_column(column_to_number(x[0])+1)+str(y[0])] = process_japanese_text(text=y[1], kana_mode=self.parent.kana_mode)
            
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
                        item = str(df.iloc[y[0]-1, column_to_number(x)])
                        
                        if item != None and item != 'nan' and item.strip() != '' and self.parent.overWrite_mode == False:
                            df.iloc[y[0]-1, column_to_number(x)] = item
                        else:
                            df.iloc[y[0]-1, column_to_number(x)] = process_japanese_text(text=y[1],kana_mode=self.parent.kana_mode)
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
                                    item = str(df.iloc[y[0]-1, column_to_number(x[1])])

                                    if item != None and item != 'nan' and item.strip() != '' and self.parent.overWrite_mode == False:
                                        df.iloc[y[0]-1, column_to_number(x[1])] = item
                                    else:
                                        df.iloc[y[0]-1, column_to_number(x[1])] = process_japanese_text(text=y[1], exclude_text=z[1],kana_mode=self.parent.kana_mode)
                        else:
                            item = str(df.iloc[y[0]-1, column_to_number(x[1])])

                            if item != None and item != 'nan' and item.strip() != '' and self.parent.overWrite_mode == False:
                                df.iloc[y[0]-1, column_to_number(x[1])] = item
                            else:
                                df.iloc[y[0]-1, column_to_number(x[1])] = process_japanese_text(text=y[1],kana_mode=self.parent.kana_mode)
                    
                    #단어 처리
                    for y in words:
                        item = str(df.iloc[y[0]-1, column_to_number(x[0])])

                        if item != None and item != 'nan' and item.strip() != '' and self.parent.overWrite_mode == False:
                            df.iloc[y[0]-1, column_to_number(x[0])] = item
                        else:
                            df.iloc[y[0]-1, column_to_number(x[0])] = process_japanese_text(text=y[1],kana_mode=self.parent.kana_mode)
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
    overWrite_mode = False

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
        filepath = QFileDialog.getOpenFileName(self, 'ファイル選択', '','Excel Files (*.xlsx *.xlsm *.csv)')
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

    def get_overWrite_btn_value(self):
        if self.overWrite_btn.isChecked():
            self.overWrite_mode = True
        else:
            self.overWrite_mode = False
            

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
        self.overWrite_btn = QCheckBox('上書きモード', self)
        self.overWrite_btn.clicked.connect(self.get_overWrite_btn_value)

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
        kana_btn_layout.addWidget(QLabel('', self))
        kana_btn_layout.addWidget(self.overWrite_btn)
        kana_btn_layout.addStretch(1)

        label_n_btn_layout = QHBoxLayout()
        label_n_btn_layout.addWidget(label_column_input)
        label_n_btn_layout.addStretch(1)
        label_n_btn_layout.addWidget(QLabel('', self))
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
