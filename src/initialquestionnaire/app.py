import toga
from toga.style import Pack
from toga.style.pack import COLUMN, ROW
import asyncio
import time
import pickle
import datetime
import os
import win32print

from .func import ConnectToGoogleAppsScript, QuestionTitle,PrintData

def make_data_label(data):
    data_new = {}
    for key, item in data.items():
        if key in QuestionTitle.question_dict:
            if key == "Q2":
                age = calc_age(item)
                item = f"{item} ({age})"
            data_new[QuestionTitle.question_dict[key]] = item
    
    # データを整形して表示
    formatted_data = []
    for key, value in data_new.items():
        formatted_data.append(f"{key}: {value}")
    
    data_text = "\n".join(formatted_data)
    return data_text

def calc_age(birthday):
    birthday_datetime = datetime.datetime.strptime(birthday, '%Y-%m-%d')
    today = datetime.date.today()
    age = today.year - birthday_datetime.year
    if (today.month, today.day) < (birthday_datetime.month, birthday_datetime.day):
        age -= 1

    # 3歳以下の場合は「X歳Yか月」
    if age <= 3:
        # 月数の計算
        months = (today.year - birthday_datetime.year) * 12 + today.month - birthday_datetime.month
        if today.day < birthday_datetime.day:
            months -= 1
        years = months // 12
        remain_months = months % 12
        if years == 0:
            return f"{remain_months}か月"
        elif remain_months == 0:
            return f"{years}歳"
        else:
            return f"{years}歳{remain_months}か月"
    else:
        return f"{age}歳"
    
class DetailWindow:
    def __init__(self, app, selected_data):
        self.window = toga.Window(title='詳細情報')
        self.app = app
        self.app_root = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
        self.selected_data = selected_data

        # メインボックス
        main_box = toga.Box(style=Pack(
            direction=COLUMN,
            margin=20,
            background_color='#f5f5f5'
        ))
        
        # タイトル
        title_label = toga.Label(
            '詳細情報',
            style=Pack(
                font_size=20,
                font_weight='bold',
                margin_bottom=20,
                color='#333333'
            )
        )
        
        data_text = make_data_label(self.selected_data)

        self.data_label = toga.MultilineTextInput(
            value=data_text,
            readonly=True,
            style=Pack(
                font_size=12,
                margin=10,
                padding=15,
                background_color='#ffffff',
                height=500,
            )
        )
        
        # 印刷ボタン
        print_button = toga.Button(
            '印刷',
            on_press=self.print_data,
            style=Pack(
                margin=10,
                padding=10,
                background_color='#4a90e2',
                color='#ffffff',
                font_size=14,
                font_weight='bold'
            )
        )

        # 完了状態ボタン
        self.complete_button = toga.Button(
            '未完了' if not self.selected_data['completed'] else '完了済み',
            on_press=self.toggle_completion,
            style=Pack(
                margin=10,
                padding=10,
                background_color='#4a90e2',
                color='#ffffff',
                font_size=14,
                font_weight='bold'
            )
        )
        
        
        main_box.add(title_label)
        main_box.add(self.data_label)
        main_box.add(self.complete_button)
        main_box.add(print_button)
        
        self.window.content = main_box
        self.window.size = (600, 800)
        self.window.show()

    def print_data(self, widget):
        print_data = PrintData(self.data_label.text)
        print_data.print_data(printer_name=self.app.selected_printer)
        self.toggle_completion(widget, only_complete=True)

    def toggle_completion(self, widget, only_complete=False):
        """選択されたデータの完了状態を切り替える"""
        if not self.selected_data:
            return

        # 選択されたデータの日付を取得
        date_str = self.selected_data['timestamp'].split()[0]
        file_path = os.path.join(self.app_root, "data", f"{date_str}.pickle")

        try:
            # 日付のデータを読み込む
            with open(file_path, 'rb') as f:
                records = pickle.load(f)

            # 選択されたデータの完了状態を切り替え
            for record in records:
                if only_complete:
                    if record['timestamp'] == self.selected_data['timestamp']:
                        record['completed'] = True
                        self.selected_data['completed'] = record['completed']
                        break
                else:
                    if record['timestamp'] == self.selected_data['timestamp']:
                        record['completed'] = not record['completed']
                        self.selected_data['completed'] = record['completed']
                        break

            # 更新されたデータを保存
            with open(file_path, 'wb') as f:
                pickle.dump(records, f)

            # 完了状態に応じてボタンのテキストを更新
            self.complete_button.text = "完了済み" if self.selected_data['completed'] else "未完了"
            
            # Mainクラスの_updateメソッドを呼び出してテーブルを更新
            self.app._update()
            
        except Exception as e:
            print(f"完了状態の更新に失敗しました: {str(e)}")

    
class Main(toga.App):
    def __init__(self):
        super().__init__()
        self.current_detail_window = None  # 現在開いている詳細画面
        self.app_root = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
        # プリンター一覧と選択中プリンター
        self.printers = [printer[2] for printer in win32print.EnumPrinters(2)]
        self.selected_printer = win32print.GetDefaultPrinter() if self.printers else None
        self.hide_completed = False
        self.sort_mode = 'timestamp'

    def startup(self):
        self.connection = ConnectToGoogleAppsScript()
        # メインボックス
        main_box = toga.Box(style=Pack(
            direction=COLUMN,
            margin=20,
            background_color='#f5f5f5'
        ))
        
        # タイトル
        title_label = toga.Label(
            'データ選択',
            style=Pack(
                font_size=24,
                font_weight='bold',
                margin_bottom=20,
                color='#333333'
            )
        )
        
        # 日付
        self.date_input = toga.DateInput(
            value=datetime.date.today(),
            on_change=self._update
            )
        
        # 追加: 完了済み非表示チェックボックス
        self.hide_completed_checkbox = toga.Switch(
            '完了済みを非表示',
            value=self.hide_completed,
            on_change=self.on_hide_completed_toggle,
            style=Pack(margin=10, padding=5, font_size=12)
        )
        # 追加: ソート方法選択用ラジオボタン（Selectionで代用）
        self.sort_selection = toga.Selection(
            items=['生年月日でソート', '回答日時でソート'],
            on_change=self.on_sort_selection,
            style=Pack(margin=10, padding=5, font_size=10,width=150)
        )
        self.sort_selection.value = '生年月日でソート'
        
        # テーブル
        self.table = toga.Table(
            headings=['回答時間', '氏名', '生年月日', '確認'],
            data=[],
            style=Pack(
                width=500,
                margin=5,
                padding=5,  # パディングを小さく
                font_size=10,  # フォントサイズを小さく
                background_color='#ffffff',
                height=500,
                flex=1
            )
        )
        
        # 表示ボタンと印刷ボタンを横並びにするためのBox
        button_box = toga.Box(style=Pack(
            direction=ROW,
            margin=10,
            padding=5
        ))

        # 表示ボタン
        show_button = toga.Button(
            '詳細を表示',
            on_press=self.show_details,
            style=Pack(
                margin=5,
                padding=10,
                background_color='#4a90e2',
                color='#ffffff',
                font_size=14,
                font_weight='bold',
                flex=1
            )
        )

        # 印刷ボタン
        print_button = toga.Button(
            "印刷",
            on_press=self.print_data,
            style=Pack(
                margin=5,
                padding=10,
                background_color='#4a90e2',
                color='#ffffff',
                font_size=14,
                font_weight='bold',
                flex=1
            )
        )

        # ボタンをBoxに追加
        button_box.add(show_button)
        button_box.add(print_button)
        
        # ステータス表示用ラベル
        self.status_label = toga.Label(
            'データ更新待機中...',
            style=Pack(
                margin=10,
                padding=5,
                font_size=12,
                color='#666666',
                background_color='#ffffff'
            )
        )
        
        # プリンター選択ドロップダウン
        self.printer_selection = toga.Selection(
            items=self.printers,
            on_change=self.on_printer_select,
            style=Pack(margin=10, padding=5, font_size=12, background_color='#ffffff')
        )
        if self.selected_printer in self.printers:
            self.printer_selection.value = self.selected_printer
        elif self.printers:
            self.printer_selection.value = self.printers[0]
            self.selected_printer = self.printers[0]
        
        # 追加: チェックボックスとソート選択を横並びに
        filter_sort_box = toga.Box(style=Pack(direction=ROW))
        filter_sort_box.add(self.hide_completed_checkbox)
        filter_sort_box.add(self.sort_selection)
        
        main_box.add(title_label)
        main_box.add(self.date_input)
        main_box.add(filter_sort_box)
        main_box.add(self.table)
        main_box.add(button_box)
        main_box.add(self.status_label)
        main_box.add(self.printer_selection)

        self.main_window = toga.MainWindow(title=self.formal_name)
        self.main_window.content = main_box
        self.main_window.size = (450, 800)
        self.main_window.show()
        
        # 初期データの更新
        self._update()
        
        # データ更新タスクを開始
        asyncio.create_task(self.update_pickle_data())

    def print_data(self, widget):
        selected_row = self.table.selection
        if selected_row is not None:
            # 選択された行の氏名と生年月日を取得
            selected_name = selected_row.氏名
            selected_birth = selected_row.生年月日
            # データから一致する項目を探す
            selected_data = next(
                (item for item in self.data if item['Q1'] == selected_name
                and item['Q2'] == selected_birth),None)
            data_text = make_data_label(selected_data)

            print_data = PrintData(data_text)
            print_data.print_data(printer_name=self.selected_printer)
            self.toggle_completed_only(widget)

    def get_selected_data(self,widget):
        selected_row = self.table.selection
        if selected_row is not None:
            # 選択された行の氏名と生年月日を取得
            selected_name = selected_row.氏名
            selected_birth = selected_row.生年月日
            
            # データから一致する項目を探す
            selected_data = next(
                (item for item in self.data if item['Q1'] == selected_name
                and item['Q2'] == selected_birth),None)
            return selected_data
        else:
            return None

    def show_details(self, widget):
        selected_data = self.get_selected_data(widget)
        if selected_data:
            # 既存の詳細画面を閉じる
            if self.current_detail_window is not None:
                try:
                    self.current_detail_window.window.close()
                except:
                    pass  # ウィンドウが既に閉じられている場合は無視
            
            # 新しい詳細画面を開く
            self.current_detail_window = DetailWindow(self, selected_data)

    def _update(self, widget=None):
        try:
            # アプリケーションのルートディレクトリからの相対パスで指定
            data_path = os.path.join(self.app_root, "data", f"{self.date_input.value.strftime('%Y-%m-%d')}.pickle")
            with open(data_path, "rb") as f:
                self.data = pickle.load(f)
            # 追加: 完了済み非表示フィルタ
            filtered_data = self.data
            if self.hide_completed:
                filtered_data = [item for item in filtered_data if not item.get('completed', False)]
            # 追加: ソート（生年月日 or timestamp）
            if self.sort_mode == 'birthday':
                filtered_data = sorted(
                    filtered_data,
                    key=lambda x: datetime.datetime.strptime(x.get('Q2', '1900-01-01').split('（')[0], '%Y-%m-%d')
                )
            else:
                filtered_data = sorted(
                    filtered_data,
                    key=lambda x: x.get('timestamp', '')
                )
            # テーブルのデータを更新（完了状態を含める）
            self.table.data = []
            self.table.data = [
                (item['timestamp'].split()[1], item['Q1'], item['Q2'], '✓' if item.get('completed', False) else '')
                for item in filtered_data
            ]
        except:
            print("データがありません")
            self.data = []
            self.table.data = []

    def toggle_completed_only(self,widget):
        # 選択されたデータの日付を取得
        selected_data = self.get_selected_data(widget)
        date_str = selected_data['timestamp'].split()[0]
        file_path = os.path.join(self.app_root, "data", f"{date_str}.pickle")

        try:
            # 日付のデータを読み込む
            with open(file_path, 'rb') as f:
                records = pickle.load(f)

            # 選択されたデータの完了状態を切り替え
            for record in records:
                if record['timestamp'] == selected_data['timestamp']:
                    record['completed'] = True
                    selected_data['completed'] = record['completed']
                    break

            # 更新されたデータを保存
            with open(file_path, 'wb') as f:
                pickle.dump(records, f)

            # Mainクラスの_updateメソッドを呼び出してテーブルを更新
            self._update()
            
        except Exception as e:
            print(f"完了状態の更新に失敗しました: {str(e)}")

    async def update_pickle_data(self):
        loop = asyncio.get_event_loop()
        while True:
            try:
                # fetch_questionnaire_dataを別スレッドで実行
                await loop.run_in_executor(
                    None,
                    self.connection.fetch_questionnaire_data,
                    self.date_input.value,
                    self.date_input.value
                )
                # データの保存が完了するのを待つ
                await asyncio.sleep(1)
                # データを更新
                self._update()
                # ステータスを更新
                self.status_label.text = f'最終更新: {time.strftime("%H:%M:%S")}'
            except Exception as e:
                self.status_label.text = f'エラーが発生しました: {str(e)}'
            # 30秒待機
            await asyncio.sleep(30)

    def on_printer_select(self, widget):
        self.selected_printer = widget.value

    def on_hide_completed_toggle(self, widget):
        self.hide_completed = widget.value
        self._update()

    def on_sort_selection(self, widget):
        if widget.value == '生年月日でソート':
            self.sort_mode = 'birthday'
        else:
            self.sort_mode = 'timestamp'
        self._update()

def main():
    return Main()

