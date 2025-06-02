import toga
from toga.style import Pack
from toga.style.pack import COLUMN, ROW
import asyncio
import time
import pickle
import datetime
import os

from .func import ConnectToGoogleAppsScript, QuestionTitle,PrintData

class DetailWindow:
    def __init__(self, app, selected_data):
        self.window = toga.Window(title='詳細情報')
        self.app = app
        
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
        
        # データ表示用のラベル
        data_new = {}
        for key, item in selected_data.items():
            if key in QuestionTitle.question_dict:
                data_new[QuestionTitle.question_dict[key]] = item
        
        # データを整形して表示
        formatted_data = []
        for key, value in data_new.items():
            formatted_data.append(f"{key}: {value}")
        
        data_text = "\n".join(formatted_data)
        self.data_label = toga.Label(
            data_text,
            style=Pack(
                font_size=14,
                margin=10,
                padding=15,
                background_color='#ffffff'
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
        
        main_box.add(title_label)
        main_box.add(self.data_label)
        main_box.add(print_button)
        
        self.window.content = main_box
        self.window.size = (500, 400)
        self.window.show()

    def print_data(self, widget):
        print_data = PrintData(self.data_label.text)
        print_data.print_data()
        
class Main(toga.App):
    def __init__(self):
        super().__init__()
        self.current_detail_window = None  # 現在開いている詳細画面
        # アプリケーションのルートディレクトリを取得
        self.app_root = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))

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
        
        # テーブル
        self.table = toga.Table(
            headings=['氏名', '生年月日', '確認'],
            data=[],
            style=Pack(
                margin=10,
                padding=5,  # パディングを小さく
                font_size=12,  # フォントサイズを小さく
                background_color='#ffffff',
                height=400,
                flex=1
            )
        )
        
        # 表示ボタン
        show_button = toga.Button(
            '詳細を表示',
            on_press=self.show_details,
            style=Pack(
                margin=10,
                padding=10,
                background_color='#4a90e2',
                color='#ffffff',
                font_size=14,
                font_weight='bold'
            )
        )
        
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
        
        main_box.add(title_label)
        main_box.add(self.date_input)
        main_box.add(self.table)
        main_box.add(show_button)
        main_box.add(self.status_label)

        self.main_window = toga.MainWindow(title=self.formal_name)
        self.main_window.content = main_box
        self.main_window.size = (800, 800)
        self.main_window.show()
        
        # 初期データの更新
        self._update()
        
        # データ更新タスクを開始
        asyncio.create_task(self.update_data())
    
    def show_details(self, widget):
        selected_row = self.table.selection
        if selected_row is not None:
            # 選択された行の氏名と生年月日を取得
            selected_name = selected_row.氏名
            selected_birth = selected_row.生年月日
            
            # データから一致する項目を探す
            selected_data = next(
                (item for item in self.data if item['Q1'] == selected_name
                and item['Q2'] == selected_birth),None)
            
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
                
            # テーブルのデータを更新（完了状態を含める）
            self.table.data = [
                (item['Q1'], item['Q2'], '済' if item.get('completed', False) else '')
                for item in self.data
            ]
            
        except:
            print("データがありません")
            self.data = []
            self.table.data = []

    async def update_data(self):
        while True:
            try:
                self.connection.fetch_questionnaire_data(start_date=self.date_input.value, end_date=self.date_input.value)
                
                # データの保存が完了するのを待つ
                await asyncio.sleep(1)  # ファイルの書き込みが完了するのを待つ
                
                # データを更新
                self._update()
                
                # ステータスを更新
                self.status_label.text = f'最終更新: {time.strftime("%H:%M:%S")}'
                
            except Exception as e:
                self.status_label.text = f'エラーが発生しました: {str(e)}'
            
            # 30秒待機
            await asyncio.sleep(30)

def main():
    return Main()

