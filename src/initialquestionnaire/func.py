"""
Google Apps Scriptのエンドポイントからデータを取得するクライアントプログラム
"""
import os
import pickle
import requests
import datetime
import json
import pytz
import re
import json
import pickle
import datetime
import os
import tempfile
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import win32api
import win32print

class QuestionTitle:

    question_dict = {
        "timestamp":"回答日時",
        "Q1":"氏名",
        "Q2":"生年月日",
        "Q3":"住所",
        "Q4":"電話番号",
        "Q5":"緊急連絡先",
        "Q6":"お薬手帳",
        "Q7":"今回処方以外の内服薬等有無",
        "Q8":"手帳に記載無い内服薬等名称",
        "Q9":"既往歴",
        "Q10":"副作用歴有無",
        "Q11":"副作用の薬品名、症状",
        "Q12":"アレルギー歴",
        "Q13":"嗜好品、運転など",
        "Q14":"妊娠有無",
        "Q15":"出産予定日",
        "Q16":"授乳有無",
        "Q17":"授乳している子供の月齢",
    }

class ConnectToGoogleAppsScript:
    def __init__(self):
        self.API_URL = "https://script.google.com/macros/s/AKfycbwTGOlcxFNifdzjOHwicLFceVprKk5nWs8iQAk63U0gwHZJnd8lKxToB3SRdF6pqxEAMQ/exec"
        self.token = "j8K2pF7xQvZmE9hN3tYbA5dR6cL1gS4w"
        # アプリケーションのルートディレクトリを取得
        self.app_root = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
        self.data_dir = os.path.join(self.app_root, "data")
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)

    def fetch_questionnaire_data(self, start_date=None, end_date=None) -> dict:
        try:
            if start_date is None:
                start_date = datetime.date.today()
            if end_date is None:
                end_date = datetime.date.today()
                
            # 日付を文字列に変換
            start_date_str = start_date.strftime("%Y-%m-%d")
            end_date_str = end_date.strftime("%Y-%m-%d")
            
        except ValueError:
            return {"error": "日付は YYYY-MM-DD 形式で入力してください。"}
            
        params = {
            'start_date': start_date_str,
            'end_date': end_date_str,
            'token': self.token
        }

        try:
            # リクエストの送信
            response = requests.get(self.API_URL, params=params)
            response.raise_for_status()
            
            # レスポンスの処理
            self.response_data = response.json()
            
            # タイムスタンプを日本時間に変換
            for record in self.response_data:
                record['timestamp'] = self._convert_to_jst(record['timestamp']).strftime('%Y-%m-%d %H:%M:%S')
                record["Q1"] = self._clean_text(record["Q1"])
                record["Q2"] = self._convert_to_jst(record["Q2"]).strftime("%Y-%m-%d")
                record["Q3"] = self._clean_text(record["Q3"])
                record["Q8"] = self._clean_text(record["Q8"])
                record["Q9"] = self._clean_text(record["Q9"])
                record["Q11"] = self._clean_text(record["Q11"])
                record["Q17"] = self._clean_text(record["Q17"])            
            
            # 日付ごとにデータを保存
            self._save_questionnaire_data_by_date(self.response_data)
            
            return self.response_data

        except requests.exceptions.RequestException as e:
            return {"error": f"通信エラー: {str(e)}"}
        except json.JSONDecodeError:
            return {"error": "レスポンスの解析に失敗しました。"}
        
    def _clean_text(self,text):

        if not isinstance(text, str):
            return text
        # 改行コードを半角スペースに置換
        text = re.sub(r'\r\n|\r|\n', ' ', text)
        # 全角スペースを半角スペースに置換
        text = re.sub(r'\u3000', ' ', text)
        # タブを半角スペースに置換
        text = re.sub(r'\t', ' ', text)
        # 連続するスペースを1つに置換
        text = re.sub(r'\s+', ' ', text)
        # 前後のスペースを削除
        text = text.strip()
        return text

    def _convert_to_jst(self, utc_time_str):
        try:
            # UTCの時間文字列をdatetimeオブジェクトに変換
            utc_time = datetime.datetime.fromisoformat(utc_time_str.replace('Z', '+00:00'))
            # 日本時間に変換
            jst = pytz.timezone('Asia/Tokyo')
            jst_time = utc_time.astimezone(jst)
            return jst_time
        except (ValueError, TypeError):
            return utc_time_str
        
    def _save_questionnaire_data_by_date(self, data):
        """
        日付ごとにアンケートデータを保存する
        """
        # データを日付ごとにグループ化
        date_groups = {}
        for record in data:
            if 'timestamp' in record:
                try:
                    date_str = record['timestamp'].split()[0]
                    record["completed"] = False
                    if date_str not in date_groups:
                        date_groups[date_str] = []
                    date_groups[date_str].append(record)
                except (ValueError, TypeError):
                    continue

        # 日付ごとにファイルに保存
        for date_str, records in date_groups.items():
            file_name = f"{date_str}.pickle"
            file_path = os.path.join(self.data_dir, file_name)
            
            # データを上書き保存
            with open(file_path, "wb") as f:
                pickle.dump(records, f)

class PrintData:
    def __init__(self, data_text):
        self.data_text = data_text

    def print_data(self):
        try:
            # 日本語フォントの設定
            font_path = "C:/Windows/Fonts/msgothic.ttc"
            pdfmetrics.registerFont(TTFont('MSGothic', font_path))
            
            # スタイルの設定
            styles = getSampleStyleSheet()
            styles.add(ParagraphStyle(
                name='Japanese',
                fontName='MSGothic',
                fontSize=10,
                leading=12
            ))
            
            # 一時ファイルを作成
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
                temp_file_path = tmp.name
            
            # PDFドキュメントの作成
            doc = SimpleDocTemplate(
                temp_file_path,
                pagesize=A4,
                rightMargin=72,
                leftMargin=72,
                topMargin=72,
                bottomMargin=72
            )
            
            # コンテンツの作成
            content = []
            
            # タイトル
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontName='MSGothic',
                fontSize=16,
                spaceAfter=30
            )
            content.append(Paragraph('問診票詳細', title_style))
            
            # 印刷日時
            date_style = ParagraphStyle(
                'CustomDate',
                parent=styles['Normal'],
                fontName='MSGothic',
                fontSize=10,
                spaceAfter=20
            )
            content.append(Paragraph(
                f'印刷日時: {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}',
                date_style
            ))
            
            # データの整形
            data = []
            for line in self.data_text.split('\n'):
                if ': ' in line:
                    key, value = line.split(': ', 1)  # 最初の': 'で分割
                    # キーと値をParagraphオブジェクトに変換
                    key_para = Paragraph(key, styles['Japanese'])
                    value_para = Paragraph(value, styles['Japanese'])
                    data.append([key_para, value_para])
            
            # テーブルの作成
            if data:  # データが存在する場合のみテーブルを作成
                # テーブルの幅を調整（A4用紙の幅に合わせる）
                available_width = A4[0] - 144  # 左右のマージン（72pt * 2）を引いた幅
                table = Table(data, colWidths=[available_width * 0.3, available_width * 0.7])
                table.setStyle(TableStyle([
                    ('FONT', (0, 0), (-1, -1), 'MSGothic'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
                    ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                    ('ALIGN', (0, 0), (0, -1), 'LEFT'),  # 左列は左揃え
                    ('ALIGN', (1, 0), (1, -1), 'LEFT'),  # 右列は左揃え
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # 上揃えに変更
                    ('PADDING', (0, 0), (-1, -1), 6),
                    ('LEFTPADDING', (0, 0), (-1, -1), 6),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                    ('TOPPADDING', (0, 0), (-1, -1), 6),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ]))
                content.append(table)
            else:
                content.append(Paragraph('データがありません', styles['Japanese']))
            
            # PDFの生成
            doc.build(content)
            
            try:
                # 利用可能なプリンターを取得
                printers = [printer[2] for printer in win32print.EnumPrinters(2)]
                if not printers:
                    raise Exception("利用可能なプリンターが見つかりません")
                
                # デフォルトプリンターを取得
                default_printer = win32print.GetDefaultPrinter()
                if not default_printer:
                    default_printer = printers[0]  # 最初のプリンターを使用
                
                # PDFを印刷
                if os.name == 'nt':  # Windowsの場合
                    import subprocess
                    # Adobe Readerを使用して印刷
                    acrobat_path = r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
                    if not os.path.exists(acrobat_path):
                        acrobat_path = r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
                    
                    if os.path.exists(acrobat_path):
                        subprocess.run([acrobat_path, '/t', temp_file_path, default_printer], shell=True)
                    else:
                        # Adobe Readerが見つからない場合は、デフォルトのPDFビューアで開く
                        os.startfile(temp_file_path)
                else:
                    raise Exception("この機能はWindowsでのみ利用可能です")
                
            except Exception as e:
                print(f"プリンターエラー: {str(e)}")
                # エラー時はPDFを開く
                os.startfile(temp_file_path)
            
            # 一時ファイルを削除（少し遅延を入れて印刷が開始されるのを待つ）
            import threading
            def delete_temp_file():
                import time
                time.sleep(5)  # 5秒待機
                try:
                    os.unlink(temp_file_path)
                except:
                    pass
            
            threading.Thread(target=delete_temp_file).start()
            
        except Exception as e:
            print(f"印刷エラー: {str(e)}")
