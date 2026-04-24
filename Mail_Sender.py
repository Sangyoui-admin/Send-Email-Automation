import os
import sys
from sys import argv
import csv
import re
from datetime import datetime
import time

import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTextEdit, QPushButton, QVBoxLayout, QWidget,
                            QFileDialog, QMessageBox, QSpinBox, QColorDialog, QInputDialog, QAction, 
                            QHBoxLayout, QDialog, QFormLayout, QLineEdit, QDialogButtonBox,QListWidget,
                            QListWidgetItem,QTableWidget,QTableWidgetItem,QLabel,QFontComboBox,QCheckBox,
                            QComboBox,QSizePolicy,QGridLayout,QFrame,QTextBrowser)
from PyQt5.QtGui import QTextCursor, QTextCharFormat, QFont, QColor, QIcon, QKeySequence
from PyQt5.QtCore import Qt,QSize,QStandardPaths
import win32com.client as win32
from PIL import Image

dir_name = os.path.dirname(os.path.abspath(argv[0])) # exe化した際は実行している方のファイルのディレクトリ
cwd = os.path.dirname(__file__)# exe化した際は一時ファイルの方のpyファイルのディレクトリ
desktop_path = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation) # デスクトップのパスを取得
headers = []  # ヘッダーを予め設定
recipients_data = []  # リストとして宛先データを保持

class ImageSizeDialog(QDialog):
    def __init__(self, original_width, original_height, parent=None):
        super().__init__(parent)
        self.setWindowTitle("画像サイズの設定")
        self.layout = QFormLayout(self)

        self.original_width = original_width
        self.original_height = original_height

        self.widthInput = QLineEdit(self)
        self.heightInput = QLineEdit(self)

        self.widthInput.setText(str(original_width))
        self.heightInput.setText(str(original_height))

        self.widthInput.textChanged.connect(self.on_width_changed)
        self.heightInput.textChanged.connect(self.on_height_changed)

        self.layout.addRow("幅 (ピクセル):", self.widthInput)
        self.layout.addRow("高さ (ピクセル):", self.heightInput)

        self.buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)
        self.layout.addWidget(self.buttonBox)
        
        # 初期化時に recipients_data を空のリストとして定義する
        self.recipients_data = []  # メールアドレスと他の列を保持するリスト
        self.sender_email = None

    def on_width_changed(self, text):
        if not self.is_changing:
            self.is_changing = True
            try:
                new_width = int(text)
                new_height = int(new_width * self.original_height / self.original_width)
                self.heightInput.blockSignals(True)  # 一時的にシグナルをブロック
                self.heightInput.setText(str(new_height))
                self.heightInput.blockSignals(False)  # シグナルのブロックを解除
            except ValueError:
                pass
            self.is_changing = False

    def on_height_changed(self, text):
        if not self.is_changing:
            self.is_changing = True
            try:
                new_height = int(text)
                new_width = int(new_height * self.original_width / self.original_height)
                self.widthInput.blockSignals(True)  # 一時的にシグナルをブロック
                self.widthInput.setText(str(new_width))
                self.widthInput.blockSignals(False)  # シグナルのブロックを解除
            except ValueError:
                pass
            self.is_changing = False

    def get_size(self):
        try:
            width = int(self.widthInput.text())
            height = int(self.heightInput.text())
            return width, height
        except ValueError:
            return None

    def __init__(self, original_width, original_height, parent=None):
        super().__init__(parent)
        self.setWindowTitle("画像サイズの設定")
        self.layout = QFormLayout(self)

        self.original_width = original_width
        self.original_height = original_height

        self.widthInput = QLineEdit(self)
        self.heightInput = QLineEdit(self)

        self.widthInput.setText(str(original_width))
        self.heightInput.setText(str(original_height))

        self.is_changing = False

        self.widthInput.textChanged.connect(self.on_width_changed)
        self.heightInput.textChanged.connect(self.on_height_changed)

        self.layout.addRow("幅 (ピクセル):", self.widthInput)
        self.layout.addRow("高さ (ピクセル):", self.heightInput)

        self.buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)
        self.layout.addWidget(self.buttonBox)

#===========================================================================================
class CSVViewer(QDialog):
    # csvファイルの内容を表示する別ウィンドウの処理
    def __init__(self, data, headers, parent=None):
        super().__init__(parent)
        self.setWindowTitle("CSV Viewer")
        self.resize(1400, 400)

        self.layout = QVBoxLayout(self)

        # Table widget to display CSV content
        self.table_widget = QTableWidget(self)
        self.table_widget.setRowCount(len(data))
        self.table_widget.setColumnCount(len(headers))
        self.table_widget.setHorizontalHeaderLabels(headers)
        
        self.fill_table(data) # csvの内容をリストに記載

        self.set_column_widths(headers) # リストの表示幅を指定
        
        # 「テスト表示ボタン」を追加
        self.test_display_button = QPushButton("プレビュー")
        self.test_display_button.clicked.connect(self.on_test_display)
        self.layout.addWidget(self.table_widget)
        self.layout.addWidget(self.test_display_button)

        # 手動更新用のフラグ（メール送信時などのプログラム更新と区別）
        self.manual_update = True
        
        # セルの変更を検知
        self.table_widget.itemChanged.connect(self.on_item_changed)

        # セルがクリックされたことを検知（列番号取得用）
        self.table_widget.cellClicked.connect(self.on_cell_clicked)
        
        # ヘッダーを保持
        self.headers = headers

        # 列幅を内容に基づいて自動調整
        self.table_widget.resizeColumnsToContents()
        
    def on_cell_clicked(self, row, column):
        # セルがクリックされたときの処理
        # 'Email'列のインデックスを取得
        if 'Email' in self.headers:
            email_col_index = self.headers.index('Email')

            # クリックされた行のEmail列にあるメールアドレスを取得
            email_item = self.table_widget.item(row, email_col_index)

            if email_item is not None:
                email_address = email_item.text()

                # ボタンのテキストにメールアドレスを表示
                new_button_text = f"{email_address}宛て原稿のプレビュー"
                self.test_display_button.setText(new_button_text)
            else:
                self.test_display_button.setText("メールアドレスが見つかりません")
        else:
            self.test_display_button.setText("Email列がありません")

        # ---------------------------------------------------------
    def on_test_display(self):
        # プレビュー画面を表示した際の処理
        selected_row = self.table_widget.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, "エラー", "行を選択してください。")
            return

        # 選択された行のデータを取得
        headers = [self.table_widget.horizontalHeaderItem(col_idx).text() for col_idx in range(self.table_widget.columnCount())]
        row_data = [self.table_widget.item(selected_row, col_idx).text() for col_idx in range(self.table_widget.columnCount())]

        # CSVのヘッダーと行のデータを基に変数を作成
        variables = {}
        for header, value in zip(headers, row_data):
            fix_header = "{"  + header + "}"
            variables[fix_header] = value # ヘッダー名を {header} 形式で変数として登録

        # 親ウィンドウが存在するか確認
        parent_window = self.parent()
        if parent_window and hasattr(parent_window, 'apply_variables_to_template'):
            
            # メイン画面のインスタンスを参照し、変数にデータを反映させたタイトルと本文を生成
            personalized_title = parent_window.apply_variables_to_template(parent_window.MailTitleEdit.text().strip(), variables)  # タイトル反映
            result_text = parent_window.apply_variables_to_template(parent_window.textEdit.toHtml().strip(), variables)  # 本文反映

            # 署名を追加
            result_text = result_text + '\n'  + parent_window.signEdit.toHtml()

            # プレビューウィンドウを表示
            self.show_test_window(personalized_title, result_text)
        else:
            QMessageBox.critical(self, "エラー", "メイン画面が見つかりません。")
            
    def show_test_window(self, title, content):
        # プレビューウィンドウを表示
        preview_window = PreviewWindow(title, content, parent=self)
        preview_window.exec_()
    # ---------------------------------------------------------

    def fill_table(self, data):
        # ステータスの状態に応じて色をつける
        for row_idx, row_data in enumerate(data):
            for col_idx, col_data in enumerate(row_data):
                item = QTableWidgetItem(str(col_data))
                self.table_widget.setItem(row_idx, col_idx, item)
                
                # Check if this is the Status column and if the value is '送信済み'
                if col_data == '送信済み' and self.table_widget.horizontalHeaderItem(col_idx).text() == 'Status':
                    item.setBackground(QColor(248,203,173))
                    # item.setForeground(QColor('white'))  # Optional: change text color for better visibility
                elif col_data == '下書き作成済み' and self.table_widget.horizontalHeaderItem(col_idx).text() == 'Status':
                    item.setBackground(QColor(255,217,102))
                elif col_data == 'アドレスを確認してください' and self.table_widget.horizontalHeaderItem(col_idx).text() == 'Status':
                    item.setBackground(QColor(204, 204, 255))

    def update_data(self, data):
        # メール送信時のプログラムによる更新なのでフラグをオフにする
        self.manual_update = False
        
        # ウィンドウに表示されているリストの内容を更新
        self.table_widget.setRowCount(len(data))
        self.fill_table(data)
        
        # 再度手動更新を可能にする
        self.manual_update = True
    
    def on_item_changed(self, item):
        global recipients_data
        # 手動更新のみ処理する（メール送信時は無視）
        if not self.manual_update:
            return

        # 変更された行の内容を反映させる
        row = item.row()
        col = item.column()

        # 変更されたセルの内容を取得
        new_value = item.text()
        # Status列の場合、セルの内容に応じて色を変更
        header_text = self.table_widget.horizontalHeaderItem(col).text()
        if header_text == 'Status':
            if new_value == '送信済み':
                item.setBackground(QColor(248, 203, 173))  # 送信済みの色
            elif new_value == '下書き作成済み':
                item.setBackground(QColor(255, 217, 102))  # 下書き作成済みの色
            elif new_value == 'アドレスを確認してください':
                item.setBackground(QColor(204, 204, 255))  # 送信失敗の色
            elif new_value == '':  # 空欄の場合は背景色をリセット
                item.setBackground(QColor(255, 255, 255))  # 白に戻す

        # TableWidgetから全行のデータを取得
        updated_data = []
        for row_idx in range(self.table_widget.rowCount()):
            row_data = []
            for col_idx in range(self.table_widget.columnCount()):
                cell_item = self.table_widget.item(row_idx, col_idx)

                # セルが空の場合や空文字列の場合に NaN を設定
                if cell_item is None or cell_item.text().strip() == '':
                    row_data.append(float('nan'))  # 無効値として NaN を追加
                else:
                    text_value = cell_item.text()
                    # 'nan' の文字列を無効値の NaN に変換
                    if text_value == str(float('nan')):
                        row_data.append(float('nan'))
                    else:
                        row_data.append(text_value)  # 通常のテキスト値を追加

            updated_data.append(row_data)

        # CSVファイルを更新
        try:
            self.save_to_csv(updated_data)
            df = pd.read_csv(file_name)
            recipients_data = []  # データリストをクリア
            recipients_data = df.to_dict('records')  # 全行データをリストとして保持
            #recipients_data = updated_data
        except Exception as e:
            print(f"CSVの更新に失敗しました: {e}")
            
    def save_to_csv(self, data):
        # CSVファイルにデータを書き込む
        try:
            with open(file_name, 'w', newline='', encoding='utf-8-sig')as csvfile:
                writer = csv.writer(csvfile)
                # ヘッダーを書き込む
                headers = [self.table_widget.horizontalHeaderItem(col).text() for col in range(self.table_widget.columnCount())]
                writer.writerow(headers)
                # データを書き込む
                writer.writerows([[cell if cell is not None else "" for cell in row] for row in data])
        except Exception as e:
            raise Exception(f"CSVファイルの書き込みに失敗しました: {e}")

    def set_column_widths(self, headers):
        # リスト一覧画面の表示幅を指定
        for idx, header in enumerate(headers):
            if header == 'Email':
                self.table_widget.setColumnWidth(idx, 250)  # Email列は幅200にする
            else:
                self.table_widget.setColumnWidth(idx, 100)  # それ以外は100にする

    def clear(self):
        self.table_widget.clearContents() # いらん内容を全消し処理

#===========================================================================================
# プレビュー表示用画面
class PreviewWindow(QDialog):
    def __init__(self, title, content, parent=None):
        super().__init__(parent)
        self.setWindowTitle("プレビュー画面")
        self.resize(1500, 1000)
        
        layout = QVBoxLayout(self)

        # タイトル用のテキストボックス
        self.title_edit = QTextBrowser(self)  # QTextBrowserに変更
        self.title_edit.setPlainText(title)   # タイトルを表示
        self.title_edit.setReadOnly(True)     # 編集不可にする
        self.title_edit.setFixedHeight(self.fontMetrics().height() * 2) 
        layout.addWidget(QLabel("タイトルプレビュー:"))
        layout.addWidget(self.title_edit)

        # コンテンツ用のテキストボックス
        self.content_edit = QTextBrowser(self)  # QTextBrowserに変更
        self.content_edit.setHtml(content)      # コンテンツを表示
        self.content_edit.setOpenExternalLinks(True)  # 外部リンクを有効にする
        layout.addWidget(QLabel("本文プレビュー:"))
        layout.addWidget(self.content_edit)
        
        # 閉じるボタン
        close_button = QPushButton("閉じる", self)
        close_button.clicked.connect(self.close)
        layout.addWidget(close_button)

#===========================================================================================
# 埋め込みリンク指定用画面
class InsertLinkDialog(QDialog):
    def __init__(self, headers, parent=None):
        super().__init__(parent)
        self.setWindowTitle('リンクを挿入')
        self.selected_url = ""

        layout = QVBoxLayout(self)

        # テキスト入力フィールド
        self.url_input = QLineEdit(self)
        self.url_input.setPlaceholderText('URLを入力してください')

        # 「URL」を含むヘッダーの選択用コンボボックス
        self.url_combo_box = QComboBox(self)
        self.url_combo_box.addItems([header for header in headers if 'URL' in header or 'url' in header])

        # コンボボックスから選択されたURLをテキストフィールドに挿入するボタン
        select_button = QPushButton('ヘッダーから選択')
        select_button.clicked.connect(self.insert_selected_url)

        # ダイアログボタン (OK/キャンセル)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        # レイアウトにウィジェットを追加
        layout.addWidget(QLabel("URLを入力するか、ヘッダーから選択してください:"))
        layout.addWidget(self.url_input)
        
        # コンボボックスとボタンのレイアウト
        combo_layout = QHBoxLayout()
        combo_layout.addWidget(self.url_combo_box)
        combo_layout.addWidget(select_button)
        layout.addLayout(combo_layout)
        
        layout.addWidget(button_box)

    def insert_selected_url(self):
        # コンボボックスから選択したURLをテキストフィールドに挿入
        selected_header = self.url_combo_box.currentText()
        if selected_header:
            selected_header = "{" + selected_header + "}"
            self.url_input.setText(selected_header)

    def get_url(self):
        return self.url_input.text()

#===========================================================================================
class HeaderWindow(QDialog):
    def __init__(self, headers, parent=None):
        super().__init__(parent)
        self.setWindowTitle('差込文章リスト')

        # グリッドレイアウトを使用
        self.layout = QGridLayout(self)

        # ヘッダーをボタンとして追加
        max_items_per_column = 6  # 1列あたりの最大ボタン数
        for index, header in enumerate(headers):
            button = QPushButton(header, self)
            button.clicked.connect(self.handle_click)  # クリックイベントを接続

            # indexから行と列を計算してグリッドに配置
            row = index % max_items_per_column
            col = index // max_items_per_column
            self.layout.addWidget(button, row, col)

        # 条件を満たす場合に曜日ボタンを追加
        if all(key in headers for key in ['Year', 'Month', 'Dates']):
            dotw_button = QPushButton('Dotw（曜日）', self)
            dotw_button.clicked.connect(self.handle_dotw_click)  # 曜日ボタンクリックイベントを接続

            # 曜日ボタンをグリッドに配置（最後の行の次の位置）
            row = len(headers) % max_items_per_column
            col = len(headers) // max_items_per_column
            self.layout.addWidget(dotw_button, row, col)

    def handle_click(self):
        clicked_button = self.sender()
        header_name = clicked_button.text()
        self.parent().insert_header_into_textbox(header_name)  # メインウィンドウのテキストボックスに挿入

    def handle_dotw_click(self):
        # Dotwボタンがクリックされたときの処理
        self.parent().insert_header_into_textbox('Dotw')  # Dotwをテキストボックスに挿入

#===========================================================================================
class EmailSenderApp(QMainWindow):
    global recipients_data
    def __init__(self):
        super().__init__()
        self.templates = {}  # これが重要です。最初に templates を定義します。
        template_path = os.path.join(dir_name, 'templates.txt')

        # ウィンドウのアイコンを設定
        exe_icon_path = os.path.join(cwd, "image\\exe_logo.png")
        self.setWindowIcon(QIcon(exe_icon_path))

        # 署名 (HTML対応)
        self.signEdit = QTextEdit(self)
        self.signEdit.setAcceptRichText(True)
        self.signEdit.setPlaceholderText("メールの署名を記載")
        
        # サイズポリシーを設定して高さを制限
        self.signEdit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.set_signature_height(12)  # 例えば、4行分の高さを設定
        font = self.signEdit.font()
        font.setPointSize(12)
        self.signEdit.setFont(font)
        
        self.load_templates_from_file(template_path) # テンプレートテキストを読み込む
        
        self.initUI()
        self.recipients = []
        self.sender_email = None
        self.attachments = []  # 添付ファイルリストを初期化
        self.attachment_list.setFixedHeight(0)  # 添付ファイルがない場合は高さ0

    def initUI(self):
        self.setWindowTitle("メール配信一括君NEO 2.0")
        self.setGeometry(100, 100, 800, 1200)

        # メインウィジェットとレイアウト
        mainWidget = QWidget(self)
        self.setCentralWidget(mainWidget)
        mainLayout = QVBoxLayout(mainWidget)
        
        # 上部の1行目ツールバー（送信者アカウント指定用）
        accountLayout = QHBoxLayout()
        self.senderLineEdit = QLineEdit(self)
        self.senderLineEdit.setPlaceholderText("このアカウントを送信者として送る")
        accountLayout.addWidget(self.senderLineEdit)

        # 2行目メールタイトル
        MailTitleLayout = QHBoxLayout()
        self.MailTitleEdit = QLineEdit(self)
        self.MailTitleEdit.setPlaceholderText("メールタイトルを指定してください")
        MailTitleLayout.addWidget(self.MailTitleEdit)

        # ---------------------------------------------------------
        # 下部の3行目ツールバー（フォントサイズ変更など）
        toolbarLayout = QHBoxLayout()

        # フォントの種類選択エリア
        self.font_label = QLabel("Font:", self)
        toolbarLayout.addWidget(self.font_label)
        self.font_combo = QFontComboBox(self)
        self.font_combo.currentFontChanged.connect(self.change_font_type)
        toolbarLayout.addWidget(self.font_combo)

        # 文字サイズ変更スピンボックス
        self.fontSizeSpinBox = QSpinBox(self)
        self.fontSizeSpinBox.setMinimum(8)
        self.fontSizeSpinBox.setMaximum(72)
        self.fontSizeSpinBox.setValue(12)
        self.fontSizeSpinBox.valueChanged.connect(self.change_font_size)
        toolbarLayout.addWidget(self.fontSizeSpinBox)

        # 色変更ボタン
        color_icon_path = os.path.join(cwd, 'image\\color_icon.png')
        colorButton = QPushButton(self)
        colorButton.setIcon(QIcon(color_icon_path))  # アイコンを設定
        colorButton.clicked.connect(self.change_color)
        colorButton.setToolTip('フォントの色を変更')
        toolbarLayout.addWidget(colorButton)

        # 太字ボタン
        bold_icon_path = os.path.join(cwd, 'image\\bold_icon.png')
        boldButton = QPushButton(self)
        boldButton.setIcon(QIcon(bold_icon_path))  # アイコンを設定
        boldButton.clicked.connect(self.toggle_bold)
        boldButton.setToolTip('太字')
        toolbarLayout.addWidget(boldButton)

        # 下線ボタン
        underline_icon_path = os.path.join(cwd, 'image\\underline_icon.png')
        underlineButton = QPushButton(self)
        underlineButton.setIcon(QIcon(underline_icon_path))  # アイコンを設定
        underlineButton.clicked.connect(self.toggle_underline)
        underlineButton.setToolTip('下線')
        toolbarLayout.addWidget(underlineButton)

        # リンクボタン
        link_icon_path = os.path.join(cwd, 'image\\link_icon.png')
        linkButton = QPushButton(self)
        linkButton.setIcon(QIcon(link_icon_path))  # アイコンを設定
        linkButton.clicked.connect(self.insert_link)
        linkButton.setToolTip('文章にリンクを設定')
        toolbarLayout.addWidget(linkButton)

        # 画像貼り付けボタン
        image_icon_path = os.path.join(cwd, 'image\\image_icon.png')
        imageButton = QPushButton(self)
        imageButton.setIcon(QIcon(image_icon_path))  # アイコンを設定
        imageButton.clicked.connect(self.insert_image)
        imageButton.setToolTip('画像を貼り付け')
        toolbarLayout.addWidget(imageButton)

        # 添付ファイル貼り付けボタン
        temp_icon_path = os.path.join(cwd, 'image\\temp_icon.png')
        tempButton = QPushButton(self)
        tempButton.setIcon(QIcon(temp_icon_path))  # アイコンを設定
        tempButton.clicked.connect(self.add_attachment)
        tempButton.setToolTip('添付ファイルを追加')
        toolbarLayout.addWidget(tempButton)

        # 仕切り線(QFrame)を作成
        separator1 = QFrame()
        separator1.setFrameShape(QFrame.VLine)  # 垂直の線を設定
        separator1.setFrameShadow(QFrame.Sunken)  # 影のスタイルを設定
        toolbarLayout.addWidget(separator1)

        # 左揃えボタン
        Left_align_icon_path = os.path.join(cwd, 'image\\Left_align_icon.png')
        Left_align_Button = QPushButton(self)
        Left_align_Button.setIcon(QIcon(Left_align_icon_path))  # アイコンを設定
        Left_align_Button.clicked.connect(self.apply_Left_alignment)
        Left_align_Button.setToolTip('文章を左寄せ')
        toolbarLayout.addWidget(Left_align_Button)

        # 中央揃えボタン
        Center_align_icon_path = os.path.join(cwd, 'image\\Center_align_icon.png')
        Center_align_Button = QPushButton(self)
        Center_align_Button.setIcon(QIcon(Center_align_icon_path))  # アイコンを設定
        Center_align_Button.clicked.connect(self.apply_Center_alignment)
        Center_align_Button.setToolTip('文章を中央寄せ')
        toolbarLayout.addWidget(Center_align_Button)

        # 右揃えボタン
        Right_align_icon_path = os.path.join(cwd, 'image\\Right_align_icon.png')
        Right_align_Button = QPushButton(self)
        Right_align_Button.setIcon(QIcon(Right_align_icon_path))  # アイコンを設定
        Right_align_Button.clicked.connect(self.apply_Right_alignment)
        Right_align_Button.setToolTip('文章を右寄せ')
        toolbarLayout.addWidget(Right_align_Button)

        
        # ---------------------------------------------------------
        # 添付ファイルの表示リスト
        self.attachment_list = QListWidget(self)
        self.attachment_list.setFixedHeight(0)  # 初期は高さ0

        # ---------------------------------------------------------
        # メールエディタ (HTML対応)
        self.textEdit = QTextEdit(self)
        self.textEdit.setAcceptRichText(True)
        self.textEdit.setPlaceholderText("メールの本文を記載")
        
        # 初期フォントサイズを12に設定
        font = self.textEdit.font()
        font.setPointSize(12)
        self.textEdit.setFont(font)
                
        # ---------------------------------------------------------
        # CSV確認ボタンエリア1
        CSVLayout1 = QHBoxLayout()

        # CSV読み込みボタン
        self.csvButton = QPushButton('CSVを読み込む📗', self)
        self.csvButton.clicked.connect(self.load_csv)
        CSVLayout1.addWidget(self.csvButton)

        # CSVテンプレートボタン
        self.csvButton = QPushButton('CSVテンプレートを作成🛠️', self)
        self.csvButton.clicked.connect(self.create_template)
        CSVLayout1.addWidget(self.csvButton)

        # ---------------------------------------------------------
        # CSV確認ボタンエリア2
        CSVLayout2 = QHBoxLayout()

        # 送信先リストを表示するボタン
        self.reopen_button = QPushButton("送信先リストを表示👥", self)
        self.reopen_button.clicked.connect(self.reopen_csv_viewer)
        self.reopen_button.setEnabled(False)
        CSVLayout2.addWidget(self.reopen_button)

        # 変数一覧リストを表示するボタン
        self.header_button = QPushButton("差込文章リストを表示💬", self)
        self.header_button.clicked.connect(self.show_header_window)
        self.header_button.setEnabled(False)
        CSVLayout2.addWidget(self.header_button)

        # ---------------------------------------------------------
        # テンプレート選択用のプルダウンメニュー
        self.templateComboBox = QComboBox(self)
        self.templateComboBox.addItem("テンプレートを選択")
        
        for template_name in self.templates.keys():
            self.templateComboBox.addItem(template_name)
            
        self.templateComboBox.currentIndexChanged.connect(self.apply_template)
        
        # --------------------------------------------------------- 
        # メール送信エリア        
        # 下書き保存のチェックボックス
        self.draft_checkbox = QCheckBox("下書きとして作成 ※このチェックを外すと直接メール送信されるようになります", self)
        self.draft_checkbox.setChecked(True)  # 初期状態でチェックを入れる

        # メール送信ボタン
        self.sendButton = QPushButton('メールの下書きを作成する📝', self)
        self.sendButton.clicked.connect(self.send_email)

        self.draft_checkbox.stateChanged.connect(self.update_button_text)  # 状態変更に応じてボタンテキストを更新

        # テストメール送信ボタン
        self.sendtestButton = QPushButton('自分のアカウントにテストメール送信する🔍', self)
        self.sendtestButton.clicked.connect(self.send_testemail)

        # ---------------------------------------------------------
        # レイアウトに追加
        mainLayout.addLayout(accountLayout)   # 送信者アカウント入力
        mainLayout.addLayout(MailTitleLayout) # メールタイトル指定入力
        mainLayout.addLayout(toolbarLayout)   # フォント変更や太字などのツール
        mainLayout.addWidget(self.templateComboBox)
        mainLayout.addWidget(self.attachment_list)  # 添付ファイルリストを本文エディタの上に追加
        mainLayout.addWidget(self.textEdit)   # メール本文のエディタ
        mainLayout.addWidget(self.signEdit)   # メール署名のエディタ
        mainLayout.addLayout(CSVLayout1)       # CSVファイルの読み込み関連
        mainLayout.addLayout(CSVLayout2)       # リストや差込文章関連
        mainLayout.addWidget(self.draft_checkbox) # 下書き作成チェック
        mainLayout.addWidget(self.sendButton) # メール送信ボタン
        mainLayout.addWidget(self.sendtestButton) # テストメール送信ボタン

        # キーボードショートカットの設定
        self.textEdit.addAction(self.create_shortcut_action(Qt.CTRL + Qt.Key_B, self.toggle_bold))
        self.textEdit.addAction(self.create_shortcut_action(Qt.CTRL + Qt.Key_U, self.toggle_underline))

        # self.recipients_data = []  # リストとして宛先データを保持
        self.csv_viewer = None  # CSVViewerインスタンスを保存するための変数

    def update_button_text(self):
        # 下書きチェックが変更された際の処理
        if self.draft_checkbox.isChecked():
            self.sendButton.setText("メールの下書きを作成する📝")
        else:
            self.sendButton.setText("メールを送信する📨")

    def create_shortcut_action(self, key_sequence, slot):
        # ショートカットアクションの処理
        action = QAction(self)
        action.setShortcut(QKeySequence(key_sequence))
        action.triggered.connect(slot)
        return action

    def change_font_type(self):
        cursor = self.textEdit.textCursor()
        if cursor.hasSelection():
            format = QTextCharFormat()
            format.setFontFamily(self.font_combo.currentFont().family())
            cursor.mergeCharFormat(format)

    def change_font_size(self):
        cursor = self.textEdit.textCursor()
        if cursor.hasSelection():
            format = QTextCharFormat()
            format.setFontPointSize(self.fontSizeSpinBox.value())
            cursor.mergeCharFormat(format)

    def change_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            cursor = self.textEdit.textCursor()
            if cursor.hasSelection():
                format = QTextCharFormat()
                format.setForeground(QColor(color))
                cursor.mergeCharFormat(format)

    def toggle_bold(self):
        cursor = self.textEdit.textCursor()
        if cursor.hasSelection():
            format = cursor.charFormat()
            format.setFontWeight(QFont.Bold if not format.fontWeight() == QFont.Bold else QFont.Normal)
            cursor.mergeCharFormat(format)

    def toggle_underline(self):
        cursor = self.textEdit.textCursor()
        if cursor.hasSelection():
            format = cursor.charFormat()
            format.setFontUnderline(not format.fontUnderline())
            cursor.mergeCharFormat(format)

    def insert_link(self):
        dialog = InsertLinkDialog(headers, self)
        if dialog.exec_() == QDialog.Accepted:
            url = dialog.get_url()
            if url:
                cursor = self.textEdit.textCursor()
                if cursor.hasSelection():
                    format = QTextCharFormat()
                    format.setAnchor(True)
                    format.setAnchorHref(url)
                    format.setForeground(QColor('blue'))
                    format.setFontUnderline(True)
                    cursor.mergeCharFormat(format)

    def insert_image(self):
        # 画像を選択する
        image_path, _ = QFileDialog.getOpenFileName(self, '画像を選択', desktop_path, '画像ファイル (*.png *.jpg *.bmp)')
        if image_path:
            # 画像のオリジナルサイズを取得
            original_size = self.get_image_size(image_path)
            if original_size:
                dialog = ImageSizeDialog(*original_size, self)
                if dialog.exec_() == QDialog.Accepted:
                    size = dialog.get_size()
                    if size:
                        width, height = size
                        # パスのバックスラッシュをスラッシュに変換
                        image_path_for_html = image_path.replace('\\', '/')
                        
                        # 画像パスをfile://プロトコルに変換してHTMLタグを生成 (コンソール用)
                        html = f'<img src="file:///{image_path_for_html}" width="{width}" height="{height}"/>'
                        cursor = self.textEdit.textCursor()
                        cursor.insertHtml(html)

                        # 埋め込み画像用に画像パスとContent-IDを辞書に保持
                        image_filename = os.path.basename(image_path)
                        content_id = f'cid:{image_filename}'
                        if not hasattr(self, 'embedded_images'):
                            self.embedded_images = {}
                        self.embedded_images[content_id] = image_path

    def get_image_size(self, image_path):
        try:
            with Image.open(image_path) as img:
                return img.width, img.height
        except Exception as e:
            QMessageBox.critical(self, 'エラー', f'画像の読み込みに失敗しました: {e}')
            return None
        
    def add_attachment(self):
        # 添付ファイルの追加と表示
        file_paths, _ = QFileDialog.getOpenFileNames(self, 'ファイルを選択', desktop_path, 'すべてのファイル (*.*)')
        if file_paths:
            for file in file_paths: # 選択されたファイルを添付ファイルリストに追加
                self.attachments.append(file)
                self.update_attachment_list()
            QMessageBox.information(self, '添付ファイル追加', f'{len(file_paths)}件のファイルが追加されました。')
    
    def update_attachment_list(self):
        # 添付ファイルリストを更新
        self.attachment_list.clear()  # 現在のリストをクリア
        for file in self.attachments:
            item = QListWidgetItem(file)
            self.attachment_list.addItem(item)

            # 削除ボタンの作成と追加
            widget = QWidget()
            layout = QHBoxLayout(widget)
            file_label = QLineEdit(file)
            file_label.setReadOnly(True)
            remove_btn = QPushButton('削除')
            remove_btn.clicked.connect(lambda _, f=file: self.remove_attachment(f))
            
            layout.addWidget(file_label)
            layout.addWidget(remove_btn)
            layout.setContentsMargins(0, 0, 0, 0)
            widget.setLayout(layout)
            
            self.attachment_list.setItemWidget(item, widget)

        # 添付ファイル数に基づいてリストの高さを調整
        item_height = self.attachment_list.sizeHintForRow(0)  # アイテムの高さを取得
        max_visible_items = 5  # 最大で表示したいアイテム数を設定
        visible_items = min(len(self.attachments), max_visible_items)

        # 添付ファイルがない場合はリストを非表示にし、ファイルがある場合はリストを表示する
        if visible_items == 0:
            self.attachment_list.setFixedHeight(0)  # 添付ファイルがない場合は高さ0
        else:
            new_height = visible_items * item_height + 20 * self.attachment_list.frameWidth()
            self.attachment_list.setFixedHeight(new_height)

        
    def remove_attachment(self, file):
        # 添付ファイルをリストから削除
        if file in self.attachments:
            self.attachments.remove(file)
            self.update_attachment_list()  # リストを更新して高さを再調整
    
    def set_signature_height(self, lines):
        # 現在のフォントサイズを取得
        font = self.signEdit.font()
        font_size = font.pointSize()

        # 一行の高さを推定 (フォントサイズに基づく)
        line_height = font_size * 1.5  # 1.5 は行間の一般的な倍率
        
        # テキストエリアの高さを設定
        self.signEdit.setFixedHeight(int(line_height * lines))

    def apply_Left_alignment(self):
        cursor = self.textEdit.textCursor()
        if cursor.hasSelection():
            # 選択範囲の始点と終点を保存
            start = cursor.selectionStart()
            end = cursor.selectionEnd()

            # 選択範囲のテキストを繰り返して処理
            cursor.setPosition(start)  # カーソルを選択範囲の先頭に移動
            while cursor.position() < end:
                cursor.select(QTextCursor.BlockUnderCursor)  # 現在のブロック（段落）を選択
                block_format = cursor.blockFormat()  # 現在のブロックのフォーマットを取得
                block_format.setAlignment(Qt.AlignLeft)  # 中央寄せに設定
                cursor.mergeBlockFormat(block_format)  # フォーマットをブロックに適用

                cursor.movePosition(QTextCursor.NextBlock)  # 次のブロックに移動
    
        # フォーカスが外れている場合でも適用されるようにテキストエディタの状態を更新
        self.textEdit.setTextCursor(cursor)
    
    def apply_Center_alignment(self):
        cursor = self.textEdit.textCursor()
        if cursor.hasSelection():
            # 選択範囲の始点と終点を保存
            start = cursor.selectionStart()
            end = cursor.selectionEnd()

            # 選択範囲のテキストを繰り返して処理
            cursor.setPosition(start)  # カーソルを選択範囲の先頭に移動
            while cursor.position() < end:
                cursor.select(QTextCursor.BlockUnderCursor)  # 現在のブロック（段落）を選択
                block_format = cursor.blockFormat()  # 現在のブロックのフォーマットを取得
                block_format.setAlignment(Qt.AlignCenter)  # 中央寄せに設定
                cursor.mergeBlockFormat(block_format)  # フォーマットをブロックに適用

                cursor.movePosition(QTextCursor.NextBlock)  # 次のブロックに移動
    
        # フォーカスが外れている場合でも適用されるようにテキストエディタの状態を更新
        self.textEdit.setTextCursor(cursor)

    def apply_Right_alignment(self):
        cursor = self.textEdit.textCursor()
        if cursor.hasSelection():
            # 選択範囲の始点と終点を保存
            start = cursor.selectionStart()
            end = cursor.selectionEnd()

            # 選択範囲のテキストを繰り返して処理
            cursor.setPosition(start)  # カーソルを選択範囲の先頭に移動
            while cursor.position() < end:
                cursor.select(QTextCursor.BlockUnderCursor)  # 現在のブロック（段落）を選択
                block_format = cursor.blockFormat()  # 現在のブロックのフォーマットを取得
                block_format.setAlignment(Qt.AlignRight)  # 中央寄せに設定
                cursor.mergeBlockFormat(block_format)  # フォーマットをブロックに適用

                cursor.movePosition(QTextCursor.NextBlock)  # 次のブロックに移動
    
        # フォーカスが外れている場合でも適用されるようにテキストエディタの状態を更新
        self.textEdit.setTextCursor(cursor)

    def load_csv(self):
        global file_name,headers,recipients_data
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, 'CSVファイルを選択', desktop_path, 'CSVファイル (*.csv)')

        if file_name:
            try:
                df = pd.read_csv(file_name)

                # データフレームが空か、もしくはヘッダー行のみの確認
                if df.empty or len(df) == 0:
                    QMessageBox.warning(self, 'エラー', 'データが存在しません。空のCSVファイルです。')
                    return

                if 'Email' in df.columns and 'Status' in df.columns:
                    # 'Status' 列を文字列型に変換
                    df['Status'] = df['Status'].astype(str)

                    # メールアドレスとして無効な形式を持つ行に対して '作成失敗' を記入
                    for index, row in df.iterrows():
                        email = str(row['Email']).strip()
                        
                        # 条件1: 'http'を含む場合
                        if 'http' in email:
                            df.at[index, 'Status'] = 'アドレスを確認してください'
                        
                        # 条件2: 半角空白を含む場合
                        elif ' ' in email:
                            df.at[index, 'Status'] = 'アドレスを確認してください'
                        
                        # 条件3: 全角空白を含む場合
                        elif '　' in email:
                            df.at[index, 'Status'] = 'アドレスを確認してください'
                        
                        # 条件4: '@'が2つ以上含まれる場合
                        elif email.count('@') != 1:
                            df.at[index, 'Status'] = 'アドレスを確認してください'
                        
                        # 条件5: メールアドレスとして無効な形式（簡易的な正規表現でチェック）
                        elif not re.match(r"^[\w\.-]+@[\w\.-]+\.\w+$", email):
                            df.at[index, 'Status'] = 'アドレスを確認してください'

                    recipients_data = df.to_dict('records')  # 全行データをリストとして保持
                    self.csv_file_path = file_name  # CSVファイルのパスを保存
                    QMessageBox.information(self, 'CSV読み込み完了', f"{len(recipients_data)}件の宛先を読み込みました。")

                    # CSVデータを表示する
                    headers = df.columns.tolist()
                    data = df.values.tolist()

                    # CSVViewerをモードレスで表示
                    if self.csv_viewer is None or not self.csv_viewer.isVisible():
                        self.csv_viewer = CSVViewer(data, headers, parent=self)
                    self.csv_viewer.show()  # モードレスで表示する

                    # 再確認ボタンを有効化
                    self.reopen_button.setEnabled(True)

                    # ヘッダービューウィンドウを有効化
                    self.header_button.setEnabled(True)
                else:
                    QMessageBox.warning(self, 'エラー', 'CSVファイルに "Email"列または"Status"列が見つかりません。')
            except Exception as e:
                QMessageBox.critical(self, 'エラー', f'CSVの読み込みに失敗しました: {e}')

    def create_template(self):
        # ファイルダイアログを開いて保存場所を選択
        file_path = dir_name + '\\CSV_template.csv'

        if file_path:
            # ヘッダー行を定義
            header = ['Email','Status','Last_Name','First_Name','Year', 'Month', 'Dates','URL1']

            # CSVファイルを作成してヘッダーを記入
            with open(file_path, mode='w', newline='', encoding='utf-8-sig') as file:
                writer = csv.writer(file)
                writer.writerow(header)
            
            # 完了メッセージを表示
            QMessageBox.information(self, '作成完了', 'CSV_template.csvを作成しました！')

    def reopen_csv_viewer(self):
        global recipients_data
        #if self.csv_viewer:
            #self.csv_viewer.show()  # 再度表示する

        df = pd.read_csv(file_name)

        if 'Email' in df.columns:
            recipients_data = []  # データリストをクリア
            self.csv_viewer.clear()  # もしCSVViewerにクリアメソッドがなければ、新しいウィジェットを作成する

            recipients_data = df.to_dict('records')  # 全行データをリストとして保持
            headers = df.columns.tolist()
            data = df.values.tolist()

            # CSVViewerをモードレスで再表示
            if self.csv_viewer is None or not self.csv_viewer.isVisible():
                self.csv_viewer = CSVViewer(data, headers, parent=self)

            self.csv_viewer.show()  # モードレスで表示する
    
    def show_header_window(self):
        global headers
        if recipients_data:
            self.header_window = HeaderWindow(headers, parent=self)
            self.header_window.show()

    def insert_header_into_textbox(self, header_name):
        cursor = self.textEdit.textCursor()  # self.textbox にアクセス
        cursor.insertText(f"{{{header_name}}}")  # 変数名を本文に挿入

                
    def send_email(self):
        html_template = self.textEdit.toHtml()  # HTMLテンプレート
        sign_text =  self.signEdit.toHtml() # 署名文章
        html_template = html_template + '\n' + sign_text # 署名文章を追加

        self.sender_email = self.senderLineEdit.text().strip()
        self.Mail_Title = self.MailTitleEdit.text().strip()
        
        if not recipients_data:
            QMessageBox.warning(self, 'エラー', '宛先がありません。CSVを読み込んでください。')
            return
        
        elif not self.Mail_Title:
            QMessageBox.warning(self, 'エラー', 'メールのタイトルを指定してください。')
            return
        
        # 実行直前の最終確認
        reply = QMessageBox.question(self, '確認', 
                                    'メール作成を実行しますか？', 
                                    QMessageBox.Yes | QMessageBox.No, 
                                    QMessageBox.No)
        if reply == QMessageBox.No:
            QMessageBox.information(self, '中止', 'メール作成を中止しました')
            return

        outlook = win32.Dispatch('outlook.application')
        namespace = outlook.GetNamespace("MAPI") # MAPIを取得し、自分のアカウント情報を取得
        updated_data = []  # 更新されたデータを保持するリスト

        for index, data in enumerate(recipients_data):
            try:
                # Status列をチェックして「送信済み」ならスキップ
                if data.get('Status') == '送信済み' or data.get('Status') == '下書き作成済み' or data.get('Status') == 'アドレスを確認してください':
                    print(f"{data['Email']} は既に処理実施済みのためスキップされました。")
                    updated_data.append(data)  # 送信済みのデータをそのままリストに追加
                    continue
            except:
                pass

            # CSVのヘッダーに基づいて変数を動的に生成する
            variables = {}
            for key, value in data.items():
                variables[key] = value

            # 固定変数の追加 (Sign や曜日のDotwなど)
            recipient_email = data['Email']
            variables['Sign'] = 'さんぎょうい株式会社'
            cc_email = variables.get('CC', None)
            bcc_email = variables.get('BCC', None)
            
            # 曜日を計算する場合（Year, Month, Datesがある場合）
            year = variables.get('Year', '1990')  # Yearの値が存在しなければデフォルトで1990
            month = variables.get('Month', '1')   # Monthの値を取得
            day = variables.get('Dates', '1')     # Datesの値を取得
            try:
                dotw = self.calculate_dotw(year, month, day)  # 曜日を計算
                variables['Dotw'] = dotw
            except Exception as e:
                variables['Dotw'] = '不明'
                print(f"曜日の計算に失敗しました: {e}")

            # テンプレートに変数を挿入
            personalized_title = self.insert_variables(self.Mail_Title, variables)
            personalized_html = self.insert_variables(html_template, variables)
            
            mail = outlook.CreateItem(0)
            mail.Subject = personalized_title
            mail.HTMLBody = personalized_html
            mail.To = recipient_email
            if self.sender_email:
                mail.SentOnBehalfOfName = self.sender_email  # 指定された送信者                
            if isinstance(cc_email, str) and cc_email.lower() != "nan":
                mail.CC = cc_email  # CCにアドレスを指定
            if isinstance(bcc_email, str) and bcc_email.lower() != "nan":
                mail.BCC = bcc_email  # BCCにアドレスを指定

            # 画像を本文に埋め込み（添付リストではなく）
            if hasattr(self, 'embedded_images'):
                for content_id, image_path in self.embedded_images.items():
                    try:
                        # パスのバックスラッシュをスラッシュに変換
                        image_path_for_html = image_path.replace('\\', '/')

                        # HTML本文の画像参照をCIDに置き換える
                        personalized_html = personalized_html.replace(f'file:///{image_path_for_html}', f'cid:{os.path.basename(image_path)}')
                        mail.HTMLBody = personalized_html  # 更新されたHTML本文を設定

                        # 埋め込み画像を添付し、ContentIdを設定
                        attachment = mail.Attachments.Add(image_path)
                        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", os.path.basename(image_path))
                    except Exception as e:
                        QMessageBox.critical(self, '画像埋め込みエラー', f'{image_path} の埋め込みに失敗しました: {e}')

            # 通常の添付ファイルを追加
            for index in range(self.attachment_list.count()):
                attachment_path = self.attachment_list.item(index).text()
                try:
                    mail.Attachments.Add(attachment_path)
                except Exception as e:
                    QMessageBox.critical(self, '添付エラー', f'{attachment_path} の添付に失敗しました:\n {e}')

            try:
                if self.draft_checkbox.isChecked():
                    # 下書きとして保存
                    mail.Save()
                    data['Status'] = '下書き作成済み'
                    updated_data.append(data)  # 更新されたデータをリストに追加
                else:
                    # メールを送信
                    time.sleep(2) # 送信時に1秒間待機する
                    mail.Send()

                    # メール送信成功時に「送信済み」と記録
                    data['Status'] = '送信済み'
                    updated_data.append(data)  # 更新されたデータをリストに追加

            except Exception as e:
                data['Status'] = 'アドレスを確認してください'
                QMessageBox.critical(self, '送信エラー', f'{recipient_email} への送信に失敗しました:\n {e}')
                print(f"{recipient_email} への送信に失敗しました:\n {e}")
                updated_data.append(data)  # エラーが発生してもデータをリストに追加

        # 送信後にCSVファイルを更新（送信済みのステータスを反映）
        self.update_csv(updated_data)

        QMessageBox.information(self, '完了', 'メールの処理が完了しました。')

    def send_testemail(self):
        # テストメールを送る際の処理
        html_template = self.textEdit.toHtml()  # HTMLテンプレート
        sign_text =  self.signEdit.toHtml() # 署名文章
        html_template = html_template + '\n' + sign_text # 署名文章を追加

        self.sender_email = self.senderLineEdit.text().strip()
        self.Mail_Title = self.MailTitleEdit.text().strip()

        if not self.Mail_Title:
            QMessageBox.warning(self, 'エラー', 'メールのタイトルを指定してください。')
            return

        outlook = win32.Dispatch('outlook.application')
        namespace = outlook.GetNamespace("MAPI") # MAPIを取得し、自分のアカウント情報を取得
        account = namespace.Accounts.Item(1)     # デフォルトアカウントを取得

        updated_data = []  # 更新されたデータを保持するリスト

        if recipients_data:
            # CSVのヘッダーと1行目のデータを取得
            headers = recipients_data[0]
        
            # CSVのヘッダーに基づいて変数を動的に生成する
            variables = {}
            for key, value in zip(headers.keys(), headers.values()):
                variables[key] = value  # ヘッダーに基づいて値を設定
        
            # 曜日を計算する場合（Year, Month, Datesがある場合）
            year = variables.get('Year', '1990')  # Yearの値が存在しなければデフォルトで1990
            month = variables.get('Month', '1')   # Monthの値を取得
            day = variables.get('Dates', '1')     # Datesの値を取得
            try:
                dotw = self.calculate_dotw(year, month, day)  # 曜日を計算
                variables['Dotw'] = dotw
            except Exception as e:
                variables['Dotw'] = '不明'
                print(f"曜日の計算に失敗しました: {e}")

            # テンプレートに変数を挿入
            personalized_title = self.insert_variables(self.Mail_Title, variables)
            personalized_html = self.insert_variables(html_template, variables)
        else:
            personalized_title = self.Mail_Title
            personalized_html = html_template

        mail = outlook.CreateItem(0)
        mail.Subject = personalized_title
        mail.HTMLBody = personalized_html
        mail.To = account

        # 送信先をテスト用のアドレスに設定（例: 自分のアドレス）
        if self.sender_email:
            mail.SentOnBehalfOfName = self.sender_email  # 指定された送信者

        # 画像を本文に埋め込み（添付リストではなく）
        if hasattr(self, 'embedded_images'):
            for content_id, image_path in self.embedded_images.items():
                try:
                    # パスのバックスラッシュをスラッシュに変換
                    image_path_for_html = image_path.replace('\\', '/')

                    # HTML本文の画像参照をCIDに置き換える
                    personalized_html = personalized_html.replace(f'file:///{image_path_for_html}', f'cid:{os.path.basename(image_path)}')
                    mail.HTMLBody = personalized_html  # 更新されたHTML本文を設定

                    # 埋め込み画像を添付し、ContentIdを設定
                    attachment = mail.Attachments.Add(image_path)
                    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", os.path.basename(image_path))
                except Exception as e:
                    QMessageBox.critical(self, '画像埋め込みエラー', f'{image_path} の埋め込みに失敗しました: {e}')


        # 通常の添付ファイルを追加
        for index in range(self.attachment_list.count()):
            attachment_path = self.attachment_list.item(index).text()
            try:
                mail.Attachments.Add(attachment_path)
            except Exception as e:
                QMessageBox.critical(self, '添付エラー', f'{attachment_path} の添付に失敗しました: {e}')

        try:
            if self.draft_checkbox.isChecked():
                # 下書きとして保存
                mail.Save()
            else:
                # メールを送信
                mail.Send()

        except Exception as e:
            QMessageBox.critical(self, '送信エラー', f'テストメールの送信に失敗しました: {e}')

        QMessageBox.information(self, '完了', 'テストメールの処理が完了しました。')

    def insert_variables(self,template, variables):
        """
        テンプレート内の変数（{}で囲まれた文字列）を対応する値に置換する関数。
        
        :param template: 変数を含むテンプレート文字列
        :param variables: 変数名と値の辞書
        :return: 変数が置換されたテンプレート文字列
        """
        for key, value in variables.items():
            # {key} の形式でテンプレート内を置換
            template = template.replace(f"{{{key}}}", str(value))
        return template
    
    def calculate_dotw(self,year, month, day):
        try:
            # 日付を作成
            date = datetime(int(year), int(month), int(day))
            # 曜日を取得 (例: "月曜日", "火曜日"など)
            dotw = date.strftime('%A')  # 英語の曜日
            dotw_jp = {
                'Monday': '月',
                'Tuesday': '火',
                'Wednesday': '水',
                'Thursday': '木',
                'Friday': '金',
                'Saturday': '土',
                'Sunday': '日'
            }
            return dotw_jp[dotw]  # 日本語の曜日に変換して返す
        except ValueError:
            return ''  # 月や日の値がそもそもない場合は空欄にする   
    
    def update_csv(self, updated_data):
        # 送信済みステータスを反映してCSVファイルを更新する関数
        if hasattr(self, 'csv_file_path'):
            try:
                # updated_data を DataFrame に変換
                df = pd.DataFrame(updated_data)
                
                # CSVファイルに書き込む（エンコーディングを指定）
                df.to_csv(self.csv_file_path, index=False, encoding='utf-8-sig')
                
                # 更新完了メッセージを表示
                QMessageBox.information(self, 'CSV更新', 'CSVファイルを更新しました。')

                # CSVViewer が表示されている場合はデータを更新する
                if self.csv_viewer and self.csv_viewer.isVisible():
                    headers = df.columns.tolist()
                    data = df.values.tolist()
                    self.csv_viewer.update_data(data)
            except Exception as e:
                # 更新失敗メッセージを表示
                QMessageBox.critical(self, 'CSV更新エラー', f'CSVファイルの更新に失敗しました: {e}')
        else:
            QMessageBox.warning(self, 'エラー', 'CSVファイルのパスが設定されていません。')

    def load_templates_from_file(self, file_path):
        # 外部ファイルからテンプレートを読み込む処理
        if not os.path.exists(file_path):
            QMessageBox.warning(self, 'エラー', f'テンプレートファイル "{file_path}" が見つかりません。')
            return

        with open(file_path, 'r', encoding='utf-8-sig') as file:
            lines = file.readlines()
            
        current_template_name = None
        current_template_content = []
        current_template_title = ""
        sign_content = []  # 署名用のコンテンツリスト
        is_sign_section = False  # 署名セクションのフラグ

        for line in lines:
            line = line.strip()
            
            if line == '=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=':
                continue  # 分け棒の行はスキップ

            if line.startswith('【テンプレート】：'):  # 新しいテンプレート名の検出
                # 既にテンプレートが存在する場合、保存
                if current_template_name:
                    self.templates[current_template_name] = {
                        'title': current_template_title,
                        'content': "\n".join(current_template_content)
                    }

                # 新しいテンプレート名を取得し、【テンプレート】：を除去
                current_template_name = line.replace('【テンプレート】：', '').strip()
                current_template_content = []  # コンテンツをリセット
                current_template_title = ""  # タイトルもリセット
                sign_content = []  # 署名もリセット
                is_sign_section = False  # 署名セクションもリセット

            elif line.startswith('【タイトル】：'):  # タイトルの検出
                current_template_title = line.replace('【タイトル】：', '').strip()
                
            elif line.startswith('【署名】'):  # 署名セクションの検出
                is_sign_section = True  # 署名セクションに切り替え 
            
            elif is_sign_section:
                # 署名セクションが開始されている場合は署名として保存
                sign_content.append(line)

            else:
                current_template_content.append(line)  # コンテンツを追加

        # 最後のテンプレートを保存（ファイルの最後にテンプレートがある場合）
        if current_template_name:
            self.templates[current_template_name] = {
                'title': current_template_title,
                'content': "\n".join(current_template_content)
            }
        
        # 署名セクションがあった場合、self.signEdit に署名をセット
        if sign_content:
            self.signEdit.setPlainText("\n".join(sign_content))

    def apply_template(self):
        selected_template_name = self.templateComboBox.currentText()
        if selected_template_name in self.templates:
            template = self.templates[selected_template_name]
            self.MailTitleEdit.setText(template['title'])  # メールタイトルに設定
            self.textEdit.setPlainText(template['content'])  # メール本文に設定
    
    def apply_variables_to_template(self, template, variables):
        # 曜日を計算する場合（Year, Month, Datesがある場合）
        year = variables.get('{Year}', '1990')  # Yearの値が存在しなければデフォルトで1990
        month = variables.get('{Month}', '1')   # Monthの値を取得
        day = variables.get('{Dates}', '1')     # Datesの値を取得
        dotw = self.calculate_dotw(year, month, day)  # 曜日を計算
        
        # 変数をテンプレートに反映させる処理
        for key, value in variables.items():
            # プレースホルダーを対応する値で置き換え
            template = template.replace(key, str(value))
        
        # 曜日だけは特殊計算なので
        template = template.replace("{Dotw}", dotw)

        return template

if __name__ == '__main__':
    app = QApplication(sys.argv)
    # アプリケーションのアイコンを設定
    exe_icon_path = os.path.join(cwd, "image\\exe_logo.png")
    app.setWindowIcon(QIcon(exe_icon_path))
    window = EmailSenderApp()
    window.show()
    sys.exit(app.exec_())
