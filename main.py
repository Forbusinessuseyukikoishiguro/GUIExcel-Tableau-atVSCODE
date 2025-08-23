"""
Tableau Excel データ前処理ツール GUI版（データ出力修正版）
Author: Claude
Description: ExcelデータをTableau分析用に効率的に前処理するためのGUIツール
"""

import pandas as pd
import numpy as np
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime
from typing import List, Dict, Optional, Union, Any
import warnings
warnings.filterwarnings('ignore')


class TableauPreprocessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Tableau Excel前処理ツール")
        self.root.geometry("1400x800")
        self.root.configure(bg='#f0f0f0')
        
        # データ関連変数
        self.data = None
        self.original_data = None
        self.data_info = {}
        self.processing_log = []
        self.file_path = None
        
        # スタイル設定
        self.setup_styles()
        
        # GUI構築
        self.create_widgets()
        
        # 初期状態設定（エクスポートボタンを有効化）
        self.update_button_states()
        self.force_enable_all_export_buttons_startup()
    
    def setup_styles(self):
        """スタイルを設定"""
        self.style = ttk.Style()
        
        # テーマ設定
        available_themes = self.style.theme_names()
        if 'clam' in available_themes:
            self.style.theme_use('clam')
        
        # カスタムスタイル
        self.style.configure('Title.TLabel', font=('Arial', 16, 'bold'))
        self.style.configure('Heading.TLabel', font=('Arial', 12, 'bold'))
        self.style.configure('Action.TButton', font=('Arial', 10, 'bold'))
        self.style.configure('Export.TButton', font=('Arial', 11, 'bold'), foreground='white')
    
    def create_widgets(self):
        """GUIウィジェットを作成"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # グリッド設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # ヘッダー
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        title_label = ttk.Label(header_frame, text="📊 Tableau Excel前処理ツール", style='Title.TLabel')
        title_label.pack()
        
        subtitle_label = ttk.Label(header_frame, text="ExcelデータをTableau分析用に効率的にクリーニング・変換")
        subtitle_label.pack()
        
        # 左側パネル（コントロール）
        self.create_control_panel(main_frame)
        
        # 右側パネル（データ表示）
        self.create_data_panel(main_frame)
    
    def create_control_panel(self, parent):
        """左側のコントロールパネルを作成"""
        control_frame = ttk.Frame(parent, width=380)
        control_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        control_frame.grid_propagate(False)
        
        # ファイル選択セクション
        file_section = ttk.LabelFrame(control_frame, text="📁 ファイル操作", padding="10")
        file_section.pack(fill=tk.X, pady=(0, 10))
        
        self.file_label = ttk.Label(file_section, text="ファイルが選択されていません", foreground='gray')
        self.file_label.pack(fill=tk.X, pady=(0, 5))
        
        file_button_frame = ttk.Frame(file_section)
        file_button_frame.pack(fill=tk.X)
        
        self.load_button = ttk.Button(file_button_frame, text="📂 ファイル選択", 
                                     command=self.load_file, style='Action.TButton')
        self.load_button.pack(side=tk.LEFT, padx=(0, 5))
        
        self.reload_button = ttk.Button(file_button_frame, text="🔄 再読み込み", 
                                       command=self.reload_file, state='disabled')
        self.reload_button.pack(side=tk.LEFT)
        
        # データ情報セクション
        info_section = ttk.LabelFrame(control_frame, text="📊 データ概要", padding="10")
        info_section.pack(fill=tk.X, pady=(0, 10))
        
        self.info_text = tk.Text(info_section, height=8, font=('Consolas', 9), 
                                state='disabled', bg='#f8f8f8')
        self.info_text.pack(fill=tk.BOTH)
        
        # クリーニングセクション
        cleaning_section = ttk.LabelFrame(control_frame, text="🧹 データクリーニング", padding="10")
        cleaning_section.pack(fill=tk.X, pady=(0, 10))
        
        self.empty_rows_button = ttk.Button(cleaning_section, text="空行削除", 
                                           command=self.remove_empty_rows, state='disabled')
        self.empty_rows_button.pack(fill=tk.X, pady=2)
        
        self.empty_cols_button = ttk.Button(cleaning_section, text="空列削除", 
                                           command=self.remove_empty_columns, state='disabled')
        self.empty_cols_button.pack(fill=tk.X, pady=2)
        
        self.duplicates_button = ttk.Button(cleaning_section, text="重複行削除", 
                                           command=self.remove_duplicates, state='disabled')
        self.duplicates_button.pack(fill=tk.X, pady=2)
        
        self.missing_button = ttk.Button(cleaning_section, text="欠損値処理", 
                                        command=self.fill_missing_dialog, state='disabled')
        self.missing_button.pack(fill=tk.X, pady=2)
        
        # 変換セクション
        transform_section = ttk.LabelFrame(control_frame, text="🔄 データ変換", padding="10")
        transform_section.pack(fill=tk.X, pady=(0, 10))
        
        self.text_clean_button = ttk.Button(transform_section, text="テキスト整形", 
                                           command=self.clean_text_data, state='disabled')
        self.text_clean_button.pack(fill=tk.X, pady=2)
        
        self.type_convert_button = ttk.Button(transform_section, text="型変換", 
                                             command=self.convert_data_types, state='disabled')
        self.type_convert_button.pack(fill=tk.X, pady=2)
        
        self.rename_button = ttk.Button(transform_section, text="列名変更", 
                                       command=self.rename_columns_dialog, state='disabled')
        self.rename_button.pack(fill=tk.X, pady=2)
        
        self.filter_button = ttk.Button(transform_section, text="データフィルタ", 
                                       command=self.filter_data_dialog, state='disabled')
        self.filter_button.pack(fill=tk.X, pady=2)
        
        # エクスポートセクション（修正版・常に表示）
        export_section = ttk.LabelFrame(control_frame, text="💾 データ出力", padding="15")
        export_section.pack(fill=tk.X, pady=(0, 10))
        
        # ===== メインエクスポートボタン（大きく目立たせる・常に有効）=====
        main_export_frame = ttk.Frame(export_section)
        main_export_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.main_export_button = tk.Button(main_export_frame, 
                                           text="📄 データを保存する", 
                                           command=self.safe_export_data,
                                           font=('Arial', 12, 'bold'),
                                           bg='#4CAF50',
                                           fg='white',
                                           activebackground='#45a049',
                                           relief='raised',
                                           bd=3,
                                           height=2,
                                           state='normal')  # 常に有効
        self.main_export_button.pack(fill=tk.X, pady=5)
        
        # 個別保存ボタン
        individual_frame = ttk.LabelFrame(export_section, text="個別保存", padding="10")
        individual_frame.pack(fill=tk.X, pady=(0, 10))
        
        export_buttons_frame = ttk.Frame(individual_frame)
        export_buttons_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.export_excel_button = tk.Button(export_buttons_frame, text="📊 Excel保存", 
                                             command=self.quick_export_excel, 
                                             state='normal',  # 常に有効
                                             font=('Arial', 10),
                                             bg='#2196F3',
                                             fg='white',
                                             width=12,
                                             relief='raised',
                                             bd=2)
        self.export_excel_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        self.export_csv_button = tk.Button(export_buttons_frame, text="📄 CSV保存", 
                                           command=self.quick_export_csv, 
                                           state='normal',  # 常に有効
                                           font=('Arial', 10),
                                           bg='#FF9800',
                                           fg='white',
                                           width=12,
                                           relief='raised',
                                           bd=2)
        self.export_csv_button.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 強制有効化ボタン（デバッグ用）
        debug_frame = ttk.Frame(individual_frame)
        debug_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.force_enable_button = ttk.Button(debug_frame, text="🔧 ボタン有効化", 
                                             command=self.force_enable_export,
                                             style='Action.TButton')
        self.force_enable_button.pack(side=tk.LEFT, padx=(0, 5))
        
        self.test_button = ttk.Button(debug_frame, text="🔍 状態確認", 
                                     command=self.test_data_status)
        self.test_button.pack(side=tk.LEFT)
        
        # エクスポート設定
        settings_frame = ttk.LabelFrame(export_section, text="保存設定", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 保存形式選択
        format_frame = ttk.Frame(settings_frame)
        format_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(format_frame, text="形式:", font=('Arial', 9, 'bold')).pack(anchor=tk.W)
        self.export_format = tk.StringVar(value="excel")
        
        format_radio_frame = ttk.Frame(settings_frame)
        format_radio_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Radiobutton(format_radio_frame, text="Excel (.xlsx)", 
                       variable=self.export_format, value="excel").pack(anchor=tk.W, padx=10)
        ttk.Radiobutton(format_radio_frame, text="CSV (.csv)", 
                       variable=self.export_format, value="csv").pack(anchor=tk.W, padx=10)
        
        # 追加オプション
        self.include_summary = tk.BooleanVar(value=True)
        self.include_log = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(settings_frame, text="処理サマリーを含める", 
                       variable=self.include_summary).pack(anchor=tk.W)
        ttk.Checkbutton(settings_frame, text="処理ログを含める", 
                       variable=self.include_log).pack(anchor=tk.W)
        
        # ステータス表示（常にポジティブ）
        self.export_status = ttk.Label(export_section, 
                                      text="✅ エクスポート可能（データ読み込み後推奨）", 
                                      foreground='green', font=('Arial', 9))
        self.export_status.pack(pady=(10, 0))
        
        # その他セクション
        other_section = ttk.LabelFrame(control_frame, text="🔧 その他", padding="10")
        other_section.pack(fill=tk.X, pady=(0, 10))
        
        self.reset_button = ttk.Button(other_section, text="🔄 データリセット", 
                                      command=self.reset_data, state='disabled')
        self.reset_button.pack(fill=tk.X, pady=2)
        
        self.log_button = ttk.Button(other_section, text="📋 処理ログ表示", 
                                    command=self.show_processing_log, state='disabled')
        self.log_button.pack(fill=tk.X, pady=2)
    
    def create_data_panel(self, parent):
        """右側のデータ表示パネルを作成"""
        data_frame = ttk.Frame(parent)
        data_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        data_frame.columnconfigure(0, weight=1)
        data_frame.rowconfigure(1, weight=1)
        
        # データ表示ヘッダー
        data_header = ttk.Label(data_frame, text="データプレビュー", style='Heading.TLabel')
        data_header.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        # データテーブル
        self.create_data_table(data_frame)
        
        # ステータスバー
        status_frame = ttk.Frame(data_frame)
        status_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        status_frame.columnconfigure(0, weight=1)
        
        self.status_label = ttk.Label(status_frame, text="ファイルを選択してください")
        self.status_label.grid(row=0, column=0, sticky=tk.W)
        
        # プログレスバー
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(status_frame, variable=self.progress_var, 
                                          maximum=100, length=200)
        self.progress_bar.grid(row=0, column=1, sticky=tk.E)
    
    def create_data_table(self, parent):
        """データテーブルを作成"""
        table_frame = ttk.Frame(parent)
        table_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        
        # Treeviewウィジェット
        self.tree = ttk.Treeview(table_frame, show='tree headings', height=20)
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # スクロールバー
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        self.tree.configure(xscrollcommand=h_scrollbar.set)
    
    def load_file(self):
        """ファイルを読み込む"""
        file_path = filedialog.askopenfilename(
            title="Excel/CSVファイルを選択",
            filetypes=[
                ("すべて対応形式", "*.xlsx;*.xls;*.csv;*.tsv;*.txt"),
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("TSV files", "*.tsv"),
                ("Text files", "*.txt"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.file_path = file_path
            try:
                self.load_data()
            except Exception as e:
                self.show_detailed_error("ファイル選択エラー", str(e))
    
    def load_data(self):
        """データを読み込む"""
        try:
            self.update_status("ファイルを読み込み中...")
            self.progress_var.set(10)
            self.root.update()
            
            # ファイル存在確認
            if not os.path.exists(self.file_path):
                raise FileNotFoundError(f"ファイルが見つかりません: {self.file_path}")
            
            # ファイル拡張子を取得
            file_extension = os.path.splitext(self.file_path)[1].lower()
            self.progress_var.set(20)
            self.root.update()
            
            print(f"DEBUG: ファイル読み込み開始 - {self.file_path}")
            print(f"DEBUG: 拡張子: {file_extension}")
            
            # Excel形式の読み込み
            if file_extension in ['.xlsx', '.xls']:
                self.data = self.load_excel_file()
            # CSV形式の読み込み  
            elif file_extension == '.csv':
                self.data = self.load_csv_with_multiple_encodings()
            # TSV形式の読み込み
            elif file_extension in ['.tsv', '.txt']:
                self.data = self.load_csv_with_multiple_encodings(separator='\t')
            # 不明な形式（CSV形式として試行）
            else:
                print(f"WARNING: 不明な拡張子 {file_extension}、CSV形式で試行")
                self.data = self.load_csv_with_multiple_encodings()
            
            self.progress_var.set(70)
            self.root.update()
            
            # データの基本検証
            if self.data is None or self.data.empty:
                raise ValueError("読み込まれたデータが空です")
            
            print(f"DEBUG: データ読み込み成功 - {self.data.shape[0]}行 × {self.data.shape[1]}列")
            
            # データの基本整理
            self.clean_loaded_data()
            
            # 元データのバックアップ
            self.original_data = self.data.copy()
            
            # GUI更新
            self.on_data_loaded()
            
        except Exception as e:
            print(f"ERROR: データ読み込み失敗 - {str(e)}")
            self.show_detailed_error("ファイル読み込みエラー", str(e))
            self.progress_var.set(0)
    
    def load_excel_file(self):
        """Excel形式ファイルを読み込む"""
        engines = ['openpyxl', 'xlrd']
        
        for engine in engines:
            try:
                print(f"DEBUG: Excel読み込み試行 - engine: {engine}")
                
                # 基本的なパラメータで読み込み
                data = pd.read_excel(
                    self.file_path, 
                    engine=engine,
                    na_values=['', 'NULL', 'null', 'None', 'none', 'NaN', 'nan', '#N/A', '#VALUE!'],
                    keep_default_na=True
                )
                
                print(f"SUCCESS: Excel読み込み成功 - engine: {engine}")
                return data
                
            except ImportError:
                print(f"WARNING: エンジン {engine} が利用できません")
                continue
            except Exception as e:
                print(f"WARNING: エンジン {engine} で失敗: {str(e)}")
                continue
        
        # 全てのエンジンで失敗した場合、pandasのデフォルトで試行
        try:
            print("DEBUG: デフォルトエンジンで Excel読み込み試行")
            data = pd.read_excel(self.file_path)
            print("SUCCESS: デフォルトエンジンで Excel読み込み成功")
            return data
        except Exception as e:
            print(f"ERROR: 全てのExcelエンジンで失敗: {str(e)}")
            raise Exception(f"Excelファイルを読み込めませんでした。\n{str(e)}\n\nCSV形式で保存し直して試してください。")
    
    def load_csv_with_multiple_encodings(self, separator=','):
        """複数エンコーディングでCSV読み込みを試行"""
        encodings = [
            'utf-8-sig',  # BOM付きUTF-8
            'utf-8',      # UTF-8
            'shift_jis',  # Shift-JIS (Windows日本語)
            'cp932',      # Windows日本語
            'iso-2022-jp', # JIS
            'euc-jp',     # EUC-JP
            'latin1',     # Latin-1
            'cp1252',     # Windows Western
        ]
        
        separators = [separator, ',', '\t', ';', '|'] if separator == ',' else [separator]
        
        for encoding in encodings:
            for sep in separators:
                try:
                    print(f"TRY: CSV読み込み (encoding: {encoding}, separator: '{sep}')")
                    
                    # パラメータ設定
                    csv_params = {
                        'encoding': encoding,
                        'sep': sep,
                        'na_values': ['', 'NULL', 'null', 'None', 'none', 'NaN', 'nan', '#N/A', '#VALUE!'],
                        'keep_default_na': True,
                        'skipinitialspace': True,
                        'skip_blank_lines': True,
                    }
                    
                    # 最初の数行で区切り文字を自動判定
                    if sep == ',' and separator == ',':
                        # 自動判定を試行
                        with open(self.file_path, 'r', encoding=encoding) as f:
                            sample = f.read(1024)
                            if sample.count('\t') > sample.count(','):
                                csv_params['sep'] = '\t'
                            elif sample.count(';') > sample.count(','):
                                csv_params['sep'] = ';'
                    
                    data = pd.read_csv(self.file_path, **csv_params)
                    
                    # データ品質チェック
                    if not data.empty and len(data.columns) > 1:
                        print(f"SUCCESS: CSV読み込み成功 (encoding: {encoding}, sep: '{csv_params['sep']}')")
                        return data
                        
                except UnicodeDecodeError:
                    continue
                except Exception as e:
                    print(f"WARNING: {encoding}/{sep} で失敗: {e}")
                    continue
        
        raise Exception(f"すべてのエンコーディング・区切り文字で読み込みに失敗しました")
    
    def clean_loaded_data(self):
        """読み込まれたデータの基本整理"""
        try:
            # 列名の重複チェック・修正
            if self.data.columns.duplicated().any():
                print("INFO: 重複列名を修正")
                new_columns = []
                seen = {}
                for col in self.data.columns:
                    if col in seen:
                        seen[col] += 1
                        new_columns.append(f"{col}_{seen[col]}")
                    else:
                        seen[col] = 0
                        new_columns.append(col)
                self.data.columns = new_columns
            
            # 完全に空の行・列を削除
            initial_shape = self.data.shape
            self.data = self.data.dropna(how='all', axis=0)  # 空行削除
            self.data = self.data.dropna(how='all', axis=1)  # 空列削除
            
            if self.data.shape != initial_shape:
                print(f"INFO: 空行・列を削除 {initial_shape} → {self.data.shape}")
            
            # インデックスをリセット
            self.data = self.data.reset_index(drop=True)
            
        except Exception as e:
            print(f"WARNING: データ整理中にエラー: {str(e)}")
            # エラーが発生しても続行
    
    def show_detailed_error(self, title, error_message):
        """詳細なエラーメッセージを表示"""
        detailed_msg = f"❌ {title}\n\n"
        detailed_msg += f"エラー詳細:\n{error_message}\n\n"
        detailed_msg += f"解決策:\n"
        detailed_msg += f"• ファイルが他のアプリケーションで開かれていないか確認\n"
        detailed_msg += f"• ファイルが破損していないか確認\n"
        detailed_msg += f"• 別の形式（Excel→CSV）で保存して試行\n"
        detailed_msg += f"• ファイルパスに日本語や特殊文字が含まれていないか確認\n"
        detailed_msg += f"• 必要なライブラリをインストール: pip install openpyxl xlrd\n\n"
        detailed_msg += f"対応形式: .xlsx, .xls, .csv, .tsv, .txt"
        
        messagebox.showerror(title, detailed_msg)
        self.progress_var.set(0)
    
    def on_data_loaded(self):
        """データ読み込み完了後の処理"""
        self.progress_var.set(100)
        
        # ファイル名表示
        filename = os.path.basename(self.file_path)
        self.file_label.config(text=f"📁 {filename}", foreground='black')
        
        # データ情報更新
        self.update_data_info()
        self.update_data_table()
        
        # ボタンを有効化（確実に実行）
        self.update_button_states(True)
        self.force_enable_all_export_buttons()
        
        self.log_action(f"ファイル読み込み完了: {filename}")
        self.update_status(f"✅ データ読み込み完了 ({self.data.shape[0]}行 × {self.data.shape[1]}列)")
        
        # プログレスバーリセット
        self.root.after(2000, lambda: self.progress_var.set(0))
    
    def force_enable_all_export_buttons_startup(self):
        """プログラム起動時にエクスポートボタンを確実に有効化"""
        try:
            print("DEBUG: 起動時エクスポートボタン有効化")
            
            # メインエクスポートボタン
            if hasattr(self, 'main_export_button'):
                self.main_export_button.config(state='normal', bg='#4CAF50')
            
            # 個別エクスポートボタン
            if hasattr(self, 'export_excel_button'):
                self.export_excel_button.config(state='normal', bg='#2196F3')
            if hasattr(self, 'export_csv_button'):
                self.export_csv_button.config(state='normal', bg='#FF9800')
            
            # ステータス更新
            if hasattr(self, 'export_status'):
                self.export_status.config(text="✅ エクスポート可能（データ読み込み後推奨）", foreground='green')
            
            # GUI更新を強制
            self.root.update()
            self.root.update_idletasks()
            
            print("DEBUG: 起動時エクスポートボタン有効化完了")
        except Exception as e:
            print(f"ERROR: 起動時エクスポートボタン有効化失敗: {e}")

    def force_enable_all_export_buttons(self):
        """エクスポートボタンをすべて強制的に有効化"""
        if self.data is not None:
            print("DEBUG: 全エクスポートボタンを強制有効化")
            
            # メインエクスポートボタン
            self.main_export_button.config(state='normal', bg='#4CAF50')
            
            # 個別エクスポートボタン
            self.export_excel_button.config(state='normal')
            self.export_csv_button.config(state='normal')
            
            # ステータス更新
            self.export_status.config(text="✅ 保存可能", foreground='green')
            
            # GUI更新を強制
            self.root.update()
            self.root.update_idletasks()
    
    def reload_file(self):
        """ファイルを再読み込み"""
        if self.file_path:
            self.load_data()
    
    def update_data_info(self):
        """データ情報を更新"""
        if self.data is None:
            return
        
        self.data_info = {
            'rows': len(self.data),
            'columns': len(self.data.columns),
            'empty_cells': self.data.isnull().sum().sum(),
            'duplicates': self.data.duplicated().sum(),
            'memory_usage': self.data.memory_usage(deep=True).sum()
        }
        
        # 情報テキスト更新
        info_text = f"""行数: {self.data_info['rows']:,}
列数: {self.data_info['columns']}
空セル数: {self.data_info['empty_cells']:,}
重複行数: {self.data_info['duplicates']:,}
メモリ使用量: {self.data_info['memory_usage']/1024/1024:.2f} MB

列情報:"""
        
        for i, (col, dtype) in enumerate(self.data.dtypes.items()):
            null_count = self.data[col].isnull().sum()
            null_rate = (null_count / len(self.data)) * 100
            info_text += f"\n{i+1:2d}. {col[:20]:<20} | {str(dtype):<10}"
            if null_count > 0:
                info_text += f" | 欠損:{null_count}({null_rate:.1f}%)"
        
        self.info_text.config(state='normal')
        self.info_text.delete(1.0, tk.END)
        self.info_text.insert(1.0, info_text)
        self.info_text.config(state='disabled')
    
    def update_data_table(self, max_rows=100):
        """データテーブルを更新"""
        if self.data is None:
            return
        
        # 既存のデータをクリア
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 列設定
        columns = list(self.data.columns)
        self.tree["columns"] = columns
        self.tree["show"] = "headings"
        
        # ヘッダー設定
        for col in columns:
            self.tree.heading(col, text=str(col))
            self.tree.column(col, width=120, minwidth=80)
        
        # データ追加（最大100行まで）
        display_data = self.data.head(max_rows)
        for index, row in display_data.iterrows():
            values = []
            for col in columns:
                value = row[col]
                if pd.isna(value):
                    values.append("")
                elif isinstance(value, float):
                    values.append(f"{value:.2f}" if abs(value) < 1000000 else f"{value:.2e}")
                else:
                    values.append(str(value)[:50])  # 長すぎる値を切り詰め
            self.tree.insert("", "end", values=values)
        
        # 省略表示
        if len(self.data) > max_rows:
            self.tree.insert("", "end", values=[f"... 他 {len(self.data) - max_rows} 行"] + [""] * (len(columns) - 1))
    
    def update_button_states(self, enabled=False):
        """ボタンの有効/無効状態を更新"""
        state = 'normal' if enabled else 'disabled'
        buttons = [
            self.reload_button, self.empty_rows_button, self.empty_cols_button,
            self.duplicates_button, self.missing_button, self.text_clean_button,
            self.type_convert_button, self.rename_button, self.filter_button,
            self.reset_button, self.log_button
        ]
        
        for button in buttons:
            button.config(state=state)
        
        # エクスポートボタンは常に有効に保つ
        self.main_export_button.config(state='normal', bg='#4CAF50')
        self.export_excel_button.config(state='normal', bg='#2196F3')
        self.export_csv_button.config(state='normal', bg='#FF9800')
        
        # ステータス更新
        if self.data is not None:
            self.export_status.config(text="✅ データ読み込み済み - 保存可能", foreground='green')
        else:
            self.export_status.config(text="⚠️ データ未読み込み - 保存は可能だが推奨しない", foreground='orange')
    
    def update_status(self, message):
        """ステータス表示を更新"""
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def log_action(self, action):
        """処理ログを記録"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {action}"
        self.processing_log.append(log_entry)
    
    def show_error(self, message):
        """エラーダイアログを表示"""
        messagebox.showerror("エラー", message)
        self.progress_var.set(0)
    
    def test_data_status(self):
        """データ状態をテスト（デバッグ用）"""
        if self.data is None:
            messagebox.showinfo("データ状態", 
                "❌ データが読み込まれていません\n\n"
                "📂 「ファイル選択」ボタンでExcel/CSVファイルを読み込んでください")
            print("DEBUG: データなし")
        else:
            info = f"✅ データが読み込まれています\n\n"
            info += f"📊 サイズ: {self.data.shape[0]:,}行 × {self.data.shape[1]}列\n"
            info += f"📁 ファイル: {os.path.basename(self.file_path) if self.file_path else 'N/A'}\n"
            info += f"🔧 実行した処理: {len(self.processing_log)}件\n\n"
            info += f"ボタン状態:\n"
            info += f"• メインボタン: {self.main_export_button['state']}\n"
            info += f"• Excel保存: {self.export_excel_button['state']}\n"
            info += f"• CSV保存: {self.export_csv_button['state']}\n\n"
            
            if self.main_export_button['state'] == 'disabled':
                info += f"⚠️ エクスポートボタンが無効です\n"
                info += f"「🔧 ボタン有効化」ボタンを押してください"
            else:
                info += f"✅ エクスポート可能です！"
            
            messagebox.showinfo("データ状態", info)
            print(f"DEBUG: データあり - {self.data.shape}")
    
    def force_enable_export(self):
        """エクスポートボタンを強制的に有効化"""
        if self.data is None:
            messagebox.showwarning("警告", 
                "❌ データが読み込まれていません。\n\n"
                "手順:\n"
                "1. 📂「ファイル選択」でExcel/CSVファイルを読み込み\n"
                "2. このボタンを再度クリック")
            return
        
        # 強制的にボタンを有効化
        self.main_export_button.config(state='normal', bg='#4CAF50')
        self.export_excel_button.config(state='normal')
        self.export_csv_button.config(state='normal')
        self.export_status.config(text="✅ 強制有効化完了", foreground='green')
        
        # GUI更新
        self.root.update()
        self.root.update_idletasks()
        
        # 成功メッセージ
        messagebox.showinfo("有効化完了", 
            "✅ エクスポートボタンを有効化しました！\n\n"
            "📄 メインボタン: データを保存する\n"
            "📊 Excelボタン: 簡単Excel保存\n"
            "📄 CSVボタン: 簡単CSV保存\n\n"
            "すべてのボタンが使用可能になりました。")
        
        print("DEBUG: エクスポートボタンを強制有効化しました")
    
    # ===== エクスポート機能（修正版）=====
    def safe_export_data(self):
        """安全なデータエクスポート"""
        if self.data is None:
            messagebox.showwarning("データ不足", 
                "❌ エクスポートするデータがありません。\n\n"
                "手順:\n"
                "1. 📂「ファイル選択」でExcel/CSVファイルを読み込み\n"
                "2. 必要に応じてデータ処理を実行\n"
                "3. このボタンでデータを保存")
            return
        
        try:
            # 保存先決定
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            export_format = self.export_format.get()
            
            if self.file_path:
                base_dir = os.path.dirname(self.file_path)
                base_name = os.path.splitext(os.path.basename(self.file_path))[0]
                extension = ".xlsx" if export_format == "excel" else ".csv"
                output_path = os.path.join(base_dir, f"{base_name}_processed_{timestamp}{extension}")
            else:
                extension = ".xlsx" if export_format == "excel" else ".csv"
                output_path = f"processed_data_{timestamp}{extension}"
            
            # エクスポート実行
            self.update_status("データを保存中...")
            self.export_status.config(text="💾 保存中...", foreground='orange')
            self.progress_var.set(10)
            
            if export_format == "excel":
                self.safe_export_excel(output_path)
            else:
                self.safe_export_csv(output_path)
            
            # 成功処理
            self.export_status.config(text="✅ 保存完了", foreground='green')
            self.log_action(f"データ保存完了: {os.path.basename(output_path)}")
            self.update_status(f"✅ 保存完了: {os.path.basename(output_path)}")
            self.progress_var.set(100)
            
            # 成功メッセージとフォルダオープン
            self.show_export_success(output_path)
            
            # プログレスバーリセット
            self.root.after(3000, lambda: self.progress_var.set(0))
            
        except Exception as e:
            self.export_status.config(text="❌ エラー", foreground='red')
            self.show_export_error(str(e))
            self.progress_var.set(0)
    
    def safe_export_excel(self, output_path):
        """安全なExcel保存"""
        try:
            # openpyxlを使用して保存
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # メインデータシート
                self.data.to_excel(writer, sheet_name='ProcessedData', index=False)
                self.progress_var.set(40)
                self.root.update()
                
                # 処理サマリーシート
                if self.include_summary.get():
                    summary_data = self.create_processing_summary()
                    summary_df = pd.DataFrame(summary_data[1:], columns=summary_data[0])
                    summary_df.to_excel(writer, sheet_name='Summary', index=False)
                    self.progress_var.set(60)
                    self.root.update()
                
                # 処理ログシート
                if self.include_log.get() and self.processing_log:
                    log_df = pd.DataFrame(self.processing_log, columns=['処理履歴'])
                    log_df.to_excel(writer, sheet_name='Log', index=False)
                    self.progress_var.set(80)
                    self.root.update()
                
        except ImportError:
            # openpyxlがない場合はシンプル保存
            print("WARNING: openpyxl not available, using simple Excel export")
            self.data.to_excel(output_path, index=False)
        except Exception as e:
            print(f"ERROR: Excel保存エラー: {e}")
            # 最後の手段として簡単保存
            self.data.to_excel(output_path, index=False)
    
    def safe_export_csv(self, output_path):
        """安全なCSV保存"""
        # BOM付きUTF-8で保存（Excelで文字化けしない）
        self.data.to_csv(output_path, index=False, encoding='utf-8-sig')
        self.progress_var.set(60)
        self.root.update()
        
        # 処理サマリーも別ファイルで保存
        if self.include_summary.get():
            summary_path = output_path.replace('.csv', '_summary.csv')
            summary_data = self.create_processing_summary()
            summary_df = pd.DataFrame(summary_data[1:], columns=summary_data[0])
            summary_df.to_csv(summary_path, index=False, encoding='utf-8-sig')
        
        # 処理ログも別ファイルで保存
        if self.include_log.get() and self.processing_log:
            log_path = output_path.replace('.csv', '_log.csv')
            log_df = pd.DataFrame(self.processing_log, columns=['処理履歴'])
            log_df.to_csv(log_path, index=False, encoding='utf-8-sig')
        
        self.progress_var.set(90)
        self.root.update()
    
    def quick_export_excel(self):
        """簡単Excel保存"""
        if self.data is None:
            messagebox.showwarning("データなし", 
                "❌ データが読み込まれていません。\n\n"
                "📂 「ファイル選択」ボタンでファイルを読み込んでください。")
            return
        
        try:
            # デフォルトファイル名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if self.file_path:
                base_dir = os.path.dirname(self.file_path)
                base_name = os.path.splitext(os.path.basename(self.file_path))[0]
                default_path = os.path.join(base_dir, f"{base_name}_excel_{timestamp}.xlsx")
            else:
                default_path = f"data_excel_{timestamp}.xlsx"
            
            # 進捗表示
            self.update_status("Excel保存中...")
            self.progress_var.set(50)
            self.root.update()
            
            # Excel保存
            self.data.to_excel(default_path, index=False)
            
            self.progress_var.set(100)
            self.log_action(f"Excel保存完了: {os.path.basename(default_path)}")
            self.update_status(f"✅ Excel保存完了")
            
            # 成功メッセージ
            result = messagebox.askyesno("Excel保存完了", 
                f"✅ Excelファイルを保存しました！\n\n"
                f"📁 {os.path.basename(default_path)}\n"
                f"📊 {len(self.data):,}行 × {len(self.data.columns)}列\n\n"
                f"保存先フォルダを開きますか？")
            
            if result:
                self.open_folder(default_path)
            
            # プログレスバーリセット
            self.root.after(2000, lambda: self.progress_var.set(0))
            
        except Exception as e:
            self.show_export_error(str(e))
            self.progress_var.set(0)
    
    def quick_export_csv(self):
        """簡単CSV保存"""
        if self.data is None:
            messagebox.showwarning("データなし", 
                "❌ データが読み込まれていません。\n\n"
                "📂 「ファイル選択」ボタンでファイルを読み込んでください。")
            return
        
        try:
            # デフォルトファイル名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            if self.file_path:
                base_dir = os.path.dirname(self.file_path)
                base_name = os.path.splitext(os.path.basename(self.file_path))[0]
                default_path = os.path.join(base_dir, f"{base_name}_csv_{timestamp}.csv")
            else:
                default_path = f"data_csv_{timestamp}.csv"
            
            # 進捗表示
            self.update_status("CSV保存中...")
            self.progress_var.set(50)
            self.root.update()
            
            # CSV保存（BOM付きUTF-8）
            self.data.to_csv(default_path, index=False, encoding='utf-8-sig')
            
            self.progress_var.set(100)
            self.log_action(f"CSV保存完了: {os.path.basename(default_path)}")
            self.update_status(f"✅ CSV保存完了")
            
            # 成功メッセージ
            result = messagebox.askyesno("CSV保存完了", 
                f"✅ CSVファイルを保存しました！\n\n"
                f"📁 {os.path.basename(default_path)}\n"
                f"📊 {len(self.data):,}行 × {len(self.data.columns)}列\n\n"
                f"保存先フォルダを開きますか？")
            
            if result:
                self.open_folder(default_path)
            
            # プログレスバーリセット
            self.root.after(2000, lambda: self.progress_var.set(0))
            
        except Exception as e:
            self.show_export_error(str(e))
            self.progress_var.set(0)
    
    def create_processing_summary(self):
        """処理サマリーデータを作成"""
        return [
            ['項目', '内容'],
            ['元ファイル名', os.path.basename(self.file_path) if self.file_path else '不明'],
            ['処理実行日時', datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
            ['データサイズ', f"{len(self.data):,}行 × {len(self.data.columns)}列"],
            ['空セル数', f"{self.data.isnull().sum().sum():,}個"],
            ['重複行数', f"{self.data.duplicated().sum():,}行"],
            ['実行処理数', f"{len(self.processing_log)}件"],
            ['メモリ使用量', f"{self.data.memory_usage(deep=True).sum()/1024/1024:.2f} MB"],
            ['Tableau対応', '✅ Tableau分析用に最適化済み'],
        ]
    
    def show_export_success(self, output_path):
        """エクスポート成功メッセージを表示"""
        file_name = os.path.basename(output_path)
        file_dir = os.path.dirname(output_path)
        file_size = os.path.getsize(output_path) / 1024  # KB
        
        success_msg = f"🎉 データ保存が完了しました！\n\n"
        success_msg += f"📁 ファイル名: {file_name}\n"
        success_msg += f"📂 保存場所: {file_dir}\n"
        success_msg += f"📊 データ: {len(self.data):,}行 × {len(self.data.columns)}列\n"
        success_msg += f"💾 ファイルサイズ: {file_size:.1f} KB\n"
        success_msg += f"🔧 実行処理: {len(self.processing_log)}件\n\n"
        success_msg += f"🚀 このファイルをTableauで直接読み込めます！"
        
        result = messagebox.askyesno("保存完了", 
            success_msg + "\n\n保存先フォルダを開きますか？")
        
        if result:
            self.open_folder(output_path)
    
    def show_export_error(self, error_message):
        """エクスポートエラーメッセージを表示"""
        error_msg = f"❌ 保存エラーが発生しました\n\n"
        error_msg += f"エラー詳細:\n{error_message}\n\n"
        error_msg += f"解決策:\n"
        
        if "Permission denied" in error_message or "being used" in error_message:
            error_msg += f"• 同名ファイルがExcelで開かれている場合は閉じてください\n"
            error_msg += f"• 保存先フォルダの書き込み権限を確認してください\n"
        elif "openpyxl" in error_message:
            error_msg += f"• pip install openpyxl を実行してください\n"
        else:
            error_msg += f"• 別の形式（CSV）を試してください\n"
            error_msg += f"• 別の保存場所を試してください\n"
        
        error_msg += f"\n💡 「🔧 ボタン有効化」を押してから再試行してください"
        
        messagebox.showerror("保存エラー", error_msg)
    
    def open_folder(self, file_path):
        """保存先フォルダを開く"""
        try:
            import platform
            import subprocess
            
            system = platform.system()
            folder_path = os.path.dirname(file_path)
            
            if system == "Windows":
                subprocess.run(['explorer', '/select,', file_path.replace('/', '\\')])
            elif system == "Darwin":  # macOS
                subprocess.run(['open', '-R', file_path])
            else:  # Linux
                subprocess.run(['xdg-open', folder_path])
                
        except Exception as e:
            print(f"フォルダを開けませんでした: {e}")
            try:
                # フォルダだけ開く
                folder_path = os.path.dirname(file_path)
                if platform.system() == "Windows":
                    os.startfile(folder_path)
                else:
                    os.system(f"open '{folder_path}'" if platform.system() == "Darwin" else f"xdg-open '{folder_path}'")
            except:
                pass
    
    # ===== データ処理メソッド =====
    def remove_empty_rows(self):
        """空行を削除"""
        if self.data is None:
            return
        
        try:
            initial_rows = len(self.data)
            self.data = self.data.dropna(how='all')
            removed_rows = initial_rows - len(self.data)
            
            self.update_data_info()
            self.update_data_table()
            
            self.log_action(f"空行削除: {removed_rows}行削除")
            self.update_status(f"✅ {removed_rows}行の空行を削除しました")
            
        except Exception as e:
            self.show_error(f"空行削除エラー: {str(e)}")
    
    def remove_empty_columns(self):
        """空列を削除"""
        if self.data is None:
            return
        
        try:
            initial_cols = len(self.data.columns)
            self.data = self.data.dropna(axis=1, how='all')
            removed_cols = initial_cols - len(self.data.columns)
            
            self.update_data_info()
            self.update_data_table()
            
            self.log_action(f"空列削除: {removed_cols}列削除")
            self.update_status(f"✅ {removed_cols}列の空列を削除しました")
            
        except Exception as e:
            self.show_error(f"空列削除エラー: {str(e)}")
    
    def remove_duplicates(self):
        """重複行を削除"""
        if self.data is None:
            return
        
        try:
            initial_rows = len(self.data)
            self.data = self.data.drop_duplicates()
            removed_rows = initial_rows - len(self.data)
            
            self.update_data_info()
            self.update_data_table()
            
            self.log_action(f"重複行削除: {removed_rows}行削除")
            self.update_status(f"✅ {removed_rows}行の重複を削除しました")
            
        except Exception as e:
            self.show_error(f"重複行削除エラー: {str(e)}")
    
    def clean_text_data(self):
        """テキストデータをクリーニング"""
        if self.data is None:
            return
        
        try:
            text_columns = self.data.select_dtypes(include=['object']).columns
            processed_cols = 0
            
            for col in text_columns:
                # 前後の空白を削除
                self.data[col] = self.data[col].astype(str).str.strip()
                # 連続する空白を1つに
                self.data[col] = self.data[col].str.replace(r'\s+', ' ', regex=True)
                processed_cols += 1
            
            self.update_data_info()
            self.update_data_table()
            
            self.log_action(f"テキストクリーニング: {processed_cols}列処理")
            self.update_status(f"✅ {processed_cols}列のテキストを整形しました")
            
        except Exception as e:
            self.show_error(f"テキストクリーニングエラー: {str(e)}")
    
    def convert_data_types(self):
        """データ型を変換"""
        if self.data is None:
            return
        
        try:
            converted_cols = 0
            
            for col in self.data.columns:
                if self.data[col].dtype == 'object':
                    # 数値変換を試行
                    numeric_series = pd.to_numeric(self.data[col], errors='coerce')
                    if numeric_series.notna().sum() / len(self.data[col]) > 0.8:
                        self.data[col] = numeric_series
                        converted_cols += 1
                    else:
                        # 日付変換を試行
                        try:
                            date_series = pd.to_datetime(self.data[col], errors='coerce')
                            if date_series.notna().sum() / len(self.data[col]) > 0.5:
                                self.data[col] = date_series
                                converted_cols += 1
                        except:
                            pass
            
            self.update_data_info()
            self.update_data_table()
            
            self.log_action(f"データ型変換: {converted_cols}列変換")
            self.update_status(f"✅ {converted_cols}列のデータ型を変換しました")
            
        except Exception as e:
            self.show_error(f"データ型変換エラー: {str(e)}")
    
    def fill_missing_dialog(self):
        """欠損値処理ダイアログを表示"""
        if self.data is None:
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("欠損値処理")
        dialog.geometry("400x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 処理方法選択
        ttk.Label(dialog, text="処理方法を選択:", font=('Arial', 10, 'bold')).pack(pady=10)
        
        method_var = tk.StringVar(value='remove')
        methods = [
            ('remove', '空行を削除'),
            ('forward', '前の値で埋める'),
            ('zero', '0で埋める'),
            ('custom', 'カスタム値で埋める')
        ]
        
        for value, text in methods:
            ttk.Radiobutton(dialog, text=text, variable=method_var, value=value).pack(anchor=tk.W, padx=20)
        
        # カスタム値入力
        custom_frame = ttk.Frame(dialog)
        custom_frame.pack(pady=10)
        ttk.Label(custom_frame, text="カスタム値:").pack(side=tk.LEFT)
        custom_var = tk.StringVar()
        ttk.Entry(custom_frame, textvariable=custom_var).pack(side=tk.LEFT, padx=5)
        
        # ボタン
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        def apply_fill():
            try:
                method = method_var.get()
                initial_nulls = self.data.isnull().sum().sum()
                
                if method == 'remove':
                    self.data = self.data.dropna()
                elif method == 'forward':
                    self.data = self.data.fillna(method='ffill')
                elif method == 'zero':
                    self.data = self.data.fillna(0)
                elif method == 'custom':
                    custom_value = custom_var.get()
                    self.data = self.data.fillna(custom_value)
                
                final_nulls = self.data.isnull().sum().sum()
                processed_count = initial_nulls - final_nulls
                
                self.update_data_info()
                self.update_data_table()
                
                self.log_action(f"欠損値処理: {processed_count}個処理 (方法: {method})")
                self.update_status(f"✅ {processed_count}個の欠損値を処理しました")
                
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("エラー", f"欠損値処理エラー: {str(e)}")
        
        ttk.Button(button_frame, text="適用", command=apply_fill).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def rename_columns_dialog(self):
        """列名変更ダイアログを表示"""
        if self.data is None:
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("列名変更")
        dialog.geometry("500x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 列リスト
        ttk.Label(dialog, text="変更する列を選択:", font=('Arial', 10, 'bold')).pack(pady=10)
        
        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20)
        
        listbox = tk.Listbox(list_frame)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=listbox.yview)
        listbox.config(yscrollcommand=scrollbar.set)
        
        for col in self.data.columns:
            listbox.insert(tk.END, col)
        
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 新しい名前入力
        input_frame = ttk.Frame(dialog)
        input_frame.pack(pady=10)
        ttk.Label(input_frame, text="新しい列名:").pack(side=tk.LEFT)
        new_name_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=new_name_var).pack(side=tk.LEFT, padx=5)
        
        # ボタン
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        def apply_rename():
            try:
                selection = listbox.curselection()
                if not selection:
                    messagebox.showwarning("警告", "列を選択してください")
                    return
                
                old_name = listbox.get(selection[0])
                new_name = new_name_var.get().strip()
                
                if not new_name:
                    messagebox.showwarning("警告", "新しい列名を入力してください")
                    return
                
                self.data = self.data.rename(columns={old_name: new_name})
                
                self.update_data_info()
                self.update_data_table()
                
                self.log_action(f"列名変更: {old_name} -> {new_name}")
                self.update_status(f"✅ 列名を変更しました: {old_name} -> {new_name}")
                
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("エラー", f"列名変更エラー: {str(e)}")
        
        ttk.Button(button_frame, text="変更", command=apply_rename).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def filter_data_dialog(self):
        """データフィルタダイアログを表示"""
        if self.data is None:
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("データフィルタ")
        dialog.geometry("450x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 列選択
        ttk.Label(dialog, text="フィルタする列:", font=('Arial', 10, 'bold')).pack(pady=5)
        column_var = tk.StringVar()
        column_combo = ttk.Combobox(dialog, textvariable=column_var, values=list(self.data.columns))
        column_combo.pack(pady=5)
        
        # 条件選択
        ttk.Label(dialog, text="条件:", font=('Arial', 10, 'bold')).pack(pady=5)
        condition_var = tk.StringVar(value='==')
        conditions = ['==', '!=', '>', '<', '>=', '<=', 'contains']
        condition_combo = ttk.Combobox(dialog, textvariable=condition_var, values=conditions)
        condition_combo.pack(pady=5)
        
        # 値入力
        ttk.Label(dialog, text="値:", font=('Arial', 10, 'bold')).pack(pady=5)
        value_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=value_var).pack(pady=5)
        
        # ボタン
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        def apply_filter():
            try:
                column = column_var.get()
                condition = condition_var.get()
                value = value_var.get()
                
                if not column or column not in self.data.columns:
                    messagebox.showwarning("警告", "有効な列を選択してください")
                    return
                
                if not value:
                    messagebox.showwarning("警告", "フィルタ値を入力してください")
                    return
                
                initial_rows = len(self.data)
                
                if condition == '==':
                    self.data = self.data[self.data[column] == value]
                elif condition == '!=':
                    self.data = self.data[self.data[column] != value]
                elif condition == '>':
                    self.data = self.data[pd.to_numeric(self.data[column], errors='coerce') > float(value)]
                elif condition == '<':
                    self.data = self.data[pd.to_numeric(self.data[column], errors='coerce') < float(value)]
                elif condition == '>=':
                    self.data = self.data[pd.to_numeric(self.data[column], errors='coerce') >= float(value)]
                elif condition == '<=':
                    self.data = self.data[pd.to_numeric(self.data[column], errors='coerce') <= float(value)]
                elif condition == 'contains':
                    self.data = self.data[self.data[column].astype(str).str.contains(value, na=False)]
                
                filtered_rows = initial_rows - len(self.data)
                
                self.update_data_info()
                self.update_data_table()
                
                self.log_action(f"データフィルタ: {column} {condition} {value} ({filtered_rows}行除外)")
                self.update_status(f"✅ {filtered_rows}行をフィルタしました")
                
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("エラー", f"フィルタエラー: {str(e)}")
        
        ttk.Button(button_frame, text="適用", command=apply_filter).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="キャンセル", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def reset_data(self):
        """データをリセット"""
        if self.original_data is None:
            return
        
        if messagebox.askyesno("確認", "データをリセットしますか？\nすべての変更が失われます。"):
            self.data = self.original_data.copy()
            self.update_data_info()
            self.update_data_table()
            self.processing_log = []
            self.log_action("データリセット完了")
            self.update_status("✅ データをリセットしました")
    
    def show_processing_log(self):
        """処理ログを表示"""
        if not self.processing_log:
            messagebox.showinfo("処理ログ", "処理ログがありません")
            return
        
        log_window = tk.Toplevel(self.root)
        log_window.title("処理ログ")
        log_window.geometry("600x400")
        
        text_widget = scrolledtext.ScrolledText(log_window, wrap=tk.WORD, font=('Consolas', 9))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        for log in self.processing_log:
            text_widget.insert(tk.END, log + '\n')
        
        text_widget.config(state='disabled')


def main():
    """メイン関数"""
    root = tk.Tk()
    app = TableauPreprocessorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
