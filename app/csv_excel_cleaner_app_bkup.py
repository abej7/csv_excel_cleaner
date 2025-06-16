# app/csv_excel_cleaner_app.py

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import pandas as pd
import os

class DataCleanerApp:
    def __init__(self, root):
        # ウィンドウ初期設定
        self.root = root
        self.root.title("CSV/Excel Data Cleaner")
        self.root.geometry("1000x620")

        # pandasデータフレーム
        self.df = None

        # UI部品構築
        self.create_widgets()

        # チェックボックス格納用
        self.column_vars = {}  # 列名: BooleanVar
        self.column_check_frame = None  # 再描画のため保持

    def create_widgets(self):
        """画面全体のレイアウト構築"""

        # === 🎨 テーマ切り替えセクション（画面上部） ===

        # 使用可能なテーマ名一覧を取得
        theme_names = self.root.style.theme_names()

        # テーマ切り替え用のラベル付きフレームを作成（見た目のセクション化）
        frame_theme = tb.LabelFrame(
            self.root,
            text="🎨 Theme Selector",     # セクションタイトル
            padding=10,
            bootstyle="secondary"
        )
        frame_theme.pack(fill=X, padx=10, pady=(5, 5))

        # 現在選択されているテーマを保持する変数（ドロップダウンと連携）
        self.theme_var = tb.StringVar(value=self.root.style.theme.name)

        # テーマ選択用のドロップダウン（Combobox）
        combo_theme = tb.Combobox(
            frame_theme,
            textvariable=self.theme_var,
            values=theme_names,
            state="readonly",       # 手入力不可にする
            width=20,
            bootstyle="info"
        )
        combo_theme.pack(side=LEFT, padx=5)

        # Applyボタン：選択されたテーマを適用する処理を実行
        btn_apply_theme = tb.Button(
            frame_theme,
            text="Apply",
            command=self.change_theme,
            bootstyle="primary"
        )
        btn_apply_theme.pack(side=LEFT, padx=5)

        # === 上部：ファイル選択エリア ===
        frame_top = tb.LabelFrame(self.root, text="📂 File Loader", padding=10, bootstyle="info")
        frame_top.pack(fill=X, padx=10, pady=(10, 5))

        self.file_path_var = tb.StringVar()
        entry = tb.Entry(frame_top, textvariable=self.file_path_var, width=60, bootstyle="info")
        entry.pack(side=LEFT, padx=5)

        btn_browse = tb.Button(frame_top, text="Browse", command=self.browse_file, bootstyle="primary-outline")
        btn_browse.pack(side=LEFT, padx=5)

        btn_load = tb.Button(frame_top, text="Load", command=self.load_file, bootstyle="success")
        btn_load.pack(side=LEFT, padx=5)

        # === CSVヘッダー行の扱いオプション ===
        self.use_header_var = tb.BooleanVar(value=True)
        chk_header = tb.Checkbutton(
            frame_top,
            text="Use First Row as Header",
            variable=self.use_header_var,
            bootstyle="info"
        )
        chk_header.pack(side=LEFT, padx=5)

        # === 中央：左右2カラム構成 ===
        frame_main = tb.Frame(self.root, padding=10)
        frame_main.pack(fill=BOTH, expand=True)

        # --- 左パネル：クリーニングオプション ---
        frame_left = tb.LabelFrame(frame_main, text="🔧 Clean Options", padding=10, bootstyle="secondary")
        frame_left.pack(side=LEFT, fill=Y, padx=(0, 10))

        # 欠損値削除チェックボックス
        self.dropna_var = tb.BooleanVar(value=True)
        chk_dropna = tb.Checkbutton(frame_left, text="Drop NaN (欠損値削除)", variable=self.dropna_var, bootstyle="info")
        chk_dropna.pack(anchor=W, pady=5)

        # 空白除去チェックボックス
        self.trim_var = tb.BooleanVar(value=True)
        chk_trim = tb.Checkbutton(frame_left, text="Trim Spaces (前後空白除去)", variable=self.trim_var, bootstyle="info")
        chk_trim.pack(anchor=W, pady=5)

        # 実行ボタン
        btn_clean = tb.Button(frame_left, text="Clean", command=self.clean_data, bootstyle="info")
        btn_clean.pack(pady=(15, 0), fill=X)

        # --- 左：列削除用セクション ---
        frame_cols = tb.LabelFrame(frame_left, text="🗃 Column Remover", padding=10, bootstyle="secondary")
        frame_cols.pack(fill=X, pady=(20, 5))

        # → 後で動的にチェックボックスを追加
        self.column_check_frame = tb.Frame(frame_cols)
        self.column_check_frame.pack(fill=X)

        btn_delete_cols = tb.Button(frame_cols, text="Delete Columns", command=self.delete_selected_columns, bootstyle="danger")
        btn_delete_cols.pack(pady=5, fill=X)

        # 小数点丸めの有効化チェック
        self.round_var = tb.BooleanVar(value=False)
        chk_round = tb.Checkbutton(
            frame_left,
            text="Round Numbers(桁丸め)",
            variable=self.round_var,
            bootstyle="info"
        )
        chk_round.pack(anchor=W, pady=(10, 2))

        # 小数点桁数の選択（0～5）
        self.round_digits_var = tb.IntVar(value=2)
        digits = [0, 1, 2, 3, 4, 5]
        combo_digits = tb.Combobox(
            frame_left,
            textvariable=self.round_digits_var,
            values=digits,
            width=5,
            state="readonly",
            bootstyle="info"
        )
        combo_digits.pack(anchor=W, padx=10)

        # === 保存形式ドロップダウンと保存ボタン ===
        frame_save = tb.LabelFrame(frame_left, text="💾 Export", padding=10, bootstyle="secondary")
        frame_save.pack(fill=X, pady=(20, 5))

        # 保存形式選択（CSV or Excel）
        self.save_format_var = tb.StringVar(value="CSV")
        combo_format = tb.Combobox(
            frame_save,
            textvariable=self.save_format_var,
            values=["CSV", "Excel"],
            state="readonly",
            width=10,
            bootstyle="info"
        )
        combo_format.pack(pady=(0, 5), anchor=W)

        # 保存ボタン
        btn_save = tb.Button(
            frame_save,
            text="Save",
            command=self.save_file,
            bootstyle="success"
        )
        btn_save.pack(fill=X)

        # --- 右パネル：データプレビュー ---
        frame_right = tb.LabelFrame(frame_main, text="📊 Data Preview", padding=10, bootstyle="info")
        frame_right.pack(side=RIGHT, fill=BOTH, expand=True)

        self.tree = tb.Treeview(frame_right, bootstyle="info")
        self.tree.pack(fill=BOTH, expand=True)

        # === 下部：ログエリア ===
        frame_log = tb.LabelFrame(self.root, text="📋 Activity Log", padding=5, bootstyle="info")
        frame_log.pack(fill=X, padx=10, pady=(0, 10))

        self.log_text = tb.ScrolledText(frame_log, height=5)
        self.log_text.pack(fill=X)

        self.log("アプリを起動しました")

    def browse_file(self):
        """ファイル選択ダイアログを表示"""
        filetypes = [("CSV files", "*.csv"), ("Excel files", "*.xlsx")]
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path:
            self.file_path_var.set(path)

    def load_file(self):
        """ファイルを読み込み、DataFrameに格納"""
        path = self.file_path_var.get()
        if not path:
            messagebox.showwarning("No file selected", "Please choose a file.")
            return

        try:
            ext = os.path.splitext(path)[-1].lower()
            if ext == ".csv":
                header_opt = 0 if self.use_header_var.get() else None
                self.df = pd.read_csv(path, header=header_opt)
            elif ext == ".xlsx":
                header_opt = 0 if self.use_header_var.get() else None
                self.df = pd.read_excel(path, header=header_opt)
            else:
                raise ValueError("Unsupported file format.")

            self.display_data()
            self.log(f"読み込み成功: {os.path.basename(path)}")
            self.create_column_checkboxes()

        except Exception as e:
            messagebox.showerror("読み込みエラー", str(e))
            self.log(f"読み込みエラー: {e}")

    def display_data(self):
        """DataFrameをTreeviewに表示"""
        self.tree.delete(*self.tree.get_children())

        if self.df is None or self.df.empty:
            self.log("データが空です")
            return

        self.tree["columns"] = list(self.df.columns)
        self.tree["show"] = "headings"

        for col in self.df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="w", width=120)

        for _, row in self.df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def clean_data(self):
        """クリーニング処理を実行（欠損値削除、空白除去）"""
        if self.df is None:
            self.log("データが読み込まれていません")
            return

        original_shape = self.df.shape

        # 欠損値削除
        if self.dropna_var.get():
            self.df = self.df.dropna()
            self.log("欠損値を含む行を削除しました")

        # 空白除去
        if self.trim_var.get():
            self.df = self.df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            self.log("文字列の前後空白を除去しました")

        # --- 数値の丸め処理（round） ---
        if self.round_var.get():
            digits = self.round_digits_var.get()
            num_cols = self.df.select_dtypes(include='number').columns

            if not num_cols.empty:
                self.df[num_cols] = self.df[num_cols].round(digits)
                self.log(f"数値を小数点以下 {digits} 桁に丸めました")
            else:
                self.log("丸め対象の数値列がありませんでした")

        self.display_data()

        new_shape = self.df.shape
        self.log(f"データ件数: {original_shape[0]} → {new_shape[0]} 行")

    def create_column_checkboxes(self):
        """列名をもとにチェックボックスを動的生成"""

        # 前回のチェックを削除
        for widget in self.column_check_frame.winfo_children():
            widget.destroy()

        self.column_vars.clear()

        # DataFrameが読み込まれていない場合はスキップ
        if self.df is None:
            return

        for col in self.df.columns:
            var = tb.BooleanVar(value=False)
            chk = tb.Checkbutton(
                self.column_check_frame,
                text=col,
                variable=var,
                bootstyle="info"
            )
            chk.pack(anchor=W)
            self.column_vars[col] = var

    def delete_selected_columns(self):
        """チェックされた列をDataFrameから削除"""
        if self.df is None:
            self.log("データが読み込まれていません")
            return

        selected = [col for col, var in self.column_vars.items() if var.get()]
        if not selected:
            self.log("削除する列が選択されていません")
            return

        self.df = self.df.drop(columns=selected)
        self.log(f"列を削除しました: {', '.join(selected)}")

        # 表示とチェックエリアを更新
        self.display_data()
        self.create_column_checkboxes()

    def save_file(self):
        """保存形式に応じてファイル保存（CSV or Excel）"""
        if self.df is None:
            self.log("保存するデータがありません")
            return

        filetypes = []
        def_ext = ""
        fmt = self.save_format_var.get()

        # ファイル形式とダイアログ設定
        if fmt == "CSV":
            filetypes = [("CSVファイル", "*.csv")]
            def_ext = ".csv"
        elif fmt == "Excel":
            filetypes = [("Excelファイル", "*.xlsx")]
            def_ext = ".xlsx"
        else:
            self.log("未対応の保存形式です")
            return

        # 保存ダイアログ表示
        path = filedialog.asksaveasfilename(defaultextension=def_ext, filetypes=filetypes)
        if not path:
            return

        try:
            if fmt == "CSV":
                self.df.to_csv(path, index=False)
            elif fmt == "Excel":
                self.df.to_excel(path, index=False)
            self.log(f"ファイルを保存しました: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("保存エラー", str(e))
            self.log(f"保存エラー: {e}")

    def change_theme(self):
        """テーマを切り替える"""
        new_theme = self.theme_var.get()
        try:
            self.root.style.theme_use(new_theme)
            self.log(f"テーマを変更しました: {new_theme}")
        except Exception as e:
            self.log(f"テーマ変更エラー: {e}")

    def log(self, message):
        """ログエリアにメッセージを出力"""
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")


# === アプリ実行 ===
if __name__ == "__main__":
    app = tb.Window(themename="journal")
    DataCleanerApp(app)
    app.mainloop()
