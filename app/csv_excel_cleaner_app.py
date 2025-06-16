# 完全版：CSV/Excel一括データクリーナー（全機能実装・日本語UI対応）
import os
import pandas as pd
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox

class DataCleanerApp:
    def __init__(self, root):
        self.root = root

        # ウィンドウタイトル（英語を維持）
        self.root.title("CSV/Excel Data Cleaner")
        self.root.geometry("1000x750")

        # pandasデータフレーム
        self.df = None

        # 列削除用チェックボックス群
        self.column_vars = {}
        self.column_check_frame = None

        self.create_widgets()

    def create_widgets(self):
        """画面全体のレイアウト構築"""

        # アプリ名タイトル（中央）
        label_title = tb.Label(
            self.root,
            text="CSV/Excel一括データクリーナー",
            font=("Helvetica", 18, "bold"),
            anchor="center"
        )
        label_title.pack(fill=X, pady=(10, 5))

        # === テーマ切替エリア ===
        theme_names = self.root.style.theme_names()
        frame_theme = tb.LabelFrame(self.root, text="🎨 Theme Selector", padding=10, bootstyle="secondary")
        frame_theme.pack(fill=X, padx=10, pady=(5, 5))

        self.theme_var = tb.StringVar(value=self.root.style.theme.name)
        combo_theme = tb.Combobox(frame_theme, textvariable=self.theme_var, values=theme_names,
                                  state="readonly", width=20, bootstyle="info")
        combo_theme.pack(side=LEFT, padx=5)

        btn_apply_theme = tb.Button(frame_theme, text="適用", command=self.change_theme, bootstyle="primary")
        btn_apply_theme.pack(side=LEFT, padx=5)

        # === ファイル読み込みエリア ===
        frame_top = tb.LabelFrame(self.root, text="📂 File Loader", padding=10, bootstyle="info")
        frame_top.pack(fill=X, padx=10, pady=(5, 5))

        self.file_path_var = tb.StringVar()
        entry = tb.Entry(frame_top, textvariable=self.file_path_var, width=60, bootstyle="info")
        entry.pack(side=LEFT, padx=5)

        btn_browse = tb.Button(frame_top, text="ファイル選択", command=self.browse_file, bootstyle="primary-outline")
        btn_browse.pack(side=LEFT, padx=5)

        btn_load = tb.Button(frame_top, text="読み込み", command=self.load_file, bootstyle="success")
        btn_load.pack(side=LEFT, padx=5)

        self.use_header_var = tb.BooleanVar(value=True)
        chk_header = tb.Checkbutton(frame_top, text="Use First Row as Header", variable=self.use_header_var,
                                    bootstyle="info")
        chk_header.pack(side=LEFT, padx=5)

        # === 中央メインレイアウト ===
        frame_main = tb.Frame(self.root, padding=10)
        frame_main.pack(fill=BOTH, expand=True)

        # === 左パネル：操作エリア ===
        frame_left = tb.LabelFrame(frame_main, text="🔧 Clean Options", padding=10, bootstyle="secondary")
        frame_left.pack(side=LEFT, fill=Y, padx=(0, 10))

        self.dropna_var = tb.BooleanVar(value=True)
        chk_dropna = tb.Checkbutton(frame_left, text="Drop NaN (欠損値削除)", variable=self.dropna_var, bootstyle="info")
        chk_dropna.pack(anchor=W)

        self.trim_var = tb.BooleanVar(value=True)
        chk_trim = tb.Checkbutton(frame_left, text="Trim Spaces (前後空白除去)", variable=self.trim_var, bootstyle="info")
        chk_trim.pack(anchor=W)

        self.round_var = tb.BooleanVar(value=False)
        chk_round = tb.Checkbutton(frame_left, text="Round Numbers(桁丸め)", variable=self.round_var, bootstyle="info")
        chk_round.pack(anchor=W, pady=(10, 2))

        self.round_digits_var = tb.IntVar(value=2)
        combo_digits = tb.Combobox(frame_left, textvariable=self.round_digits_var, values=[0,1,2,3,4,5],
                                   width=5, state="readonly", bootstyle="info")
        combo_digits.pack(anchor=W, padx=10)

        btn_clean = tb.Button(frame_left, text="クリーニング", command=self.clean_data, bootstyle="primary")
        btn_clean.pack(pady=(15, 0), fill=X)

        # === 列削除セクション ===
        frame_cols = tb.LabelFrame(frame_left, text="🗃 Column Remover", padding=10, bootstyle="secondary")
        frame_cols.pack(fill=X, pady=(20, 5))

        self.column_check_frame = tb.Frame(frame_cols)
        self.column_check_frame.pack(fill=X)

        btn_delete_cols = tb.Button(frame_cols, text="列を削除", command=self.delete_selected_columns,
                                    bootstyle="danger")
        btn_delete_cols.pack(pady=5, fill=X)

        # === 保存セクション ===
        frame_save = tb.LabelFrame(frame_left, text="💾 Export", padding=10, bootstyle="secondary")
        frame_save.pack(fill=X, pady=(20, 5))

        self.save_format_var = tb.StringVar(value="CSV")
        combo_format = tb.Combobox(frame_save, textvariable=self.save_format_var,
                                   values=["CSV", "Excel"], state="readonly", width=10, bootstyle="info")
        combo_format.pack(pady=(0, 5), anchor=W)

        btn_save = tb.Button(frame_save, text="保存", command=self.save_file, bootstyle="success")
        btn_save.pack(fill=X)

        # === 右パネル：データ表示 ===
        frame_right = tb.LabelFrame(frame_main, text="📊 Data Preview", padding=10, bootstyle="info")
        frame_right.pack(side=RIGHT, fill=BOTH, expand=True)

        self.tree = tb.Treeview(frame_right, bootstyle="info")
        self.tree.pack(fill=BOTH, expand=True)

        # === ログ出力 ===
        frame_log = tb.LabelFrame(self.root, text="📋 Activity Log", bootstyle="info", padding=5)
        frame_log.pack(fill=X, padx=10, pady=(0, 10))

        self.log_text = tb.ScrolledText(frame_log, height=8)
        self.log_text.pack(fill=X, padx=5, pady=(0, 5))

        self.log("アプリを起動しました")


        # ファイルダイアログでCSVまたはExcelファイルを選択し、パスを入力欄に反映
    def browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        if path:
            self.file_path_var.set(path)


        # 選択されたファイルをpandasで読み込み、DataFrameに格納
    def load_file(self):
        path = self.file_path_var.get()
        if not path:
            messagebox.showwarning("No file selected", "Please choose a file.")
            return
        try:
            ext = os.path.splitext(path)[-1].lower()
            header_opt = 0 if self.use_header_var.get() else None
            if ext == ".csv":
                self.df = pd.read_csv(path, header=header_opt)
            elif ext == ".xlsx":
                self.df = pd.read_excel(path, header=header_opt)
            else:
                raise ValueError("Unsupported file format.")

            self.display_data()
            self.create_column_checkboxes()
            self.log(f"読み込み成功: {os.path.basename(path)}")

        except Exception as e:
            messagebox.showerror("読み込みエラー", str(e))
            self.log(f"読み込みエラー: {e}")


        # DataFrameの内容をTreeviewに表示
    def display_data(self):
        self.tree.delete(*self.tree.get_children())
        if self.df is None or self.df.empty:
            return
        self.tree["columns"] = list(self.df.columns)
        self.tree["show"] = "headings"
        for col in self.df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="w", width=120)
        for _, row in self.df.iterrows():
            self.tree.insert("", "end", values=list(row))


        # チェックされたクリーニング処理を実行（欠損・空白・丸め）
    def clean_data(self):
        if self.df is None:
            self.log("データが読み込まれていません")
            return
        if self.dropna_var.get():
            self.df = self.df.dropna()
            self.log("欠損値を含む行を削除しました")
        if self.trim_var.get():
            self.df = self.df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            self.log("文字列の前後空白を除去しました")
        if self.round_var.get():
            digits = self.round_digits_var.get()
            num_cols = self.df.select_dtypes(include='number').columns
            if len(num_cols) > 0:
                self.df[num_cols] = self.df[num_cols].round(digits)
                self.log(f"数値を小数点以下 {digits} 桁に丸めました")
        self.display_data()


        # DataFrameの列名に応じた削除対象列のチェックボックスを動的に生成
    def create_column_checkboxes(self):
        if self.column_check_frame is None:
            return
        for widget in self.column_check_frame.winfo_children():
            widget.destroy()
        self.column_vars.clear()
        if self.df is None:
            return
        for col in self.df.columns:
            var = tb.BooleanVar(value=False)
            chk = tb.Checkbutton(self.column_check_frame, text=col, variable=var, bootstyle="info")
            chk.pack(anchor=W)
            self.column_vars[col] = var


        # チェックされた列をDataFrameから削除
    def delete_selected_columns(self):
        if self.df is None:
            self.log("データが読み込まれていません")
            return
        selected = [col for col, var in self.column_vars.items() if var.get()]
        if not selected:
            self.log("削除する列が選択されていません")
            return
        self.df = self.df.drop(columns=selected)
        self.log(f"列を削除しました: {', '.join(selected)}")
        self.display_data()
        self.create_column_checkboxes()


        # CSVまたはExcelとしてDataFrameを保存（選択形式に応じて処理）
    def save_file(self):
        if self.df is None:
            self.log("保存するデータがありません")
            return
        fmt = self.save_format_var.get()
        ext = ".csv" if fmt == "CSV" else ".xlsx"
        types = [("CSVファイル", "*.csv")] if fmt == "CSV" else [("Excelファイル", "*.xlsx")]
        path = filedialog.asksaveasfilename(defaultextension=ext, filetypes=types)
        if not path:
            return
        try:
            if fmt == "CSV":
                self.df.to_csv(path, index=False)
            else:
                self.df.to_excel(path, index=False)
            self.log(f"ファイルを保存しました: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("保存エラー", str(e))
            self.log(f"保存エラー: {e}")


        # 選択されたBootstrapテーマにアプリの外観を切り替え
    def change_theme(self):
        new_theme = self.theme_var.get()
        try:
            self.root.style.theme_use(new_theme)
            self.log(f"テーマを変更しました: {new_theme}")
        except Exception as e:
            self.log(f"テーマ変更エラー: {e}")


        # ログメッセージを下部ログ表示欄に追加
    def log(self, message):
        # メッセージをログ欄に追記し、スクロールを一番下に移動
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")


if __name__ == "__main__":
    app = tb.Window(themename="journal")
    DataCleanerApp(app)
    app.mainloop()
