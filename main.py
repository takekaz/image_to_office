

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import json
import functools
import re # 連番付加のために追加

# 既存の処理関数をインポート
from image_to_excel import insert_images_to_excel
from image_to_pptx import insert_images_to_pptx

class ImageToOfficeApp:
    def __init__(self, master):
        self.master = master
        master.title("Image to Office Converter")
        master.geometry("800x600")

        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_path = os.path.join(self.base_dir, "config.json")

        self.create_widgets()
        self.load_config()

    def create_widgets(self):
        # フレームの作成
        input_frame = tk.LabelFrame(self.master, text="入力パスと出力パス", padx=10, pady=10)
        input_frame.pack(padx=10, pady=5, fill="x")

        settings_frame = tk.LabelFrame(self.master, text="画像領域とExcel/PowerPoint設定", padx=10, pady=10)
        settings_frame.pack(padx=10, pady=5, fill="both", expand=True)

        buttons_frame = tk.Frame(self.master, padx=10, pady=10)
        buttons_frame.pack(padx=10, pady=5, fill="x")

        # --- 入力パスと出力パスのセクション ---
        # 画像フォルダ
        tk.Label(input_frame, text="画像フォルダ:").grid(row=0, column=0, sticky="w", pady=2)
        self.image_folder_var = tk.StringVar(value=os.path.join(self.base_dir, "img"))
        tk.Entry(input_frame, textvariable=self.image_folder_var, width=60).grid(row=0, column=1, padx=5, pady=2)
        tk.Button(input_frame, text="選択", command=functools.partial(self.browse_folder, self.image_folder_var)).grid(row=0, column=2, pady=2)

        # Excel出力パス
        tk.Label(input_frame, text="Excel出力パス:").grid(row=1, column=0, sticky="w", pady=2)
        self.excel_output_path_var = tk.StringVar(value=os.path.join(self.base_dir, "output_images.xlsx"))
        tk.Entry(input_frame, textvariable=self.excel_output_path_var, width=60).grid(row=1, column=1, padx=5, pady=2)
        tk.Button(input_frame, text="選択", command=functools.partial(self.browse_file, self.excel_output_path_var, [("Excel files", "*.xlsx")])).grid(row=1, column=2, pady=2)

        # PowerPoint出力パス
        tk.Label(input_frame, text="PowerPoint出力パス:").grid(row=2, column=0, sticky="w", pady=2)
        self.pptx_output_path_var = tk.StringVar(value=os.path.join(self.base_dir, "output_images.pptx"))
        tk.Entry(input_frame, textvariable=self.pptx_output_path_var, width=60).grid(row=2, column=1, padx=5, pady=2)
        tk.Button(input_frame, text="選択", command=functools.partial(self.browse_file, self.pptx_output_path_var, [("PowerPoint files", "*.pptx")])).grid(row=2, column=2, pady=2)

        # --- 設定表示と変更ボタンのセクション ---
        # Treeviewで領域とセル座標のペアを表示
        self.tree = ttk.Treeview(settings_frame, columns=("img_region", "excel_pos"), show="headings")
        self.tree.heading("img_region", text="画像領域 [x1, y1, x2, y2]")
        self.tree.heading("excel_pos", text="Excelセル座標")
        self.tree.column("img_region", width=200)
        self.tree.column("excel_pos", width=100)
        self.tree.pack(fill="both", expand=True, padx=5, pady=5)

        # スクロールバー
        scrollbar = ttk.Scrollbar(settings_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tree.pack_forget() # パックし直すために一度非表示
        self.tree.pack(fill="both", expand=True, padx=5, pady=5, side="left") # Treeviewを左側に配置

        # 領域確認・変更ボタン
        tk.Button(settings_frame, text="領域確認・変更", command=self.open_region_editor).pack(pady=5)
        # 出力セル確認・変更ボタン
        tk.Button(settings_frame, text="出力セル確認・変更", command=self.open_excel_cell_editor).pack(pady=5)

        # --- 出力ボタンのセクション ---
        tk.Button(buttons_frame, text="Excel出力", command=self.run_excel_export).pack(side="left", padx=10, pady=5)
        tk.Button(buttons_frame, text="PowerPoint出力", command=self.run_pptx_export).pack(side="left", padx=10, pady=5)

    def browse_folder(self, var):
        folder_selected = filedialog.askdirectory(initialdir=os.path.dirname(var.get()))
        if folder_selected:
            var.set(folder_selected)

    def browse_file(self, var, filetypes):
        file_selected = filedialog.asksaveasfilename(
            initialdir=os.path.dirname(var.get()),
            initialfile=os.path.basename(var.get()),
            defaultextension=filetypes[0][1].split('*')[-1], # '.xlsx' など
            filetypes=filetypes
        )
        if file_selected:
            var.set(file_selected)

    def load_config(self):
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.config = json.load(f)
            self.update_config_display()
            messagebox.showinfo("設定読み込み", "設定ファイルを読み込みました。")
        except FileNotFoundError:
            messagebox.showwarning("設定読み込み", f"設定ファイルが見つかりません: {self.config_path}\nデフォルト値で作成します。")
            self.config = {
                "image_regions_and_excel_coords": [
                    {"img_region": [100, 100, 200, 150], "excel_pos": "B2"}
                ],
                "excel_to_pptx_conversion_params": {
                    "col_width_pix": 64, "row_height_pix": 20, "dpi": 96
                }
            }
            self._save_initial_config() # 初期設定をファイルに保存
        except json.JSONDecodeError:
            messagebox.showerror("設定読み込みエラー", f"設定ファイルのJSON形式が不正です: {self.config_path}")
            self.config = {
                "image_regions_and_excel_coords": [
                    {"img_region": [100, 100, 200, 150], "excel_pos": "B2"}
                ],
                "excel_to_pptx_conversion_params": {
                    "col_width_pix": 64, "row_height_pix": 20, "dpi": 96
                }
            }
            self._save_initial_config() # 不正な場合も初期設定をファイルに保存

    def _save_initial_config(self):
        # 設定ファイルを保存するヘルパー関数
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            print(f"初期設定ファイルが作成されました: {self.config_path}")
        except Exception as e:
            print(f"初期設定ファイルの保存中にエラーが発生しました: {e}")

    def update_config_display(self):
        # Treeviewの既存データをクリア
        for i in self.tree.get_children():
            self.tree.delete(i)

        # 新しいデータを挿入
        regions_and_coords = self.config.get("image_regions_and_excel_coords", [])
        for item in regions_and_coords:
            self.tree.insert("", "end", values=(str(item.get("img_region")), item.get("excel_pos")))

    def open_region_editor(self):
        # 新しいトップレベルウィンドウを作成
        editor_window = tk.Toplevel(self.master)
        # region_editor.py の RegionEditor クラスをインスタンス化
        # self.update_main_config_display をコールバックとして渡す
        from region_editor import RegionEditor # ここでインポート
        RegionEditor(editor_window, self.config_path, self.image_folder_var.get(), self.update_main_config_display)

    def update_main_config_display(self):
        # region_editorで設定が保存された後にmain.pyの表示を更新
        self.load_config() # config.jsonを再読み込み
        messagebox.showinfo("設定更新", "画像領域設定が更新されました。")

    def open_excel_cell_editor(self):
        editor_window = tk.Toplevel(self.master)
        # excel_cell_editor.py の ExcelCellEditor クラスをインスタンス化
        from excel_cell_editor import ExcelCellEditor # ここでインポート
        ExcelCellEditor(editor_window, self.config_path, self.update_main_config_display)

    def _handle_file_overwrite(self, original_filepath):
        """
        ファイルが存在する場合、上書きまたは連番付加をユーザーに確認する。
        最終的な出力ファイルパスを返す。
        """
        if not os.path.exists(original_filepath):
            return original_filepath # ファイルが存在しない場合はそのまま返す

        # ファイルが存在する場合
        result = messagebox.askyesnocancel(
            "ファイルが存在します",
            f"出力ファイル '{os.path.basename(original_filepath)}' は既に存在します。\n\n"
            "上書きしますか？\n"
            "「はい」：ファイルを上書きします。\n"
            "「いいえ」：ファイル名に連番を付加して保存します。\n"
            "「キャンセル」：処理を中止します。"
        )

        if result is True: # はい (上書き)
            return original_filepath
        elif result is False: # いいえ (連番付加)
            base, ext = os.path.splitext(original_filepath)
            # 既存の連番 (例: (1)) があればそれを考慮して次の連番を探す
            match = re.search(r'\((\d+)\)$', base)
            if match:
                base = base[:match.start()] # 連番部分を除去
                start_num = int(match.group(1)) + 1
            else:
                start_num = 1

            i = start_num
            new_filepath = f"{base}({i}){ext}"
            while os.path.exists(new_filepath):
                i += 1
                new_filepath = f"{base}({i}){ext}"
            return new_filepath
        else: # キャンセル
            return None # 処理を中止することを示す

    def run_excel_export(self):
        image_folder = self.image_folder_var.get()
        original_excel_output_path = self.excel_output_path_var.get()
        regions_and_coords = self.config.get("image_regions_and_excel_coords", [])

        if not os.path.isdir(image_folder):
            messagebox.showerror("エラー", f"画像フォルダが見つかりません: {image_folder}")
            return
        if not regions_and_coords:
            messagebox.showwarning("警告", "設定ファイルに画像領域とセル座標のペアが定義されていません。")
            return

        # ファイルの上書き確認とパスの決定
        excel_output_path = self._handle_file_overwrite(original_excel_output_path)
        if excel_output_path is None: # ユーザーがキャンセルした場合
            messagebox.showinfo("処理中止", "Excelファイルの出力がキャンセルされました。")
            return

        try:
            insert_images_to_excel(excel_output_path, image_folder, regions_and_coords)
            messagebox.showinfo("成功", f"Excelファイルが正常に生成されました:\n{excel_output_path}")
            # もしファイル名が変更された場合、UIのパスも更新する
            if excel_output_path != original_excel_output_path:
                self.excel_output_path_var.set(excel_output_path)
        except Exception as e:
            messagebox.showerror("エラー", f"Excelファイルの生成中にエラーが発生しました:\n{e}")

    def run_pptx_export(self):
        image_folder = self.image_folder_var.get()
        original_pptx_output_path = self.pptx_output_path_var.get()
        regions_and_coords = self.config.get("image_regions_and_excel_coords", [])
        excel_conv_params = self.config.get("excel_to_pptx_conversion_params", {})

        if not os.path.isdir(image_folder):
            messagebox.showerror("エラー", f"画像フォルダが見つかりません: {image_folder}")
            return
        if not regions_and_coords:
            messagebox.showwarning("警告", "設定ファイルに画像領域とセル座標のペアが定義されていません。")
            return
        if not excel_conv_params:
            messagebox.showwarning("警告", "設定ファイルにExcel-PowerPoint変換パラメータが定義されていません。")
            return

        # ファイルの上書き確認とパスの決定
        pptx_output_path = self._handle_file_overwrite(original_pptx_output_path)
        if pptx_output_path is None: # ユーザーがキャンセルした場合
            messagebox.showinfo("処理中止", "PowerPointファイルの出力がキャンセルされました。")
            return

        try:
            insert_images_to_pptx(pptx_output_path, image_folder, regions_and_coords, excel_conv_params)
            messagebox.showinfo("成功", f"PowerPointファイルが正常に生成されました:\n{pptx_output_path}")
            # もしファイル名が変更された場合、UIのパスも更新する
            if pptx_output_path != original_pptx_output_path:
                self.pptx_output_path_var.set(pptx_output_path)
        except Exception as e:
            messagebox.showerror("エラー", f"PowerPointファイルの生成中にエラーが発生しました:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ImageToOfficeApp(root)
    root.mainloop()

