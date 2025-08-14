

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import json
import functools
import re
import sys

# 既存の処理関数をインポート
from image_to_excel import insert_images_to_excel
from image_to_pptx import insert_images_to_pptx

class ImageToOfficeApp:
    def __init__(self, master):
        self.master = master
        master.title("Image to Office Converter")
        master.geometry("800x600")

        if getattr(sys, 'frozen', False):
            self.app_exe_dir = os.path.dirname(sys.executable)
            self.resource_base_dir = self.app_exe_dir
        else:
            self.app_exe_dir = os.path.dirname(os.path.abspath(__file__))
            self.resource_base_dir = self.app_exe_dir

        self.config_path = os.path.join(self.app_exe_dir, "config.json")

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
        # 画像フォルダ (exe/スクリプトパス内のimgフォルダを使用)
        tk.Label(input_frame, text="画像フォルダ:").grid(row=0, column=0, sticky="w", pady=2)
        self.image_folder_var = tk.StringVar(value=os.path.join(self.app_exe_dir, "img"))
        tk.Entry(input_frame, textvariable=self.image_folder_var, width=60).grid(row=0, column=1, padx=5, pady=2)
        tk.Button(input_frame, text="選択", command=functools.partial(self.browse_folder, self.image_folder_var)).grid(row=0, column=2, pady=2)

        # Excel出力パス (exe/スクリプトパスを使用)
        tk.Label(input_frame, text="Excel出力パス:").grid(row=1, column=0, sticky="w", pady=2)
        self.excel_output_path_var = tk.StringVar(value=os.path.join(self.app_exe_dir, "output_images.xlsx"))
        tk.Entry(input_frame, textvariable=self.excel_output_path_var, width=60).grid(row=1, column=1, padx=5, pady=2)
        tk.Button(input_frame, text="選択", command=functools.partial(self.browse_file, self.excel_output_path_var, [("Excel files", "*.xlsx")])).grid(row=1, column=2, pady=2)

        # PowerPoint出力パス (exe/スクリプトパスを使用)
        tk.Label(input_frame, text="PowerPoint出力パス:").grid(row=2, column=0, sticky="w", pady=2)
        self.pptx_output_path_var = tk.StringVar(value=os.path.join(self.app_exe_dir, "output_images.pptx"))
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
        self.tree.pack_forget()
        self.tree.pack(fill="both", expand=True, padx=5, pady=5, side="left")

        # 設定変更関連のボタン
        settings_buttons_frame = tk.Frame(settings_frame) # 新しいフレームにボタンをまとめる
        settings_buttons_frame.pack(side="right", fill="y", padx=5, pady=5)

        tk.Button(settings_buttons_frame, text="領域確認・変更", command=self.open_region_editor).pack(pady=5)
        tk.Button(settings_buttons_frame, text="出力セル確認・変更", command=self.open_excel_cell_editor).pack(pady=5)
        # 選択した領域を削除ボタンを追加
        tk.Button(settings_buttons_frame, text="選択した領域を削除", command=self.delete_selected_region).pack(pady=5) # 追加

        # --- 出力ボタンのセクション ---
        tk.Button(buttons_frame, text="Excel出力", command=self.run_excel_export).pack(side="left", padx=10, pady=5)
        tk.Button(buttons_frame, text="PowerPoint出力", command=self.run_pptx_export).pack(side="left", padx=10, pady=5)

    def browse_folder(self, var):
        initial_dir = os.path.dirname(var.get()) if os.path.exists(var.get()) else self.app_exe_dir
        folder_selected = filedialog.askdirectory(initialdir=initial_dir)
        if folder_selected:
            var.set(folder_selected)

    def browse_file(self, var, filetypes):
        current_path = var.get()
        initial_dir = os.path.dirname(current_path) if os.path.exists(os.path.dirname(current_path)) else self.app_exe_dir
        initial_file = os.path.basename(current_path)

        file_selected = filedialog.asksaveasfilename(
            initialdir=initial_dir,
            initialfile=initial_file,
            defaultextension=filetypes[0][1].split('*')[-1],
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
            self._save_config() # 初期設定をファイルに保存
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
            self._save_config() # 不正な場合も初期設定をファイルに保存

    def _save_config(self): # ヘルパー関数名を変更 (initial_configでなく汎用的に)
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            print(f"設定ファイルが保存されました: {self.config_path}")
        except Exception as e:
            print(f"設定ファイルの保存中にエラーが発生しました:\n{e}")
            messagebox.showerror("保存エラー", f"設定ファイルの保存中にエラーが発生しました:\n{e}") # UIにも表示

    def update_config_display(self):
        # Treeviewの既存データをクリア
        for i in self.tree.get_children():
            self.tree.delete(i)

        # 新しいデータを挿入
        regions_and_coords = self.config.get("image_regions_and_excel_coords", [])
        for item in regions_and_coords:
            self.tree.insert("", "end", values=(str(item.get("img_region")), item.get("excel_pos")))

    def delete_selected_region(self): # 新規追加
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("削除", "削除する領域を選択してください。")
            return

        if messagebox.askyesno("確認", "選択された領域を削除しますか？"):
            # 選択されたアイテムのインデックスを特定し、configから削除
            # Treeviewの表示順とconfigのリストの順序が一致している前提
            # 逆順で削除することで、インデックスのずれを防ぐ
            selected_indices = sorted([self.tree.index(item) for item in selected_items], reverse=True)
            for idx in selected_indices:
                if 0 <= idx < len(self.config["image_regions_and_excel_coords"]):
                    del self.config["image_regions_and_excel_coords"][idx]
            
            self._save_config() # 変更を保存
            self.update_config_display() # Treeviewを更新
            messagebox.showinfo("削除", "選択された領域を削除しました。")


    def open_region_editor(self):
        editor_window = tk.Toplevel(self.master)
        from region_editor import RegionEditor
        RegionEditor(editor_window, self.config_path, self.image_folder_var.get(), self.update_main_config_display)

    def update_main_config_display(self):
        self.load_config()
        messagebox.showinfo("設定更新", "画像領域設定が更新されました。")

    def open_excel_cell_editor(self):
        editor_window = tk.Toplevel(self.master)
        from excel_cell_editor import ExcelCellEditor
        ExcelCellEditor(editor_window, self.config_path, self.update_main_config_display)

    def _handle_file_overwrite(self, original_filepath):
        if not os.path.exists(original_filepath):
            return original_filepath

        result = messagebox.askyesnocancel(
            "ファイルが存在します",
            f"出力ファイル '{os.path.basename(original_filepath)}' は既に存在します。\n\n"
            "上書きしますか？\n"
            "「はい」：ファイルを上書きします。\n"
            "「いいえ」：ファイル名に連番を付加して保存します。\n"
            "「キャンセル」：処理を中止します。"
        )

        if result is True:
            return original_filepath
        elif result is False:
            base, ext = os.path.splitext(original_filepath)
            match = re.search(r'\((\d+)\)$', base)
            if match:
                base = base[:match.start()]
                start_num = int(match.group(1)) + 1
            else:
                start_num = 1

            i = start_num
            new_filepath = f"{base}({i}){ext}"
            while os.path.exists(new_filepath):
                i += 1
                new_filepath = f"{base}({i}){ext}"
            return new_filepath
        else:
            return None

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

        excel_output_path = self._handle_file_overwrite(original_excel_output_path)
        if excel_output_path is None:
            messagebox.showinfo("処理中止", "Excelファイルの出力がキャンセルされました。")
            return

        try:
            insert_images_to_excel(excel_output_path, image_folder, regions_and_coords)
            messagebox.showinfo("成功", f"Excelファイルが正常に生成されました:\n{excel_output_path}")
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

        pptx_output_path = self._handle_file_overwrite(original_pptx_output_path)
        if pptx_output_path is None:
            messagebox.showinfo("処理中止", "PowerPointファイルの出力がキャンセルされました。")
            return

        try:
            template_file = self.config.get("pptx_template_file", None)
            template_pptx_path = os.path.join(self.app_exe_dir, template_file) if template_file else None
            insert_images_to_pptx(pptx_output_path, image_folder, regions_and_coords, excel_conv_params, template_pptx_path)
            messagebox.showinfo("成功", f"PowerPointファイルが正常に生成されました:\n{pptx_output_path}")
            if pptx_output_path != original_pptx_output_path:
                self.pptx_output_path_var.set(pptx_output_path)
        except Exception as e:
            messagebox.showerror("エラー", f"PowerPointファイルの生成中にエラーが発生しました:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ImageToOfficeApp(root)
    root.mainloop()

