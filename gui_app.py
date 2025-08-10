
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import json
from excel_image_embedder import embed_images_to_excel
from ppt_image_embedder import embed_images_to_ppt

class ImageToOfficeApp:
    def __init__(self, master):
        self.master = master
        master.title("Image to Office Converter")
        master.geometry("800x600")

        self.config_data = {}
        self.load_config()

        self.create_widgets()
        self.set_initial_paths()

    def load_config(self):
        script_dir = os.path.dirname(__file__)
        config_path = os.path.join(script_dir, "config.json")
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                self.config_data = json.load(f)
            print(f"Config loaded from {config_path}")
        except FileNotFoundError:
            messagebox.showerror("Error", f"設定ファイルが見つかりません: {config_path}")
            self.config_data = self.create_default_config(config_path)
        except json.JSONDecodeError:
            messagebox.showerror("Error", f"設定ファイルの読み込みに失敗しました: {config_path}")
            self.config_data = self.create_default_config(config_path)


    def create_default_config(self, config_path):
        """デフォルトの設定ファイルを作成し、保存します。"""
        default_config = {
            "image_regions_and_coords": [
                {"img_region": [100, 100, 200, 150], "excel_pos": "B2"}
            ],
            "excel_to_ppt_conversion_params": {
                "col_width_pixels_per_char": 7.0,
                "default_col_width_chars": 8.43,
                "row_height_points_per_row": 15.0,
                "dpi": 96
            }
        }
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=4, ensure_ascii=False)
            messagebox.showinfo("情報", f"設定ファイルが見つからなかったため、デフォルト設定で作成しました: {config_path}")
            return default_config
        except Exception as e:
            messagebox.showerror("エラー", f"デフォルト設定ファイルの作成に失敗しました: {e}")
            return {}

    def set_initial_paths(self):
        script_dir = os.path.dirname(__file__)
        self.image_folder_var.set(os.path.join(script_dir, "img"))
        self.excel_output_path_var.set(os.path.join(script_dir, "output_images.xlsx"))
        self.ppt_output_path_var.set(os.path.join(script_dir, "output_images.pptx"))

    def create_widgets(self):
        # Frame for paths
        path_frame = tk.LabelFrame(self.master, text="ファイルパス設定", padx=10, pady=10)
        path_frame.pack(padx=10, pady=10, fill="x")

        # Image Folder
        tk.Label(path_frame, text="画像フォルダ:").grid(row=0, column=0, sticky="w", pady=2)
        self.image_folder_var = tk.StringVar()
        self.image_folder_entry = tk.Entry(path_frame, textvariable=self.image_folder_var, width=60)
        self.image_folder_entry.grid(row=0, column=1, padx=5, pady=2)
        tk.Button(path_frame, text="選択", command=self.select_image_folder).grid(row=0, column=2, padx=5, pady=2)

        # Excel Output Path
        tk.Label(path_frame, text="Excel出力パス:").grid(row=1, column=0, sticky="w", pady=2)
        self.excel_output_path_var = tk.StringVar()
        self.excel_output_path_entry = tk.Entry(path_frame, textvariable=self.excel_output_path_var, width=60)
        self.excel_output_path_entry.grid(row=1, column=1, padx=5, pady=2)
        tk.Button(path_frame, text="選択", command=self.select_excel_output_path).grid(row=1, column=2, padx=5, pady=2)

        # PowerPoint Output Path
        tk.Label(path_frame, text="PowerPoint出力パス:").grid(row=2, column=0, sticky="w", pady=2)
        self.ppt_output_path_var = tk.StringVar()
        self.ppt_output_path_entry = tk.Entry(path_frame, textvariable=self.ppt_output_path_var, width=60)
        self.ppt_output_path_entry.grid(row=2, column=1, padx=5, pady=2)
        tk.Button(path_frame, text="選択", command=self.select_ppt_output_path).grid(row=2, column=2, padx=5, pady=2)

        # Frame for region/cell display (placeholder)
        display_frame = tk.LabelFrame(self.master, text="領域・セル座標設定", padx=10, pady=10)
        display_frame.pack(padx=10, pady=10, fill="both", expand=True)

        tk.Label(display_frame, text="画像領域とセル座標のペア:").pack(anchor="w", pady=2)
        self.regions_text = tk.Text(display_frame, height=10, wrap="word")
        self.regions_text.pack(fill="both", expand=True)
        self.update_regions_display()

        # Action Buttons for region/cell
        action_buttons_frame = tk.Frame(self.master)
        action_buttons_frame.pack(pady=5)
        tk.Button(action_buttons_frame, text="領域確認・変更", command=self.confirm_change_regions).pack(side="left", padx=5)
        tk.Button(action_buttons_frame, text="出力セル確認・変更", command=self.confirm_change_output_cells).pack(side="left", padx=5)
        
        # Output Buttons
        output_buttons_frame = tk.Frame(self.master)
        output_buttons_frame.pack(pady=10)
        tk.Button(output_buttons_frame, text="Excel出力", command=self.run_excel_embedder, height=2, width=15).pack(side="left", padx=10)
        tk.Button(output_buttons_frame, text="PowerPoint出力", command=self.run_ppt_embedder, height=2, width=15).pack(side="left", padx=10)

    def update_regions_display(self):
        self.regions_text.delete(1.0, tk.END)
        regions_data = self.config_data.get("image_regions_and_coords", [])
        if regions_data:
            for item in regions_data:
                self.regions_text.insert(tk.END, f"  Img Region: {item.get('img_region')}, Excel Pos: {item.get('excel_pos')}\n")
        else:
            self.regions_text.insert(tk.END, "設定ファイルに領域データが見つかりません。\n")

    def select_image_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.image_folder_var.set(folder_selected)

    def select_excel_output_path(self):
        file_selected = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if file_selected:
            self.excel_output_path_var.set(file_selected)

    def select_ppt_output_path(self):
        file_selected = filedialog.asksaveasfilename(defaultextension=".pptx",
                                                     filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")])
        if file_selected:
            self.ppt_output_path_var.set(file_selected)

    def confirm_change_regions(self):
        messagebox.showinfo("機能未実装", "領域確認・変更機能はまだ実装されていません。")

    def confirm_change_output_cells(self):
        messagebox.showinfo("機能未実装", "出力セル確認・変更機能はまだ実装されていません。")

    def run_excel_embedder(self):
        excel_path = self.excel_output_path_var.get()
        image_folder = self.image_folder_var.get()
        regions_data = self.config_data.get("image_regions_and_coords", [])

        if not regions_data:
            messagebox.showwarning("警告", "設定ファイルに埋め込みたい領域データがありません。")
            return
        
        if not os.path.exists(image_folder):
            messagebox.showerror("エラー", f"画像フォルダが見つかりません: {image_folder}")
            return

        try:
            embed_images_to_excel(excel_path, image_folder, regions_data)
            messagebox.showinfo("完了", f"Excelファイル '{os.path.basename(excel_path)}' が正常に作成されました。")
        except Exception as e:
            messagebox.showerror("エラー", f"Excelファイルの作成中にエラーが発生しました: {e}")

    def run_ppt_embedder(self):
        ppt_path = self.ppt_output_path_var.get()
        image_folder = self.image_folder_var.get()
        regions_data = self.config_data.get("image_regions_and_coords", [])
        excel_conversion_params = self.config_data.get("excel_to_ppt_conversion_params", {})

        if not regions_data:
            messagebox.showwarning("警告", "設定ファイルに埋め込みたい領域データがありません。")
            return
        
        if not os.path.exists(image_folder):
            messagebox.showerror("エラー", f"画像フォルダが見つかりません: {image_folder}")
            return

        if not excel_conversion_params:
            messagebox.showwarning("警告", "設定ファイルにPowerPoint変換パラメータがありません。")
            return

        try:
            embed_images_to_ppt(ppt_path, image_folder, regions_data, excel_conversion_params)
            messagebox.showinfo("完了", f"PowerPointファイル '{os.path.basename(ppt_path)}' が正常に作成されました。")
        except Exception as e:
            messagebox.showerror("エラー", f"PowerPointファイルの作成中にエラーが発生しました: {e}")

if __name__ == '__main__':
    root = tk.Tk()
    app = ImageToOfficeApp(root)
    root.mainloop()

