

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
import json
import string

class ExcelCellEditor:
    def __init__(self, master, config_path, main_app_callback):
        self.master = master
        self.config_path = config_path
        self.main_app_callback = main_app_callback # main.pyに設定更新を通知するコールバック

        self.current_config = self.load_config()
        self.original_regions_and_coords = [item.copy() for item in self.current_config.get("image_regions_and_excel_coords", [])]
        self.regions_data = [item.copy() for item in self.original_regions_and_coords] # 編集用データ

        self.cell_width = 75 # デフォルトのセル幅 (ピクセル)
        self.cell_height = 20 # デフォルトのセル高さ (ピクセル)
        self.header_offset = 30 # ラベル表示用のオフセット

        # ウィンドウの初期サイズを設定
        self._set_initial_window_size()

        self.setup_ui()
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        # ドラッグ操作用の変数
        self.drag_start_x = None
        self.drag_start_y = None
        self.selected_rect_id = None # 現在選択中の矩形ID (canvas item id)
        self.selected_region_idx = None # 現在選択中の領域データインデックス

    def load_config(self):
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            messagebox.showwarning("設定読み込み", f"設定ファイル '{self.config_path}' が見つからないか、不正です。")
            return {}

    def _set_initial_window_size(self):
        # 画面の最大サイズを取得
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()

        # 少なくとも表示したいグリッドの範囲 (例: 20列 x 30行)
        min_grid_width = 20 * self.cell_width + self.header_offset
        min_grid_height = 30 * self.cell_height + self.header_offset

        # ウィンドウの推奨サイズ
        desired_width = min(min_grid_width, screen_width * 0.9) # 画面幅の90%を上限
        desired_height = min(min_grid_height, screen_height * 0.8) # 画面高さの80%を上限

        # 最低サイズを設定
        final_width = max(800, int(desired_width))
        final_height = max(600, int(desired_height))
        
        self.master.geometry(f"{final_width}x{final_height}")

    def setup_ui(self):
        self.master.title("出力セル確認・変更")

        self.canvas = tk.Canvas(self.master, bg="white", cursor="cross")
        self.canvas.pack(fill="both", expand=True, side="left")

        # スクロールバー
        self.hbar = ttk.Scrollbar(self.master, orient="horizontal", command=self.canvas.xview)
        self.hbar.pack(side="bottom", fill="x")
        self.vbar = ttk.Scrollbar(self.master, orient="vertical", command=self.canvas.yview)
        self.vbar.pack(side="right", fill="y")
        
        # canvasx/canvasyメソッドがスクロールオフセットを返すように設定
        self.canvas.config(xscrollcommand=self.hbar.set, yscrollcommand=self.vbar.set)
        
        # スクロールイベントを直接フックしてヘッダーを再描画
        # configureイベントはウィンドウのリサイズ、MouseWheelはスクロール
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.canvas.bind("<Button-4>", self.on_scroll_event) # Linux/macOS Scroll up
        self.canvas.bind("<Button-5>", self.on_scroll_event) # Linux/macOS Scroll down
        self.canvas.bind("<MouseWheel>", self.on_scroll_event) # Windows MouseWheel

        # イベントバインディング
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)
        self.canvas.bind("<ButtonPress-3>", self.on_right_click) # 右クリック

        # ズーム機能 (Ctrl + マウスホイール)はここでは使わないが残す
        # self.canvas.bind("<Control-MouseWheel>", self.on_zoom)

        self.draw_grid_and_regions() # 初期描画

    def on_canvas_configure(self, event):
        # ウィンドウサイズ変更時に再描画
        self.draw_grid_and_regions()

    def on_scroll_event(self, event):
        # スクロール後のヘッダー再描画を遅延させる
        # `after_idle` を使うことで、Tkinterがイベントキューのアイドル状態になったときに実行される
        # これにより、連続的なスクロールイベントでの過度な再描画を防ぐ
        self.master.after_idle(self.draw_grid_and_regions)


    def on_zoom(self, event):
        # ズーム機能を実装する場合のプレースホルダー
        # 今回のタスクではズームは不要だが、将来的な拡張のために残す
        # このメソッドは現在、イベントバインディングされていない
        if event.delta > 0: # Scroll up = zoom in
            scale_factor = 1.1
        else: # Scroll down = zoom out
            scale_factor = 0.9

        # セルサイズをズーム
        self.cell_width = int(self.cell_width * scale_factor)
        self.cell_height = int(self.cell_height * scale_factor)
        
        # 最小・最大セルサイズを制限 (オプション)
        self.cell_width = max(20, min(200, self.cell_width))
        self.cell_height = max(10, min(50, self.cell_height))
        
        self.draw_grid_and_regions()


    def draw_grid_and_regions(self):
        self.canvas.delete("all") # 全てクリア

        # キャンバスの現在の表示領域 (スクロール位置) を取得
        x_view = self.canvas.canvasx(0)
        y_view = self.canvas.canvasy(0)
        
        # 表示されている範囲の計算 (header_offsetを考慮)
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()

        # --- グリッド線の描画 ---
        # 縦線
        for i in range(50): # 50列
            x = self.header_offset + i * self.cell_width
            self.canvas.create_line(x, self.header_offset, x, self.header_offset + 100 * self.cell_height, fill="lightgray", tags="grid")
            
        # 横線
        for i in range(100): # 100行
            y = self.header_offset + i * self.cell_height
            self.canvas.create_line(self.header_offset, y, self.header_offset + 50 * self.cell_width, y, fill="lightgray", tags="grid")

        # キャンバスのスクロール領域を最大グリッド範囲に合わせて設定
        max_grid_width = self.header_offset + 50 * self.cell_width
        max_grid_height = self.header_offset + 100 * self.cell_height
        self.canvas.config(scrollregion=(0, 0, max_grid_width, max_grid_height))

        # --- ヘッダーの背景描画 ---
        # コーナーの四角
        self.canvas.create_rectangle(
            x_view, y_view, x_view + self.header_offset, y_view + self.header_offset,
            fill="lightgray", outline="", tags="header_bg"
        )
        # 上部ヘッダーの背景
        self.canvas.create_rectangle(
            x_view + self.header_offset, y_view, x_view + canvas_width, y_view + self.header_offset,
            fill="lightgray", outline="", tags="header_bg"
        )
        # 左側ヘッダーの背景
        self.canvas.create_rectangle(
            x_view, y_view + self.header_offset, x_view + self.header_offset, y_view + canvas_height,
            fill="lightgray", outline="", tags="header_bg"
        )

        # --- ヘッダーラベルの描画 ---
        # 列ヘッダー
        for i in range(50): # 50列
            col_letter = self._col_to_letter(i + 1)
            x_pos = self.header_offset + i * self.cell_width
            self.canvas.create_text(
                x_pos + self.cell_width / 2, y_view + self.header_offset / 2, # x_view +... で固定位置に
                text=col_letter, fill="black", font=("Arial", 9, "bold"),
                tags="header_label"
            )
            
        # 行ヘッダー
        for i in range(100): # 100行
            y_pos = self.header_offset + i * self.cell_height
            self.canvas.create_text(
                x_view + self.header_offset / 2, y_pos + self.cell_height / 2, # y_view +... で固定位置に
                text=str(i + 1), fill="black", font=("Arial", 9, "bold"),
                tags="header_label"
            )
            
        # --- 領域の描画 ---
        for i, item in enumerate(self.regions_data):
            excel_pos = item["excel_pos"]
            x1, y1, x2, y2 = self._cell_to_coords(excel_pos)

            # 矩形を描画 (ヘッダーオフセットを考慮したCanvas座標)
            rect_id = self.canvas.create_rectangle(
                self.header_offset + x1, self.header_offset + y1,
                self.header_offset + x2, self.header_offset + y2,
                outline="blue", width=2, fill="lightblue", # 矩形内でもドラッグできるようfillを設定
                tags=("cell_rect", f"region_{i}")
            )
            # 領域番号を表示
            self.canvas.create_text(
                self.header_offset + (x1 + x2) / 2, self.header_offset + (y1 + y2) / 2,
                text=str(i + 1), fill="black", font=("Arial", 12, "bold"), tags=("cell_text", f"region_text_{i}")
            )
        
        # 矩形とテキストをグリッド線の前面に表示
        self.canvas.tag_raise("cell_rect")
        self.canvas.tag_raise("cell_text")
        
        # ヘッダー背景とラベルを最前面に配置（常に表示されるように）
        self.canvas.tag_raise("header_bg")
        self.canvas.tag_raise("header_label")


    def _col_to_letter(self, col_idx):
        # 1-based index to A, B, C...
        result = ""
        while col_idx > 0:
            col_idx, remainder = divmod(col_idx - 1, 26)
            result = string.ascii_uppercase[remainder] + result
        return result

    def _letter_to_col(self, col_str):
        # A, B, C... to 1-based index
        col_idx = 0
        for char in col_str:
            col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
        return col_idx

    def _cell_to_coords(self, excel_pos):
        # 例: "B2" -> (x1, y1, x2, y2) ピクセル座標 (ヘッダーオフセットは含まない)
        col_str = "".join(filter(str.isalpha, excel_pos)).upper()
        row_str = "".join(filter(str.isdigit, excel_pos))
        
        col_idx = self._letter_to_col(col_str)
        row_idx = int(row_str)

        x1 = (col_idx - 1) * self.cell_width
        y1 = (row_idx - 1) * self.cell_height
        x2 = x1 + self.cell_width
        y2 = y1 + self.cell_height
        return x1, y1, x2, y2

    def _coords_to_cell(self, x_coord_canvas, y_coord_canvas):
        # キャンバス座標をExcelセル座標に変換
        # ヘッダーオフセットを考慮
        col_idx = int((x_coord_canvas - self.header_offset) / self.cell_width) + 1
        row_idx = int((y_coord_canvas - self.header_offset) / self.cell_height) + 1
        
        # 負のインデックスにならないようにする (最小は1)
        col_idx = max(1, col_idx)
        row_idx = max(1, row_idx)

        col_letter = self._col_to_letter(col_idx)
        return f"{col_letter}{row_idx}"

    def on_button_press(self, event):
        self.drag_start_x = self.canvas.canvasx(event.x) # Canvas座標に変換
        self.drag_start_y = self.canvas.canvasy(event.y) # Canvas座標に変換
        self.selected_rect_id = None
        self.selected_region_idx = None

        # クリックされたアイテムのタグを確認
        item_ids = self.canvas.find_overlapping(event.x, event.y, event.x, event.y)
        
        for item_id in reversed(item_ids): # 最前面のアイテムから順に確認
            tags = self.canvas.gettags(item_id)
            if "cell_rect" in tags:
                for tag in tags:
                    if tag.startswith("region_"):
                        self.selected_region_idx = int(tag.split('_')[1])
                        self.selected_rect_id = item_id
                        self.canvas.tag_raise(self.selected_rect_id) # 選択された矩形を最前面に
                        self.canvas.tag_raise(f"region_text_{self.selected_region_idx}") # テキストも最前面に
                        self.canvas.itemconfig(self.selected_rect_id, outline="green") # 選択色
                        return # 最初の矩形が見つかったら終了
        
        # ヘッダー領域内でのクリックはドラッグしない (セルのドラッグとは区別)
        if event.x < self.header_offset or event.y < self.header_offset:
            self.drag_start_x = None # ドラッグを無効化
            self.drag_start_y = None


    def on_mouse_drag(self, event):
        if self.selected_rect_id is not None and self.selected_region_idx is not None:
            # 現在のCanvas上のマウス座標を取得 (スクロールオフセットを考慮)
            current_canvas_x = self.canvas.canvasx(event.x)
            current_canvas_y = self.canvas.canvasy(event.y)
            
            dx = current_canvas_x - self.drag_start_x
            dy = current_canvas_y - self.drag_start_y
            
            # Canvasアイテムを移動
            self.canvas.move(self.selected_rect_id, dx, dy)
            self.canvas.move(f"region_text_{self.selected_region_idx}", dx, dy)

            self.drag_start_x = current_canvas_x
            self.drag_start_y = current_canvas_y

            # 内部データの更新はon_button_releaseで行う

    def on_button_release(self, event):
        if self.selected_rect_id is not None and self.selected_region_idx is not None:
            # 矩形の現在座標を取得 (header_offset込みのcanvas座標)
            x1_canvas, y1_canvas, x2_canvas, y2_canvas = self.canvas.coords(self.selected_rect_id)

            # 新しいExcelセル座標に変換 (左上座標を使用)
            new_excel_pos = self._coords_to_cell(x1_canvas, y1_canvas)

            # 内部データを更新
            self.regions_data[self.selected_region_idx]["excel_pos"] = new_excel_pos
            messagebox.showinfo("更新", f"領域 {self.selected_region_idx + 1} のExcelセル座標を '{new_excel_pos}' に変更しました。")
            
            # 矩形の色を元に戻す
            self.canvas.itemconfig(self.selected_rect_id, outline="blue")
            
            # グリッドと領域を再描画して、矩形を新しいセル位置にスナップさせる
            self.draw_grid_and_regions()

        self.selected_rect_id = None
        self.selected_region_idx = None
        self.drag_start_x = None
        self.drag_start_y = None

    def on_right_click(self, event):
        item_id = self.canvas.find_closest(event.x, event.y)[0]
        tags = self.canvas.gettags(item_id)

        region_idx = None
        if "cell_rect" in tags:
            for tag in tags:
                if tag.startswith("region_"):
                    region_idx = int(tag.split('_')[1])
                    break

        if region_idx is not None:
            self.show_context_menu(event.x, event.y, region_idx)

    def show_context_menu(self, x, y, region_idx):
        menu = tk.Menu(self.master, tearoff=0)
        menu.add_command(label="Excelセル座標を変更", command=lambda: self.change_excel_pos_dialog(region_idx))
        menu.tk_popup(x, y)

    def change_excel_pos_dialog(self, region_idx):
        current_pos = self.regions_data[region_idx]["excel_pos"]
        new_pos = simpledialog.askstring("Excelセル座標の変更",
                                         f"領域 {region_idx + 1} のExcelセル座標を入力してください:",
                                         initialvalue=current_pos)
        if new_pos:
            # 入力されたセル座標が有効か簡易的にチェック (例: A1, B2など)
            import re
            if not re.match(r"^[A-Z]+[0-9]+$", new_pos.upper()):
                messagebox.showerror("入力エラー", "無効なExcelセル座標です。例: A1, B2")
                return

            self.regions_data[region_idx]["excel_pos"] = new_pos.upper() # 大文字に変換して保存
            messagebox.showinfo("更新", f"領域 {region_idx + 1} のExcelセル座標を '{new_pos}' に変更しました。")
            self.draw_grid_and_regions() # 更新を反映するために再描画

    def on_closing(self):
        if self.regions_data != self.original_regions_and_coords: # 変更があるか確認
            if messagebox.askyesno("確認", "変更を保存して閉じますか？"):
                self.save_config()
                self.main_app_callback() # main.pyに設定更新を通知
        self.master.destroy()

    def save_config(self):
        self.current_config["image_regions_and_excel_coords"] = self.regions_data
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.current_config, f, indent=4, ensure_ascii=False)
            messagebox.showinfo("保存", "設定ファイルに保存しました。")
        except Exception as e:
            messagebox.showerror("保存エラー", f"設定ファイルの保存中にエラーが発生しました:\n{e}")

if __name__ == '__main__':
    # 動作確認用のダミー設定と画像フォルダ
    root = tk.Tk()

    current_dir = os.path.dirname(os.path.abspath(__file__))
    dummy_config_path = os.path.join(current_dir, "config.json")

    # ダミーのコールバック関数
    def dummy_callback():
        print("Main app callback triggered (config reloaded).")

    app = ExcelCellEditor(root, dummy_config_path, dummy_callback)
    root.mainloop()

