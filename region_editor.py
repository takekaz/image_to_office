

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk # ttk追加
from PIL import Image, ImageTk
import os
import json
import copy # copyモジュールをインポート

class RegionEditor:
    def __init__(self, master, config_path, image_folder_path, main_app_callback):
        self.master = master
        self.config_path = config_path
        self.image_folder_path = image_folder_path
        self.main_app_callback = main_app_callback # main.pyに設定更新を通知するコールバック

        self.current_config = self.load_config()
        # original_regionsをdeepcopyで完全に独立させる
        self.original_regions = copy.deepcopy(self.current_config.get("image_regions_and_excel_coords", [])) # 修正
        # regions_dataもdeepcopyで完全に独立させる
        self.regions_data = copy.deepcopy(self.current_config.get("image_regions_and_excel_coords", [])) # 修正

        # 複数画像のサポートを削除し、最初の画像のみを扱う
        self.images = self.load_images()
        if not self.images:
            messagebox.showerror("エラー", f"画像フォルダ '{self.image_folder_path}' に画像が見つかりません。")
            master.destroy()
            return
        # 最初の画像のみを使用
        self.current_pil_img, self.current_image_filename = self.images[0]

        # オフセットの初期化
        self.offset_x = 0
        self.offset_y = 0
        
        # ウィンドウの初期サイズを設定
        self._set_initial_window_size()

        self.setup_ui()
        # display_imageはsetup_uiの後に呼び出し、キャンバスサイズを取得できるようにする
        # キャンバスが実際に配置されるのを待つ & サイズ確定
        self.canvas.update_idletasks()
        self.display_image()

        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        # マウスイベントの管理変数
        self.selected_region_id = None # 現在選択中の矩形ID
        self.drag_start_x = None
        self.drag_start_y = None
        self.drag_mode = None # "move", "resize_corner", "resize_line", "new_region"
        self.resize_handle_id = None # リサイズ中のハンドルのID

        self.new_region_rect_id = None # 新規領域作成中の矩形ID

        # ツールチップ関連の初期化
        self.tooltip_window = None
        self.tooltip_label = None

    def load_config(self):
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            messagebox.showwarning("設定読み込み", f"設定ファイル '{self.config_path}' が見つからないか、不正です。")
            return {}

    def load_images(self):
        supported_extensions = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff')
        image_files = sorted([f for f in os.listdir(self.image_folder_path) if f.lower().endswith(supported_extensions)])

        images = []
        if image_files: # 少なくとも1枚の画像があればそれを返す
            filename = image_files[0]
            try:
                path = os.path.join(self.image_folder_path, filename)
                img = Image.open(path)
                images.append((img, filename))
            except Exception as e:
                print(f"画像 '{filename}' の読み込み中にエラーが発生しました: {e}")
        return images

    def _set_initial_window_size(self):
        # 画像の元のサイズ
        img_width, img_height = self.current_pil_img.size

        # 画面の最大サイズを取得
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()

        # スクロールバーやウィンドウの枠、タイトルバーなどのための余白 (調整)
        # 経験的に、スクロールバーの幅や高さは約15-20px
        # ウィンドウのタイトルバーやボーダーも考慮に入れる
        # 最低限、スクロールバーが表示されるための余白を確保
        scrollbar_thickness = 20 # 仮定 (実際のOSやテーマで変動)
        padding_extra = 30 # その他のウィンドウボーダーやマージン

        # 希望するウィンドウサイズ (画像サイズ + スクロールバーのスペース + 余白)
        desired_width = img_width + scrollbar_thickness + padding_extra
        desired_height = img_height + scrollbar_thickness + padding_extra

        # ウィンドウサイズは画面の最大サイズを超えないようにする
        final_width = min(desired_width, screen_width)
        final_height = min(desired_height, screen_height)

        # 最小サイズを設定することで、スクロールバーが表示されるべき時に隠れないようにする
        # ただし、画像が画面より小さい場合は小さくても良い
        # 画像サイズが画面サイズより大きい場合にのみ、画面最大化を考慮
        if img_width > screen_width - (scrollbar_thickness + padding_extra) or \
           img_height > screen_height - (scrollbar_thickness + padding_extra):
            # 画像のいずれかの寸法が画面より大きい場合、ウィンドウを最大化する
            # 'zoomed'はWindowsで最大化、macOSでは通常全画面、Linuxでは環境依存
            self.master.state('zoomed')
        else:
            self.master.geometry(f"{final_width}x{final_height}")

        # ウィンドウの最小サイズを設定 (スクロールバーが隠れないように)
        # 例えば、画像が小さくても、スクロールバーの領域は確保したい場合
        # self.master.minsize(img_width + scrollbar_thickness + padding_extra, img_height + scrollbar_thickness + padding_extra)
        # ここでは、画像が画面より小さい場合はそのままのサイズで表示し、不要なスクロールバーを出さないため、
        # minsizeは設定しないか、より柔軟に。
        # 現在のロジックでは、desired_width/heightが最低限確保される。


    def setup_ui(self):
        self.master.title("領域確認・変更")
        self.canvas = tk.Canvas(self.master, bg="lightgray", cursor="cross")
        self.canvas.pack(fill="both", expand=True, side="left") # キャンバスを左に配置

        # スクロールバー
        self.hbar = ttk.Scrollbar(self.master, orient="horizontal", command=self.canvas.xview)
        self.hbar.pack(side="bottom", fill="x")
        self.vbar = ttk.Scrollbar(self.master, orient="vertical", command=self.canvas.yview)
        self.vbar.pack(side="right", fill="y")
        self.canvas.config(xscrollcommand=self.hbar.set, yscrollcommand=self.vbar.set)

        # イベントバインディング
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)
        self.canvas.bind("<ButtonPress-3>", self.on_right_click) # 右クリック

        # ズーム機能 (Ctrl + マウスホイール)を無効化
        # self.canvas.bind("<Control-MouseWheel>", self.on_zoom)
        # Shift + マウスドラッグでパン (スクロールバーと共存可能だが、主にズーム時に使う想定)
        self.canvas.bind("<Shift-ButtonPress-1>", self.on_pan_start)
        self.canvas.bind("<Shift-B1-Motion>", self.on_pan_drag)
        self.canvas.bind("<Shift-ButtonRelease-1>", self.on_pan_release)

    def display_image(self):
        self.canvas.delete("all")

        if not self.images:
            return

        # 画像を実寸で表示（縮小・拡大なし）
        display_img_pil = self.current_pil_img # 原寸大のPIL Imageを使用
        self.tk_img = ImageTk.PhotoImage(display_img_pil)

        # 画像の表示位置を計算 (左上基準とし、オフセットは考慮)
        self.img_x = self.offset_x
        self.img_y = self.offset_y

        self.canvas.create_image(self.img_x, self.img_y, image=self.tk_img, anchor="nw", tags="current_image")

        # スクロール領域を画像の実寸サイズに合わせて設定
        self.canvas.config(scrollregion=(0, 0, self.current_pil_img.width, self.current_pil_img.height))

        self.draw_regions()

    def draw_regions(self):
        self.canvas.delete("region_rect") # 既存の領域を削除
        self.canvas.delete("handle")      # 既存のハンドルを削除
        self.canvas.delete("region_text_bg") # 既存のテキスト背景を削除
        self.canvas.delete("region_text") # 既存のテキストを削除

        # 領域の描画は画像ピクセル座標とCanvas座標を一致させる（オフセットのみ考慮）
        for i, item in enumerate(self.regions_data):
            img_region = item["img_region"]

            # 画像座標をキャンバス座標に変換 (オフセットのみ考慮)
            x1_canvas = self.img_x + img_region[0]
            y1_canvas = self.img_y + img_region[1]
            x2_canvas = self.img_x + img_region[2]
            y2_canvas = self.img_y + img_region[3]

            # 矩形を描画
            self.canvas.create_rectangle(
                x1_canvas, y1_canvas, x2_canvas, y2_canvas,
                outline="blue", width=2, tags=("region_rect", f"region_{i}")
            )
            # ハンドルとテキストは別の関数で描画
            self.draw_handles_and_text_for_region(i)


    def draw_handles_and_text_for_region(self, i):
        item = self.regions_data[i]
        img_region = item["img_region"]

        # 画像座標をキャンバス座標に変換 (オフセットのみ考慮)
        x1_canvas = self.img_x + img_region[0]
        y1_canvas = self.img_y + img_region[1]
        x2_canvas = self.img_x + img_region[2]
        y2_canvas = self.img_y + img_region[3]

        center_x = (x1_canvas + x2_canvas) / 2
        center_y = (y1_canvas + y2_canvas) / 2

        # 領域番号のテキストと背景のサイズを計算
        text_content = str(i + 1)
        font_size = 12
        # フォントメトリクスを正確に取得するのはTkinterでは少し手間だが、ここでは固定値を仮定
        # 大体の文字幅と高さを予測
        # 例: 1文字あたり約10-12px幅、15-18px高さ
        text_width_approx = len(text_content) * 8 + 4 # 文字数*平均文字幅 + 左右のパディング
        text_height_approx = font_size + 4 # フォントサイズ + 上下のパディング

        # テキスト背景用の矩形を描画
        self.canvas.create_rectangle(
            center_x - text_width_approx / 2,
            center_y - text_height_approx / 2,
            center_x + text_width_approx / 2,
            center_y + text_height_approx / 2,
            fill="black",
            outline="black", # 枠線も黒で背景と一体化
            tags=("region_rect", "region_text_bg", f"region_text_bg_{i}")
        )

        # 領域番号を表示
        self.canvas.create_text(
            center_x, center_y, text=text_content, fill="white",
            font=("Arial", font_size, "bold"), tags=("region_rect", f"region_text_{i}")
        )

        # リサイズハンドルを描画
        handle_size = 8
        # Corners
        self.canvas.create_rectangle(x1_canvas - handle_size/2, y1_canvas - handle_size/2,
                                    x1_canvas + handle_size/2, y1_canvas + handle_size/2,
                                    fill="red", outline="white", tags=("handle", f"handle_corner_{i}_nw"))
        self.canvas.create_rectangle(x2_canvas - handle_size/2, y1_canvas - handle_size/2,
                                    x2_canvas + handle_size/2, y1_canvas + handle_size/2,
                                    fill="red", outline="white", tags=("handle", f"handle_corner_{i}_ne"))
        self.canvas.create_rectangle(x1_canvas - handle_size/2, y2_canvas - handle_size/2,
                                    x1_canvas + handle_size/2, y2_canvas + handle_size/2,
                                    fill="red", outline="white", tags=("handle", f"handle_corner_{i}_sw"))
        self.canvas.create_rectangle(x2_canvas - handle_size/2, y2_canvas - handle_size/2,
                                    x2_canvas + handle_size/2, y2_canvas + handle_size/2,
                                    fill="red", outline="white", tags=("handle", f"handle_corner_{i}_se"))
        # Midpoints for lines
        self.canvas.create_rectangle(x1_canvas - handle_size/2, center_y - handle_size/2,
                                    x1_canvas + handle_size/2, center_y + handle_size/2,
                                    fill="green", outline="white", tags=("handle", f"handle_line_{i}_w"))
        self.canvas.create_rectangle(x2_canvas - handle_size/2, center_y - handle_size/2,
                                    x2_canvas + handle_size/2, center_y + handle_size/2,
                                    fill="green", outline="white", tags=("handle", f"handle_line_{i}_e"))
        self.canvas.create_rectangle(center_x - handle_size/2, y1_canvas - handle_size/2,
                                    center_x + handle_size/2, y1_canvas + handle_size/2,
                                    fill="green", outline="white", tags=("handle", f"handle_line_{i}_n"))
        self.canvas.create_rectangle(center_x - handle_size/2, y2_canvas - handle_size/2,
                                    center_x + handle_size/2, y2_canvas + handle_size/2,
                                    fill="green", outline="white", tags=("handle", f"handle_line_{i}_s"))


    # on_zoomメソッドは無効化 (削除済み)

    def on_pan_start(self, event):
        self.pan_start_x = self.canvas.canvasx(event.x) # キャンバス座標に変換
        self.pan_start_y = self.canvas.canvasy(event.y) # キャンバス座標に変換
        self.canvas.config(cursor="fleur")
        # パン開始時にツールチップ表示
        self.show_tooltip(event.x, event.y, "パン中")

    def on_pan_drag(self, event):
        # 現在のキャンバスビューのオフセットを計算
        new_x = self.canvas.canvasx(event.x)
        new_y = self.canvas.canvasy(event.y)
        self.canvas.xview_scroll(int(self.pan_start_x - new_x), "units")
        self.canvas.yview_scroll(int(self.pan_start_y - new_y), "units")

        # ツールチップの更新
        self.update_tooltip(event.x, event.y, "パン中")

    def on_pan_release(self, event):
        self.canvas.config(cursor="cross")
        self.hide_tooltip() # パン終了時にツールチップ非表示

    def show_tooltip(self, x, y, text):
        if self.tooltip_window:
            self.hide_tooltip()

        self.tooltip_window = tk.Toplevel(self.master)
        self.tooltip_window.wm_overrideredirect(True) # ウィンドウの装飾をなくす
        # ルートウィンドウの座標を取得して相対位置を計算
        master_x = self.master.winfo_x()
        master_y = self.master.winfo_y()
        self.tooltip_window.wm_geometry(f"+{master_x + x + 10}+{master_y + y + 10}")

        self.tooltip_label = tk.Label(self.tooltip_window, text=text, background="lightyellow", relief="solid", borderwidth=1,
                                     font=("Arial", 9, "normal"))
        self.tooltip_label.pack(padx=1, pady=1)

    def update_tooltip(self, x, y, text):
        if self.tooltip_window:
            self.tooltip_label.config(text=text)
            master_x = self.master.winfo_x()
            master_y = self.master.winfo_y()
            self.tooltip_window.wm_geometry(f"+{master_x + x + 10}+{master_y + y + 10}")

    def hide_tooltip(self):
        if self.tooltip_window:
            self.tooltip_window.destroy()
        self.tooltip_window = None
        self.tooltip_label = None

    def _format_coords_for_tooltip(self, coords):
        return f" ({int(coords[0])},{int(coords[1])})-({int(coords[2])},{int(coords[3])})"

    def on_button_press(self, event):
        self.drag_start_x = self.canvas.canvasx(event.x) # キャンバス座標に変換
        self.drag_start_y = self.canvas.canvasy(event.y) # キャンバス座標に変換
        self.selected_region_id = None # クリック時に一旦選択を解除
        self.hide_tooltip() # 新しい操作開始時にツールチップを非表示にする

        # クリックされたアイテムのタグを確認
        item_id = self.canvas.find_closest(self.drag_start_x, self.drag_start_y)[0]
        tags = self.canvas.gettags(item_id)

        found_region_tag = None
        for tag in tags:
            # region_{i} の形式のタグを探す。ただし "region_rect" や "region_text_{i}" は除外する
            if tag.startswith("region_") and tag != "region_rect" and not tag.startswith("region_text_"):
                found_region_tag = tag
                break

        if found_region_tag:
            self.selected_region_id = int(found_region_tag.split('_')[1])
            self.drag_mode = "move"
            # 枠のクリックで緑色に
            self.canvas.itemconfig(f"region_{self.selected_region_id}", outline="green")
            # 移動開始時にツールチップを表示
            current_region_coords = self.regions_data[self.selected_region_id]["img_region"]
            self.show_tooltip(event.x, event.y, "移動中" + self._format_coords_for_tooltip(current_region_coords))
        elif "handle" in tags:
            # ハンドルがクリックされた場合 (リサイズまたは色変更)
            for tag in tags:
                if tag.startswith("handle_corner_") or tag.startswith("handle_line_"):
                    parts = tag.split('_')
                    self.selected_region_id = int(parts[2]) # region index
                    self.resize_handle_id = tag # 例: handle_corner_0_nw
                    self.drag_mode = "resize"
                    # 角のクリックで赤色に
                    self.canvas.itemconfig(f"region_{self.selected_region_id}", outline="red")
                    # リサイズ開始時にツールチップを表示
                    current_region_coords = self.regions_data[self.selected_region_id]["img_region"]
                    self.show_tooltip(event.x, event.y, "リサイズ中" + self._format_coords_for_tooltip(current_region_coords))
                    break
        else:
            # 何もクリックされていない場合 (新規領域作成の可能性)
            self.drag_mode = "new_region_potential"
            self.new_region_rect_id = None # 初期化
            # ここではツールチップを表示しない。ドラッグが20px以上になったら表示

    def on_mouse_drag(self, event):
        current_canvas_x = self.canvas.canvasx(event.x)
        current_canvas_y = self.canvas.canvasy(event.y)
        tooltip_text = ""
        current_region_coords = None

        if self.selected_region_id is not None and self.drag_mode == "move":
            # 領域の移動
            dx = current_canvas_x - self.drag_start_x
            dy = current_canvas_y - self.drag_start_y

            # Canvas上の矩形、テキスト、ハンドルを移動
            self.canvas.move(f"region_{self.selected_region_id}", dx, dy)
            self.canvas.move(f"region_text_bg_{self.selected_region_id}", dx, dy) # 背景も移動
            self.canvas.move(f"region_text_{self.selected_region_id}", dx, dy)
            self.canvas.move(f"handle_corner_{self.selected_region_id}_nw", dx, dy)
            self.canvas.move(f"handle_corner_{self.selected_region_id}_ne", dx, dy)
            self.canvas.move(f"handle_corner_{self.selected_region_id}_sw", dx, dy)
            self.canvas.move(f"handle_corner_{self.selected_region_id}_se", dx, dy)
            self.canvas.move(f"handle_line_{self.selected_region_id}_w", dx, dy)
            self.canvas.move(f"handle_line_{self.selected_region_id}_e", dx, dy)
            self.canvas.move(f"handle_line_{self.selected_region_id}_n", dx, dy)
            self.canvas.move(f"handle_line_{self.selected_region_id}_s", dx, dy)

            # 内部データを更新 (画像ピクセル座標はCanvas座標と一致するので、そのままdx, dyを加算)
            region_item = self.regions_data[self.selected_region_id]
            region_item["img_region"][0] += dx
            region_item["img_region"][1] += dy
            region_item["img_region"][2] += dx
            region_item["img_region"][3] += dy

            self.drag_start_x = current_canvas_x # drag_startを更新して次の移動量計算に備える
            self.drag_start_y = current_canvas_y

            tooltip_text = "移動中"
            current_region_coords = region_item["img_region"]


        elif self.selected_region_id is not None and self.drag_mode == "resize":
            # 領域のリサイズ
            current_rect_coords_canvas = list(self.canvas.coords(f"region_{self.selected_region_id}"))
            handle_tag = self.resize_handle_id

            if handle_tag.endswith("_nw"):
                current_rect_coords_canvas[0], current_rect_coords_canvas[1] = current_canvas_x, current_canvas_y
            elif handle_tag.endswith("_ne"):
                current_rect_coords_canvas[2], current_rect_coords_canvas[1] = current_canvas_x, current_canvas_y
            elif handle_tag.endswith("_sw"):
                current_rect_coords_canvas[0], current_rect_coords_canvas[3] = current_canvas_x, current_canvas_y
            elif handle_tag.endswith("_se"):
                current_rect_coords_canvas[2], current_rect_coords_canvas[3] = current_canvas_x, current_canvas_y
            elif handle_tag.endswith("_w"):
                current_rect_coords_canvas[0] = current_canvas_x
            elif handle_tag.endswith("_e"):
                current_rect_coords_canvas[2] = current_canvas_x
            elif handle_tag.endswith("_n"):
                current_rect_coords_canvas[1] = current_canvas_y
            elif handle_tag.endswith("_s"):
                current_rect_coords_canvas[3] = current_canvas_y

            self.canvas.coords(f"region_{self.selected_region_id}", *current_rect_coords_canvas)

            # リサイズ中はテキストとハンドルを再描画
            self.canvas.delete(f"region_text_bg_{self.selected_region_id}") # 背景も削除
            self.canvas.delete(f"region_text_{self.selected_region_id}")
            self.canvas.delete("handle")
            self.draw_handles_and_text_for_region(self.selected_region_id)

            # 画像座標はCanvas座標と一致するので、そのまま取得
            x1_img = (current_rect_coords_canvas[0] - self.img_x)
            y1_img = (current_rect_coords_canvas[1] - self.img_y)
            x2_img = (current_rect_coords_canvas[2] - self.img_x)
            y2_img = (current_rect_coords_canvas[3] - self.img_y)
            current_region_coords = [min(x1_img, x2_img), min(y1_img, y2_img), max(x1_img, x2_img), max(y1_img, y2_img)]

            tooltip_text = "リサイズ中"

        elif self.drag_mode == "new_region_potential":
            dx = current_canvas_x - self.drag_start_x
            dy = current_canvas_y - self.drag_start_y

            if abs(dx) > 20 or abs(dy) > 20:
                self.drag_mode = "new_region" # モードを確定
                # ここで初めて仮の矩形を生成
                self.new_region_rect_id = self.canvas.create_rectangle(
                    self.drag_start_x, self.drag_start_y, current_canvas_x, current_canvas_y,
                    outline="orange", width=2, tags="new_region_temp"
                )
                tooltip_text = "新規追加中"
                # 新規領域の画像座標 (Canvas座標と一致)
                x1_img = (self.drag_start_x - self.img_x)
                y1_img = (self.drag_start_y - self.img_y)
                x2_img = (current_canvas_x - self.img_x)
                y2_img = (current_canvas_y - self.img_y)
                current_region_coords = [min(x1_img, x2_img), min(y1_img, y2_img), max(x1_img, x2_img), max(y1_img, y2_img)]
                # 新規追加の初回ドラッグでツールチップを表示
                self.show_tooltip(event.x, event.y, tooltip_text + self._format_coords_for_tooltip(current_region_coords))
            else:
                self.hide_tooltip() # 20ピクセル未満ではツールチップを非表示
                return # 20ピクセル未満では何もしない
        
        elif self.drag_mode == "new_region":
            # 仮の矩形を更新
            self.canvas.coords(self.new_region_rect_id, self.drag_start_x, self.drag_start_y, current_canvas_x, current_canvas_y)
            tooltip_text = "新規追加中"
            # 新規領域の画像座標 (Canvas座標と一致)
            x1_img = (self.drag_start_x - self.img_x)
            y1_img = (self.drag_start_y - self.img_y)
            x2_img = (current_canvas_x - self.img_x)
            y2_img = (current_canvas_y - self.img_y)
            current_region_coords = [min(x1_img, x2_img), min(y1_img, y2_img), max(x1_img, x2_img), max(y1_img, y2_img)]
            # 新規追加の継続ドラッグでツールチップを更新
            self.update_tooltip(event.x, event.y, tooltip_text + self._format_coords_for_tooltip(current_region_coords))
            return # このパスでは下部のif文は実行しない

        if tooltip_text and current_region_coords:
            self.update_tooltip(event.x, event.y, tooltip_text + self._format_coords_for_tooltip(current_region_coords))


    def on_button_release(self, event):
        self.hide_tooltip() # リリース時にツールチップを非表示にする

        if self.drag_mode == "resize" and self.selected_region_id is not None:
            # リサイズ後の内部データを更新
            current_coords_canvas = self.canvas.coords(f"region_{self.selected_region_id}")
            # キャンバス座標から画像ピクセルに逆変換（ここではそのまま）
            x1_img = (current_coords_canvas[0] - self.img_x)
            y1_img = (current_coords_canvas[1] - self.img_y)
            x2_img = (current_coords_canvas[2] - self.img_x)
            y2_img = (current_coords_canvas[3] - self.img_y)

            # 負のサイズにならないように正規化し、整数に変換
            self.regions_data[self.selected_region_id]["img_region"] = [
                int(min(x1_img, x2_img)), int(min(y1_img, y2_img)),
                int(max(x1_img, x2_img)), int(max(y1_img, y2_img))
            ]
            self.draw_regions() # ハンドルとテキストの位置を更新するために再描画

        elif self.drag_mode == "new_region":
            # 新規領域の確定
            if self.new_region_rect_id:
                coords = self.canvas.coords(self.new_region_rect_id)
                self.canvas.delete(self.new_region_rect_id) # 仮の矩形を削除

                # キャンバス座標から画像ピクセルに逆変換（ここではそのまま）
                x1_img = (coords[0] - self.img_x)
                y1_img = (coords[1] - self.img_y)
                x2_img = (coords[2] - self.img_x)
                y2_img = (coords[3] - self.img_y)

                # サイズが0でないことを確認し、整数に変換
                if abs(x2_img - x1_img) > 1 and abs(y2_img - y1_img) > 1:
                    new_region = {
                        "img_region": [int(min(x1_img, x2_img)), int(min(y1_img, y2_img)),
                                       int(max(x1_img, x2_img)), int(max(y1_img, y2_img))],
                        "excel_pos": "A1" # デフォルト値を設定、後で変更可能にする
                    }
                    self.regions_data.append(new_region)
                    self.draw_regions() # 新しい領域を描画
                else:
                    messagebox.showwarning("新規領域作成", "領域のサイズが小さすぎます。20ピクセル以上のドラッグが必要です。")


        self.drag_mode = None
        self.resize_handle_id = None
        self.new_region_rect_id = None
        self.selected_region_id = None # ドラッグ終了時に選択を解除し、色をリセット

        # 全ての矩形の色を青に戻す (リリース時に選択状態を解除)
        for i in range(len(self.regions_data)):
            self.canvas.itemconfig(f"region_{i}", outline="blue")


    def on_right_click(self, event):
        item_id = self.canvas.find_closest(event.x, event.y)[0]
        tags = self.canvas.gettags(item_id)

        region_idx = None
        # クリックされたのが領域自体またはハンドルの場合、その領域のインデックスを取得
        for tag in tags:
            if tag.startswith("region_") and tag != "region_rect" and not tag.startswith("region_text_bg_") and not tag.startswith("region_text_"):
                # region_X のX部分を取得
                try:
                    region_idx = int(tag.split('_')[1])
                except ValueError:
                    continue
                break
            elif tag.startswith("handle_corner_") or tag.startswith("handle_line_"):
                # handle_type_X_direction のX部分を取得
                try:
                    region_idx = int(tag.split('_')[2])
                except ValueError:
                    continue
                break

        if region_idx is not None:
            self.selected_region_id = region_idx # 右クリックされた領域を選択状態にする
            self.show_context_menu(event.x, event.y, region_idx)
        else:
            # どこでもない場所で右クリックした場合は、メニューを表示しない
            pass

    def show_context_menu(self, x, y, region_idx):
        menu = tk.Menu(self.master, tearoff=0)
        menu.add_command(label="領域を削除", command=lambda: self.delete_region(region_idx))
        menu.add_command(label="Excelセル座標を変更", command=lambda: self.change_excel_pos(region_idx))
        menu.tk_popup(x, y)

    def delete_region(self, region_idx):
        if messagebox.askyesno("確認", f"領域 {region_idx + 1} を削除しますか？"):
            del self.regions_data[region_idx]
            self.draw_regions() # 領域を再描画して更新を反映

    def change_excel_pos(self, region_idx):
        current_pos = self.regions_data[region_idx]["excel_pos"]
        new_pos = simpledialog.askstring("Excelセル座標の変更",
                                         f"領域 {region_idx + 1} のExcelセル座標を入力してください:",
                                         initialvalue=current_pos)
        if new_pos:
            self.regions_data[region_idx]["excel_pos"] = new_pos
            messagebox.showinfo("更新", f"領域 {region_idx + 1} のExcelセル座標を '{new_pos}' に変更しました。")
            # 画面には直接影響しないが、内部データが更新されたことを確認

    def on_closing(self):
        print("--- on_closing called ---")
        print(f"regions_data (current): {self.regions_data}")
        print(f"original_regions (initial): {self.original_regions}")
        print(f"Are they different? {self.regions_data != self.original_regions}")

        if self.regions_data != self.original_regions: # 変更があるか確認
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
            error_message = f"設定ファイルの保存中にエラーが発生しました:\n{e}"
            print(error_message) # コンソールにも出力
            messagebox.showerror("保存エラー", error_message)

if __name__ == '__main__':
    # 動作確認用のダミー設定と画像フォルダ
    # 本番環境では main.py から呼び出される
    root = tk.Tk()
    root.geometry("1000x800")

    current_dir = os.path.dirname(os.path.abspath(__file__))
    dummy_config_path = os.path.join(current_dir, "config.json")
    dummy_image_folder = os.path.join(current_dir, "img")

    # ダミーのコールバック関数
    def dummy_callback():
        print("Main app callback triggered (config reloaded).")

    app = RegionEditor(root, dummy_config_path, dummy_image_folder, dummy_callback)
    root.mainloop()


