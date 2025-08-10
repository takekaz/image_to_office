


import os
import json
import tempfile # 追加
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

def excel_coord_to_inches(excel_pos: str, params: dict):
    """
    Excelのセル座標をPowerPointのインチ座標に変換します。
    この関数は仮の実装であり、実際のExcelの列幅や行高さに基づいて
    より複雑な計算が必要になる場合があります。

    Args:
        excel_pos (str): Excelのセル座標 (例: "B2")
        params (dict): 変換に必要なパラメータ (例: {"col_width_pix": 64, "row_height_pix": 20, "dpi": 96})

    Returns:
        tuple: (x_inches, y_inches) - PowerPoint上でのx, y座標 (インチ単位)
    """
    col_width_pix = params.get("col_width_pix", 64) # Excelのデフォルト列幅（ピクセル）
    row_height_pix = params.get("row_height_pix", 20) # Excelのデフォルト行高さ（ピクセル）
    dpi = params.get("dpi", 96) # スクリーンDPI

    # Excel座標を列と行のインデックスに変換
    col_str = "".join(filter(str.isalpha, excel_pos)).upper()
    row_str = "".join(filter(str.isdigit, excel_pos))

    col_idx = 0
    for char in col_str:
        col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
    row_idx = int(row_str)

    # ピクセル単位での概算位置 (左上セルA1を(0,0)とする)
    # これは非常に単純な仮の計算であり、実際のExcelのセルサイズは複雑です
    # 列幅や行高さはopenpyxlで取得可能ですが、ここでは簡略化
    x_pix = (col_idx - 1) * col_width_pix
    y_pix = (row_idx - 1) * row_height_pix

    # ピクセルをインチに変換
    x_inches = x_pix / dpi
    y_inches = y_pix / dpi

    return Inches(x_inches), Inches(y_inches)

def insert_images_to_pptx(pptx_filepath: str, image_folder_path: str, regions_and_coords: list, excel_conv_params: dict):
    """
    指定された画像フォルダ内の画像を読み込み、その領域をPowerPointスライドの指定座標に貼り付けます。
    画像ごとに新しいスライドを作成し、スライド右上に画像ファイル名を表記します。

    Args:
        pptx_filepath (str): 出力するPowerPointファイルのパス。
        image_folder_path (str): 画像が保存されているフォルダのパス。
        regions_and_coords (list): 領域とセル座標のペアのリスト。
                                   例: [{"img_region": [x1, y1, x2, y2], "excel_pos": "B2"}, ...]
                                   img_region: [left, upper, right, lower] (Pillowのcrop形式)
        excel_conv_params (dict): Excelのセル座標をインチに変換するためのパラメータ。
                                  例: {"col_width_pix": 64, "row_height_pix": 20, "dpi": 96}
    """
    # 既存のファイルがあっても上書きで新規作成
    prs = Presentation()

    # レイアウトの選択 (ここでは空白のスライドレイアウトを使用)
    blank_slide_layout = prs.slide_layouts[6] # 通常、6番目が空白レイアウト

    supported_extensions = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff')

    # 画像フォルダ内の全ての画像ファイルを処理
    for image_filename in sorted(os.listdir(image_folder_path)):
        if image_filename.lower().endswith(supported_extensions):
            image_path = os.path.join(image_folder_path, image_filename)

            # 画像ごとに新しいスライドを作成
            slide = prs.slides.add_slide(blank_slide_layout)
            print(f"Processing image: {image_filename} on new slide")

            # スライド右上に画像ファイル名を表記
            # テキストボックスのサイズと位置を調整
            left = Inches(prs.slide_width.inches - 2) # スライド右端から2インチ左
            top = Inches(0.1) # スライド上端から0.1インチ下
            width = Inches(1.9)
            height = Inches(0.5)
            textbox = slide.shapes.add_textbox(left, top, width, height)
            text_frame = textbox.text_frame
            text_frame.text = image_filename
            text_frame.word_wrap = True

            # フォントサイズを調整
            p = text_frame.paragraphs[0]
            p.font.size = Pt(10)

            try:
                original_image = Image.open(image_path)

                for item in regions_and_coords:
                    img_region = item["img_region"]
                    excel_pos = item["excel_pos"]

                    # 画像領域を切り抜き
                    cropped_image = original_image.crop(img_region)

                    # 切り抜いた画像を一時ファイルとして保存
                    # tempfileモジュールを使用して一時ファイルを安全に作成
                    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
                        temp_image_path = temp_file.name
                        cropped_image.save(temp_image_path)

                    # Excelセル座標をPowerPointのインチ座標に変換
                    x_inches, y_inches = excel_coord_to_inches(excel_pos, excel_conv_params)

                    # 画像をスライドに貼り付け
                    # widthとheightを元の切り抜き画像のサイズに基づいて自動調整させるために指定しない
                    pic = slide.shapes.add_picture(temp_image_path, x_inches, y_inches)
                    print(f"  - Cropped region {img_region} from {image_filename} and inserted at {excel_pos} ({x_inches.inches:.2f}in, {y_inches.inches:.2f}in)")

                    # 一時ファイルを削除 (NamedTemporaryFileでdelete=Falseにしたので手動で削除)
                    os.remove(temp_image_path)

            except Exception as e:
                print(f"Error processing {image_filename}: {e}")
                continue

    # プレゼンテーションを保存
    try:
        prs.save(pptx_filepath)
        print(f"PowerPoint file saved successfully to {pptx_filepath}")
    except Exception as e:
        print(f"Error saving PowerPoint file: {e}")
    finally:
        # tempfileを使用しているので、個別にリスト管理して削除する必要がなくなりました
        pass

if __name__ == '__main__':
    current_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(current_dir, "config.json")
    image_dir = os.path.join(current_dir, "img")
    output_pptx_file = os.path.join(current_dir, "output_images.pptx")

    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)

        regions_and_coords = config.get("image_regions_and_excel_coords", [])
        excel_conversion_parameters = config.get("excel_to_pptx_conversion_params", {})

        # 関数を実行
        insert_images_to_pptx(output_pptx_file, image_dir, regions_and_coords, excel_conversion_parameters)

    except FileNotFoundError:
        print(f"Error: config file not found at {config_path}")
    except json.JSONDecodeError:
        print(f"Error: Could not decode JSON from {config_path}. Check file format.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


