"""
生成 MD-Word Renderer 應用程式圖標
結合 Markdown 和 Word 的視覺元素
"""
from PIL import Image, ImageDraw, ImageFont
import os


def create_icon():
    """創建應用程式圖標"""
    
    # 創建多個尺寸的圖標
    sizes = [16, 32, 48, 64, 128, 256]
    images = []
    
    for size in sizes:
        # 創建帶透明背景的圖像
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        
        # 計算比例
        scale = size / 256
        
        # === 背景：漸層藍色圓角矩形（代表文檔） ===
        padding = int(20 * scale)
        radius = int(30 * scale)
        
        # 繪製圓角矩形背景（深藍到淺藍漸層效果）
        for i in range(size - 2 * padding):
            alpha = int(255 * (1 - i / (size - 2 * padding) * 0.3))
            color = (
                int(41 + (100 - 41) * i / (size - 2 * padding)),  # R: 深藍到藍
                int(98 + (149 - 98) * i / (size - 2 * padding)),   # G: 深藍到藍
                int(255),  # B: 保持最大值
                alpha
            )
            draw.rectangle(
                [padding, padding + i, size - padding - 1, padding + i + 1],
                fill=color
            )
        
        # 繪製白色邊框
        border_width = max(2, int(3 * scale))
        draw.rounded_rectangle(
            [padding, padding, size - padding - 1, size - padding - 1],
            radius=radius,
            outline=(255, 255, 255, 255),
            width=border_width
        )
        
        # === 左側：Markdown 符號 (MD) ===
        # MD 文字或 # 符號
        if size >= 64:
            try:
                # 嘗試使用系統字體
                font_size = int(size * 0.35)
                try:
                    font = ImageFont.truetype("arial.ttf", font_size)
                except:
                    try:
                        font = ImageFont.truetype("segoeui.ttf", font_size)
                    except:
                        font = ImageFont.load_default()
                
                # 左側繪製 "MD"
                text = "MD"
                bbox = draw.textbbox((0, 0), text, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                
                text_x = padding + int((size - 2 * padding) * 0.25) - text_width // 2
                text_y = (size - text_height) // 2 - int(5 * scale)
                
                # 添加文字陰影
                draw.text((text_x + 2, text_y + 2), text, fill=(0, 0, 0, 100), font=font)
                draw.text((text_x, text_y), text, fill=(255, 255, 255, 255), font=font)
                
            except Exception as e:
                print(f"警告: 無法載入字體 (size {size}): {e}")
                # 使用簡單的幾何圖形代替
                pass
        
        # === 中間：箭頭 ===
        arrow_y = size // 2
        arrow_start_x = int(size * 0.42)
        arrow_end_x = int(size * 0.58)
        arrow_width = max(2, int(4 * scale))
        arrow_head = max(3, int(8 * scale))
        
        # 箭頭線條
        draw.line(
            [(arrow_start_x, arrow_y), (arrow_end_x, arrow_y)],
            fill=(255, 255, 255, 255),
            width=arrow_width
        )
        
        # 箭頭頭部
        draw.polygon(
            [
                (arrow_end_x, arrow_y),
                (arrow_end_x - arrow_head, arrow_y - arrow_head // 2),
                (arrow_end_x - arrow_head, arrow_y + arrow_head // 2)
            ],
            fill=(255, 255, 255, 255)
        )
        
        # === 右側：Word 文檔符號 (W) ===
        if size >= 64:
            try:
                # 繪製 "W"
                text = "W"
                bbox = draw.textbbox((0, 0), text, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                
                text_x = padding + int((size - 2 * padding) * 0.75) - text_width // 2
                text_y = (size - text_height) // 2 - int(5 * scale)
                
                # 添加文字陰影
                draw.text((text_x + 2, text_y + 2), text, fill=(0, 0, 0, 100), font=font)
                draw.text((text_x, text_y), text, fill=(255, 255, 255, 255), font=font)
                
            except Exception as e:
                print(f"警告: 無法載入字體 (size {size}): {e}")
        
        images.append(img)
    
    # 保存為 .ico 文件（包含多個尺寸）
    ico_path = os.path.join(os.path.dirname(__file__), 'assets', 'app_icon.ico')
    os.makedirs(os.path.dirname(ico_path), exist_ok=True)
    
    images[0].save(
        ico_path,
        format='ICO',
        sizes=[(s, s) for s in sizes]
    )
    
    print(f"✅ 圖標已創建: {ico_path}")
    
    # 也保存 PNG 版本（用於文檔等）
    png_path = os.path.join(os.path.dirname(__file__), 'assets', 'app_icon.png')
    images[-1].save(png_path, format='PNG')
    print(f"✅ PNG 圖標已創建: {png_path}")
    
    return ico_path, png_path


if __name__ == '__main__':
    try:
        ico_path, png_path = create_icon()
        print("\n圖標生成成功！")
        print(f"ICO: {ico_path}")
        print(f"PNG: {png_path}")
    except Exception as e:
        print(f"❌ 錯誤: {e}")
        import traceback
        traceback.print_exc()
