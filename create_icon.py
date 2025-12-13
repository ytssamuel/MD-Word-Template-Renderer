"""
生成 MD-Word Renderer 應用程式圖標
現代簡約科技風格設計
"""
from PIL import Image, ImageDraw, ImageFont
import os
import math


def create_icon():
    """創建應用程式圖標 - 現代科技風格"""
    
    # 創建多個尺寸的圖標
    sizes = [16, 32, 48, 64, 128, 256]
    images = []
    
    for size in sizes:
        # 創建帶透明背景的圖像
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        
        # 計算比例
        scale = size / 256
        
        # === 設計：簡約漸層背景 + 抽象轉換符號 ===
        
        # 背景：圓角矩形，帶現代漸層
        padding = int(8 * scale)
        radius = int(40 * scale)
        
        # 繪製漸層背景（深藍到青色，科技感）
        for y in range(padding, size - padding):
            # 計算漸層進度
            progress = (y - padding) / (size - 2 * padding)
            
            # 從深藍 #1a1a2e 漸變到 #16213e 再到 #0f3460
            r = int(26 + (15 - 26) * progress)
            g = int(26 + (51 - 26) * progress)
            b = int(46 + (96 - 46) * progress)
            
            draw.line([(padding, y), (size - padding - 1, y)], fill=(r, g, b, 255))
        
        # 繪製圓角遮罩
        mask = Image.new('L', (size, size), 0)
        mask_draw = ImageDraw.Draw(mask)
        mask_draw.rounded_rectangle(
            [padding, padding, size - padding - 1, size - padding - 1],
            radius=radius,
            fill=255
        )
        
        # 應用遮罩
        img.putalpha(mask)
        draw = ImageDraw.Draw(img)
        
        # === 核心圖案：抽象的文件轉換符號 ===
        center_x = size // 2
        center_y = size // 2
        
        # 左側：代表 Markdown 的簡約符號（三條橫線）
        line_width = max(2, int(4 * scale))
        line_length = int(45 * scale)
        line_gap = int(16 * scale)
        left_x = center_x - int(50 * scale)
        
        # 三條漸層線（代表文字/Markdown）
        for i, offset in enumerate([-line_gap, 0, line_gap]):
            # 線條長度遞減，更有層次感
            current_length = line_length - i * int(8 * scale)
            y = center_y + offset
            
            # 青色漸層 #00d9ff → #00fff2
            alpha = 255 - i * 40
            draw.rounded_rectangle(
                [left_x, y - line_width//2, 
                 left_x + current_length, y + line_width//2],
                radius=line_width//2,
                fill=(0, 217, 255, alpha)
            )
        
        # 中間：流動的轉換箭頭
        arrow_start = center_x - int(5 * scale)
        arrow_end = center_x + int(35 * scale)
        arrow_y = center_y
        
        # 箭頭主體（漸層效果）
        arrow_width = max(2, int(3 * scale))
        draw.line(
            [(arrow_start, arrow_y), (arrow_end - int(10*scale), arrow_y)],
            fill=(0, 255, 242, 255),
            width=arrow_width
        )
        
        # 箭頭頭部（三角形）
        arrow_head_size = max(4, int(12 * scale))
        draw.polygon([
            (arrow_end, arrow_y),
            (arrow_end - arrow_head_size, arrow_y - arrow_head_size//2),
            (arrow_end - arrow_head_size, arrow_y + arrow_head_size//2)
        ], fill=(0, 255, 242, 255))
        
        # 右側：代表 Word 文件的簡約符號（文件圖示）
        doc_left = center_x + int(40 * scale)
        doc_width = int(50 * scale)
        doc_height = int(60 * scale)
        doc_top = center_y - doc_height // 2
        fold_size = int(12 * scale)
        
        # 文件主體（帶折角）
        doc_points = [
            (doc_left, doc_top + fold_size),  # 左上（折角下方）
            (doc_left, doc_top + doc_height),  # 左下
            (doc_left + doc_width, doc_top + doc_height),  # 右下
            (doc_left + doc_width, doc_top),  # 右上
            (doc_left + fold_size, doc_top),  # 折角右
            (doc_left, doc_top + fold_size),  # 回到左上
        ]
        draw.polygon(doc_points, fill=(255, 255, 255, 230))
        
        # 折角效果
        fold_points = [
            (doc_left, doc_top + fold_size),
            (doc_left + fold_size, doc_top + fold_size),
            (doc_left + fold_size, doc_top),
        ]
        draw.polygon(fold_points, fill=(200, 200, 200, 200))
        
        # 文件內的線條（代表內容）
        content_x = doc_left + int(8 * scale)
        content_width = doc_width - int(16 * scale)
        for i, offset in enumerate([int(20*scale), int(32*scale), int(44*scale)]):
            line_w = content_width - i * int(6 * scale)
            draw.rounded_rectangle(
                [content_x, doc_top + offset,
                 content_x + line_w, doc_top + offset + int(4*scale)],
                radius=int(2*scale),
                fill=(100, 100, 100, 150)
            )
        
        # === 添加微妙的光暈效果（大尺寸才加）===
        if size >= 64:
            # 在箭頭附近添加光暈
            glow_radius = int(8 * scale)
            for r in range(glow_radius, 0, -1):
                alpha = int(30 * (r / glow_radius))
                # 這裡簡化，不加複雜光暈
        
        images.append(img)
    
    # 保存為 .ico 文件（包含多個尺寸）
    ico_path = os.path.join(os.path.dirname(__file__), 'assets', 'app_icon.ico')
    os.makedirs(os.path.dirname(ico_path), exist_ok=True)
    
    # 保存 ICO（從大到小排列）
    images[-1].save(
        ico_path,
        format='ICO',
        sizes=[(s, s) for s in sizes],
        append_images=images[:-1]
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
