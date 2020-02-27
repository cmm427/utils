# -*- coding: UTF-8 -*-

from pathlib import Path
from PIL import Image, ImageDraw, ImageFont


'''
生成测试照片
'''


def create_image(width=1920, height=1080, ext="jpg"):
    img = Image.new('RGB', (width, height), color=(73, 109, 137))
    d = ImageDraw.Draw(img)
    fnt = ImageFont.truetype('C:\Windows\Fonts\Arial.ttf', size=42)
    text_str = "width: " + str(width) + "\nheight: " + str(height) + "\nformat: " + ext
    d.text((width/2, height/2), text_str, font=fnt, fill=(255, 255, 0))
    file_name = '{0}{1}{2}{3}{4}'.format(str(width), '_', str(height), '.', ext)
    dir_name = Path.cwd() / "images"
    if not dir_name.exists():
        Path.mkdir(dir_name)
    img.save(str(dir_name / file_name))
    print(file_name + " is created")


def main():
    widths = [1280, 1920, 2688]
    heights = [720, 1080, 1242, 1280]
    ext_lst = ['jpg', 'jpeg', 'png']
    for w in widths:
        for h in heights:
            for ext in ext_lst:
                create_image(w, h, ext)


if __name__ == "__main__":
    main()
