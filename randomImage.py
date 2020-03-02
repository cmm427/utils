# -*- coding: UTF-8 -*-

from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from random import *


# 获取文件存放目录
def get_dir():
    dir_name = Path.cwd() / "images"
    if not dir_name.exists():
        Path.mkdir(dir_name)

    return dir_name


# 生成测试照片
def create_image(width=1920, height=1080, ext="jpg"):
    img = Image.new('RGB', (width, height), color=(73, 109, 137))
    d = ImageDraw.Draw(img)

    fnt = ImageFont.truetype('C:\Windows\Fonts\Arial.ttf', size=42)
    text_str = "width: " + str(width) + "\nheight: " + str(height) + "\nformat: " + ext
    d.text((width/2, height/2), text_str, font=fnt, fill=(255, 255, 0))

    file_name = '{0}{1}{2}{3}{4}'.format(str(width), '_', str(height), '.', ext)

    img.save(str(get_dir() / file_name))
    print(file_name + " is created")


# 生成GIF
def create_gif(width, height):
    names = ['img{:02d}.gif'.format(i) for i in range(20)]

    im = Image.new("RGB", (width, height), color=(73, 109, 137))

    pos = randrange(0, width/2)
    for n in names:
        frame = im.copy()
        draw = ImageDraw.Draw(frame)
        draw.ellipse((pos, pos, 50+pos, 50+pos), 'red')
        frame.save(str(get_dir() / n))

        pos += 10

    images = []

    for n in names:
        frame = Image.open(str(get_dir() / n))
        images.append(frame)

    # 保存所用的frame为GIF格式动画
    file_name = str(width)+"_"+str(height)+'.gif'
    images[0].save(str(get_dir() / file_name), save_all=True, append_images=images[1:], duration=100, loop=0)
    print(file_name + " is created")


def main():
    resolutions = [(720, 1280), (1080, 1920), (1242, 2688)]
    ext_lst = ['jpg', 'png']

    # 生成图片
    for r in resolutions:
        for fmt in ext_lst:
            create_image(r[0], r[1], fmt)

    # 生成GIF
    for r in resolutions:
        create_gif(r[0], r[1])


if __name__ == "__main__":
    main()
