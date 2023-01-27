'''Убираем фон картинки методом прозрачности не нужных пикселей и усиливаем насыщенность нужных'''
from PIL import Image
from HSV.hsv import rgb_to_hsv, hsv_to_rgb

def GO(fail):
    img = Image.open(fail)
    img = img.convert("RGBA")
    datas = img.getdata()
    ONE_255 = 1.0 / 255.0
    xxx = 222
    newData = []

    for item in datas:
        r, g, b = (item[0], item[1], item[2])
        h, s, v = rgb_to_hsv(r * ONE_255, g * ONE_255, b * ONE_255)
        h, s, v = (h * 360.0, s * 100.0, v * 100.0)

        if r > xxx and g > xxx and b > xxx:
            newData.append((0, 0, 0, 0))
        else:
            s = s + 30 if s + 30 < 100 else s
            if h < 200 or h > 250:
                newData.append((0, 0, 0, 0))
            else:
                r, g, b = hsv_to_rgb(h / 360, s / 100, v / 100)
                r, g, b = (r * 255.0, g * 255.0, b * 255.0)
                newData.append((int(r), int(g), int(b), 255))

    fl = fail[:-3] + ".png"
    img.putdata(newData)
    img.save(rf"{fl}", "PNG")
    return fl


if __name__ == "__main__":
    fail = r"C:\vxvproj\tnnc-Excel\CorrectReportApp\CorrectReport\Подписи\Шукалова.jpg"
    GO(fail)