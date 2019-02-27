import pyautogui
import PIL
print(dir(PIL))
from PIL import ImageGrab
bbox = (760, 0, 1160, 1080)

print(pyautogui.position())

im1 = ImageGrab.grab()
pyautogui.moveTo(654, 112)
pyautogui.dragTo(834, 357)
im2 = ImageGrab.grab()
#  from (654, 112) to (834, 357)

print('im',dir(im2))
# print('im1',im1.Info)
# print('im2',im2.Info)
im1.save('pirture1.png')
im2.save('pirture2.png')

# from PIL import Image
# from PIL import ImageChops
#
#
# def compare_images(path_one, path_two, diff_save_location):
#     """
#     比较图片，如果有不同则生成展示不同的图片
#
#     @参数一: path_one: 第一张图片的路径
#     @参数二: path_two: 第二张图片的路径
#     @参数三: diff_save_location: 不同图的保存路径
#     """
#     image_one = Image.open('pirture1.png')
#     image_two = Image.open('pirture2.png')
#     try:
#         diff = ImageChops.difference(image_one, image_two)
#
#         if diff.getbbox() is None:
#             # 图片间没有任何不同则直接退出
#             print("【+】We are the same!")
#         else:
#             diff.save(diff_save_location)
#     except ValueError as e:
#         text = ("表示图片大小和box对应的宽度不一致，参考API说明：Pastes another image into this image."
#                 "The box argument is either a 2-tuple giving the upper left corner, a 4-tuple defining the left, upper, "
#                 "right, and lower pixel coordinate, or None (same as (0, 0)). If a 4-tuple is given, the size of the pasted "
#                 "image must match the size of the region.使用2纬的box避免上述问题")
#         print("【{0}】{1}".format(e, text))
#
#
# if __name__ == '__main__':
#     compare_images('1.png',
#                    '2.png',                   '我们不一样.png')