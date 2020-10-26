# -*- coding: utf-8 -*-
import PIL.Image as PILI


def cutPngIntoGrid(fn, rowNum, colNum, keepList=None):
    #
    """
    fn: png文件名
    png等距切割成rowNum行 * colNum列的小png，每个png按从左到右、从上到下，自0编号。
    keepList: 小png列表中需要保存的图的序号
    """
    img = PILI.open(fn)
    w, h = img.size
    if rowNum <= h and colNum <= w:
        print('Original image info: %sx%s, %s, %s' % (w, h, img.format, img.mode))
        i, j = 0, 0
        h = h // rowNum
        w = w // colNum
        fn = fn.split('.')[0]
        for r in range(rowNum):
            for c in range(colNum):
                box = (c * w, r * h, (c + 1) * w, (r + 1) * h)
                i += 1
                if i in keepList:
                    img.crop(box).save(fn + '_' + str(j) + ".png", "png")
                    j += 1