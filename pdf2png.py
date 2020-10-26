# -*- coding: utf-8 -*-
from wand.image import Image
from wand.color import Color
from PyPDF2 import PdfFileReader, PdfFileWriter
import io


def convertPdfIntoPng(fn, on, pageIndex=0, res=200):
    """

    fn: 输入pdf文件名
    on: 输出png文件名
    pageIndex: 待处理的页号
    res: 输出png分辨率
    """
    pdf = PdfFileReader(fn, strict=False)
    pageObj = pdf.getPage(pageIndex)
    dst_pdf = PdfFileWriter()
    dst_pdf.addPage(pageObj)
    pdf_bytes = io.BytesIO()
    dst_pdf.write(pdf_bytes)
    pdf_bytes.seek(0)
    img = Image(file=pdf_bytes, resolution=res)
    img.format = "png"
    img.compression_quality = 90
    img.background_color = Color("white")
    img.save(filename=on)
    img.destroy()
