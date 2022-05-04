from pdf2docx import Converter
pdf_file = r"E:\Spyderworkplace\excel\湖南\2019_pdf\附件：湖南省注册会计师协会关于我省注册会计师2019年度任职资格检查合格人员名单（第一批）.pdf"
docx_file = r"E:\Spyderworkplace\excel\湖南\2019_pdf\附件：湖南省注册会计师协会关于我省注册会计师2019年度任职资格检查合格人员名单（第一批）.docx"
cv = Converter(pdf_file)
cv.convert(docx_file,start=0,end=None)
cv.close()