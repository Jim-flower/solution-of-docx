import os
from win32com import client
from op3 import get_pictures
def doc2docx(fn):
    word = client.Dispatch("Word.Application") # 打开word应用程序
    #for file in files:
    doc = word.Documents.Open(fn) #打开word文件
    doc.SaveAs("{}x".format(fn), 12)#另存为后缀为".docx"的文件，其中参数12或16指docx文件
     #关闭原来word文件
    doc.Close()
    word.Quit()
if __name__ == '__main__':
    c_lsit = os.listdir()
    for x in c_lsit:
        if (".doc" in x )&(x[-1]=="c"):
            doc2docx("C:\\Users\\Administrator\\Desktop\\demo\\"+x)
            path2 = r"C:\Users\Administrator\Desktop\demo"
            get_pictures(x+"x", path2)
            os.remove(x+"x")
            os.remove(x)
        elif(".docx")in x:
            path2 = r"C:\Users\Administrator\Desktop\demo"
            get_pictures(x, path2)
            os.remove(x)
        elif(".pdf")in x:
            os.remove(x)
