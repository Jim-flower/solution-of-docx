from win32com import client

#转换doc为docx
def doc2docx(fn):
    word = client.Dispatch("Word.Application") # 打开word应用程序
    #for file in files:
    doc = word.Documents.Open(fn) #打开word文件
    doc.SaveAs("{}x".format(fn), 12)#另存为后缀为".docx"的文件，其中参数12或16指docx文件
     #关闭原来word文件
    doc.Close()
    word.Quit()





doc2docx(r"C:\Users\Administrator\Desktop\test.doc")