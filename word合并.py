import os
import win32com.client as win32

word = win32.gencache.EnsureDispatch("Word.Application")
# 启动word对象应用
word.Visible = False
path = r"D:/word文件合并"
# 需要合并的文件所在路径
files = []
filename = os.listdir(path)
filename.sort(key=lambda x:int(x[:-5]))
# 按照数据顺序进行排序，既对".docx"前面的进行排序
for filename in filename:
    filename = os.path.join(path, filename)
    files.append(filename)
    # 获取目录下所有文件的路径

output = word.Documents.Add()
# 新建空的WORD文档
for file in files:
    output.Application.Selection.InsertFile(file)
    # 拼接文档

doc = output.Range(output.Content.Start, output.Content.End)
# 获取合并后文档的内容

output.SaveAs("d:/word合并.docx")
# 把汇总文件保存到指定路径
output.Close()  # 关闭
