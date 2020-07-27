# 批量生成word文件
from docx import Document
import pandas as pd

# 读取数据
df = pd.read_excel('./姓名.xls', 'Sheet1')

# 替换函数
def replace_text(old_text, new_text):
    paragraphs = document.paragraphs# 所有段落
    for paragraph in paragraphs:# 遍历段落
        #print(paragraph)
        for run in paragraph.runs:# 遍历段落内容
            #print(run.text)
            run_text = run.text.replace(old_text, new_text)# 替换生成副本
            run.text = run_text# 副本替换实际内容# 易错点 run.text写成run_text, 则无法对run进行替换
            #print(run.text)
# 行
for row in range(0,df.shape[0]):# 遍历行
    document= Document('./模板.docx')# 读取模板文件
    my_col = df.columns.tolist()# 读取数据表列名
    
    for col in range(0,df.shape[1]):# 在每行下遍历列
        new_text = df.iloc[row,col]# 读取特定行列下内容
        old_text = my_col[col]# 读取对应列列名， 即为模板中内容
        # 替换
        replace_text(str(old_text),str(new_text))
#         print(old_text,new_text)
        
    document.save('./{}.docx'.format(df.iloc[row,0]))# 输出保存，以名字命名文件，每一行的数据为一个文件
