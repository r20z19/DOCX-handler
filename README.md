# DOCX-handler
DOCX论文文件排版器，方便学生党的快速排版神器

### 安装：
git clone https://github.com/r20z19/DOCX-handler.git

cd DOCX-handler

pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

### 使用：
python app.py 然后浏览器进入127.0.0.1:5000

1.首先按照想要的四级标识进行写docx或doc文档

<img width="1170" height="1016" alt="屏幕截图 2026-01-09 203709" src="https://github.com/user-attachments/assets/0deafb30-2207-4ac3-899e-3d9824712a8d" />

2.然后清除全部格式，以免造成代码识别错误

<img width="1169" height="1247" alt="屏幕截图 2026-01-09 203724" src="https://github.com/user-attachments/assets/f5f0c6e2-f604-4d41-abb2-f67f7acc5d79" />

3.接下来把文档内容全部粘贴到新的空的docx或doc文档，以免不可见字符造成代码识别错误

<img width="1898" height="1055" alt="屏幕截图 2026-01-09 203758" src="https://github.com/user-attachments/assets/23a02239-bc20-4572-8896-9317731c3f41" />

4.最后进入127.0.0.1:5000 选择对应的标识，字体，字号，是否加粗，再上传文档即可

<img width="1038" height="1183" alt="屏幕截图 2026-01-09 203859" src="https://github.com/user-attachments/assets/a5a09bc0-8e86-4845-8629-40941f5e922c" />


### 注意：
1.docx里面不能出现回车自动排版的格式，比如一级标题是（一）  （二） （三），回车生成的（四）是不会被代码检测到的，必须是手动打的（四）才会被代码检测并处理。

2.某些隐形的杂格式也会导致代码处理失败，建议新建一个空docx，然后把所有内容粘贴到新docx中，再进行处理。

### 如果喜欢的点个star吧，如果多的话，会考虑更新新版本排版器
