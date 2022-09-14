import docx
import json
from collections import defaultdict

doc = docx.Document(r"./explanation.docx")
nounExp = defaultdict(list)
LastStyle = ""
for paragraph in doc.paragraphs:
    styleName = paragraph.style.name
    Text = paragraph.text
    if styleName ==  "Heading 3":
        LastStyle = styleName
        Lastkey = Text
    elif styleName ==  "Normal":
        if LastStyle == "Heading 3":
            nounExp[Lastkey].append(Text)
    else:
        LastStyle = styleName


# 保存文件
tf = open("myDictionary.json", "w")
json.dump(nounExp,tf)
tf.close()











