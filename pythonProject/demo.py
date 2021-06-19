import docx
import json
import re
from loguru import logger
import win32com.client as wc
import shutil
import os
logger.add("log/interface_log_{time}.log", rotation="500MB", encoding="utf-8", enqueue=True, compression="zip",
           retention="10 days")
newfileNames = os.listdir(r"data/")
l = 1
e = 1
# lines = []
# with open("法律种类.txt", 'r', encoding='utf-8') as fp:
#     lines = fp.readlines()
#     for i in range(len(lines)):
#         lines[i] = re.findall(r'[《](.*?)[》]', lines[i])[0]
#     fp.close()
# dict_lines = dict.fromkeys(lines, 0)
# print(dict_lines)
dict_lines = dict()
for file in newfileNames:
    try:
        # 注意SaveAs会打开保存后的文件，有时可能看不到，但后台一定是打开的
        path = "data/" + file.split(".docx")[0] + ".docx"
        # 获取文档的所有段落 path : 相对路径包含文档名称
        docx_temp = docx.Document(path)
        for para in docx_temp.paragraphs:
                if re.findall("(.*)判决如下", str(para.text)):
                    t = re.findall("(.*)判决如下", str(para.text))[0]
                    law_list = re.findall(r'[《](.*?)[》]', t)
                    for i in law_list:
                        if i in dict_lines.keys():
                            dict_lines[i] += 1
                        else:
                            dict_lines.update({i:0})
                        logger.info(str(l)+"<:::>"+str(dict_lines[i])+"<::>"+i)
                    l += 1
                    break
                #     if "中华人民共和国消费者权益保护法" in law_list:
                #     with open("keywords/中华人民共和国消费者权益保护法/综上x判决如下.txt", 'a+', encoding='utf-8') as fp:
                #         fp.write(str(file) + '\n')
                #         fp.close()
                #     logger.info("[" + str(l) + "]" + str(file))
                #     l += 1
                #     break
                #     print(str(l) + "<:>" + str(re.findall("根据(.*)判决如下", str(para.text))))
                #     print(str(l) + "<:>" +str(re.findall("(.*)判决如下", str(para.text))[0]))
                #     l += 1
    except:
        with open("keywords/中华人民共和国消费者权益保护法_print_error.txt", 'a+', encoding='utf-8') as fp:
            fp.write(str(file) + '\n')
            fp.close()
        logger.info("[" + str(e) + "]" + str(file))
        e += 1
        continue
dict_lines = sorted(dict_lines.items(), key=lambda x:x[1])
logger.info(dict_lines)