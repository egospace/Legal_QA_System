import os
import re
import shutil
from loguru import logger
import docx
import win32com.client as wc

logger.add("log/interface_log_{time}.log", rotation="500MB", encoding="utf-8", enqueue=True, compression="zip",
           retention="10 days")


# doc文件另存为docx
def doc_transform_docx():
    word = wc.Dispatch("Word.Application")
    filenames = os.listdir(r"rawData")
    l = 1
    print(len(filenames))
    for file in filenames:
        logger.info(str(l) + ":" + str(file))
        l += 1
        try:
            doc = word.Documents.Open("D:\\transform\\rawData" + file)
            # 上面的地方只能使用完整绝对地址，相对地址找不到文件，且，只能用“\\”，不能用“/”，哪怕加了 r 也不行，涉及到将反斜杠看成转义字符。
            doc.SaveAs("D:\\transform\\newData" + file.split(".doc")[0] + ".docx", 12)  # 或直接简写
            if l % 1000 == 0:
                os.system("taskkill /F /IM WINWORD.EXE")
                word = wc.Dispatch("Word.Application")
        except:
            continue


def docx_transform_classify():
    newfileNames = os.listdir(r"D:\\Code\\pythoncode\\pythonProject\\data\\")
    l = 1
    e = 1
    for file in newfileNames:
        # 注意SaveAs会打开保存后的文件，有时可能看不到，但后台一定是打开的
        path = "D:\\Code\\pythoncode\\pythonProject\\data\\" + file
        # print(path)
        # 获取文档的所有段落 path : 相对路径包含文档名称
        try:

            docx_temp = docx.Document(path)

            for para in docx_temp.paragraphs:
                if re.findall("本院经审理确认如下事实", str(para.text)):
                    with open("keywords/本院经审理确认如下事实.txt", 'a+', encoding='utf-8') as fp:
                        fp.write(str(file) + '\n')
                        fp.close()
                    logger.info("[" + str(l) + "]" + str(file))
                    l += 1
                    break
        except:
            with open("keywords/本院经审理确认如下事实_error.txt", 'a+', encoding='utf-8') as fp:
                fp.write(str(file) + '\n')
                fp.close()
            logger.info("[" + str(e) + "]" + str(file))
            e += 1
            continue
        # fact = str(findpara(getpara(path), "事实与理由")).split("事实与理由：")[1]
        # laws = (re.findall(r'[《](.*?)[》]', str(findpara(getpara(path), "综上所述")).split("综上所述，")[1]))
        # res = {'fact': fact, 'laws': laws}
        # res2json = json.dumps(res, ensure_ascii=False)
        # with open(resPath, 'a+', encoding='utf-8') as fp:
        #     fp.write(res2json + '\n')
        #     fp.close()


def move_doc():
    word = wc.Dispatch("Word.Application")
    filenames = os.listdir(r"data")
    l = 1
    print(len(filenames))
    for f in range(893):
        print(str(l) + ":" + str(filenames[f]))
        l += 1
        shutil.move(r"D:\\Code\\pythoncode\\pythonProject\\data\\" + filenames[f],
                    r"D:\\Code\\pythoncode\\pythonProject\\demo16\\")


def move_docx():
    with open("keywords/诉称null.txt", 'r', encoding='utf-8') as fp:
        lines = fp.readlines()
        l = 1
        e = 1
        print(len(lines))
        for file in lines:
            try:
                logger.info(str(l) + ":" + str(file))
                shutil.move(r"dataxf/诉称/" + file.strip("\n"),
                            r"dataxf/法条/诉称/")
                l += 1
            except:
                with open("keywords/move_docx_error.txt", 'a+', encoding='utf-8') as fp:
                    fp.write(str(file) + '\n')
                    fp.close()
                logger.info("["+str(e)+"]"+"error:"+str(file))
                e += 1
        fp.close()


if __name__ == "__main__":
    # srcPath = input()
    # desPath = input()
    # doc_transform_docx(srcPath, desPath)
    # move_doc()
    # docx_transform_classify()
    move_docx()
    # pass
    # doc_transform_docx()

