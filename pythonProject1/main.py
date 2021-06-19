from flask import Flask, request, jsonify
import pymysql
import predict
app = Flask(__name__)
conn = pymysql.connect(host="localhost", port=3306, user="root", password="123456", database="law",
                       charset="utf8")


def db(ft):
    # 得到一个可以执行SQL语句的光标对象
    cursor = conn.cursor()
    laws = []
    for i in ft:
        # 定义要执行的SQL语句
        sql = "select * from ft where law_id=%s;"
        # 执行SQL语句
        cursor.execute(sql, [i])
        ret = cursor.fetchall()
        result = {'laws': ret[0][1], 'content': ret[0][2]}
        laws.append(result)
    # 关闭光标对象
    cursor.close()
    ret = {"data": laws}
    return ret


@app.route('/fact', methods=['POST'])
def fact():
    laws = predict.get_label(eval(request.get_data())['fact'])
    if len(laws) > 0:
        return jsonify(db(laws))
    else:
        return jsonify({"data": [{"laws": "抱歉! 关于此类事件，系统正在学习中~", "content": ""}]})


if __name__ == '__main__':
    app.run(port=8080, debug=True)
    # 关闭数据库连接
    conn.close()
