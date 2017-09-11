# _*_ coding: utf-8 _*_
# @Time    : 2017/8/27 22:50
# @Author  : Jimmy Yu
# @File    : updateJdeCus.py
# @Remark  : 同步JDE客户信息至OA
import cx_Oracle,myUtil.SqlEnv,myUtil.MailUtil,myUtil.JobLog,datetime,smtplib,math

oaConn = cx_Oracle.connect(myUtil.SqlEnv.MAIN_OA_CONNECT_STRING)
oaCursor = oaConn.cursor()
try:
    oaCursor.execute("select c_id from uf_customers")
    rows = oaCursor.fetchall()
    for row in rows:
        oaCursor.callproc('updatejdecus', [row[0]])
except Exception as e:
    raise

finally:
    oaCursor.close()
    oaConn.close()