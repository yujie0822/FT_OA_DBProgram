# _*_ coding: utf-8 _*_
# @Time    : 2017/8/9 13:02
# @Author  : Jimmy Yu
# @File    : Job_UpdateLSZQ.py
# @Remark  : 临时账期停用检查，停用到期日为明日的临时账期

import cx_Oracle,myUtil.SqlEnv,myUtil.MailUtil,myUtil.JobLog,datetime,smtplib,math

oaConn = cx_Oracle.connect(myUtil.SqlEnv.MAIN_OA_CONNECT_STRING)
oaCursor = oaConn.cursor()
myUtil.JobLog.logger.debug("临时账期停用：OA Database Connected")
try:
    endType = oaCursor.var(cx_Oracle.NUMBER)
    msg = oaCursor.var(cx_Oracle.STRING)
    oaCursor.callproc('P_JOB_UPDATEZQ',[msg,endType])
    if endType.getvalue() == 0:
        myUtil.JobLog.logger.info('Run Successfully')
        myUtil.MailUtil.sendTextMailTo(["jimmyyu@fortune-co.com", "yujie0822@163.com"], "Run With Successfully-临时账期停用程序",
                                       msg.getvalue())
    elif endType.getvalue() == 1:
        myUtil.JobLog.logger.info('Run with Warnings')
        myUtil.MailUtil.sendTextMailTo(["jimmyyu@fortune-co.com", "yujie0822@163.com"], "Run With Warning-临时账期停用程序",
                                       msg.getvalue())
    else:
        myUtil.JobLog.logger.error('Run with Error')
        myUtil.MailUtil.sendTextMailTo(["jimmyyu@fortune-co.com", "yujie0822@163.com"], "Run With Error-临时账期停用程序",
                                       msg.getvalue())
    myUtil.JobLog.logger.info(msg.getvalue())

except Exception as e:
    myUtil.JobLog.logger.error("临时账期停用：%s".format(e), exc_info=True)
    myUtil.MailUtil.sendTextMailTo(["jimmyyu@fortune-co.com", "yujie0822@163.com"], "Run With Exception-临时账期停用程序",e)
    raise

finally:
    oaCursor.close()
    oaConn.close()
    myUtil.JobLog.logger.debug("临时账期停用：OA Database Disconnected")