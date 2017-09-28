# _*_ coding: utf-8 _*_
# @Time    : 2017/9/11 21:36
# @Author  : Jimmy Yu
# @File    : Job_CheckWfproError.py
# @Remark  : 检测工作流中使用的存储过程错误信息
import myUtil.MailUtil
import myUtil.JobLog
from sqlalchemy import create_engine
import os
import pandas as pd
from pandas.io import sql

os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
pd.set_option('display.max_colwidth', -1)
e_oa = create_engine('oracle://oadb:oracle@192.168.0.89:1521/OADB',echo = True)

receiver = ['jimmyyu@fortune-co.com']
subject = "OA Procedure Error Log"
try:
    errdata = pd.read_sql_query("select t.REQUESTID,t.DATETIME,t.PRONAME,t.ERRORLOG from cuspro_errorLog t",e_oa)
    htmlContent = """
    <!DOCTYPE html>
    <html>
    
    <head>
      <meta charset="utf-8">
      <style>
      table {
          border-collapse: collapse;
          width: 100%;
      }
    
      th, td {
          text-align: left;
          padding: 8px;
      }
    
      tr:nth-child(even){background-color: #f2f2f2}
    
      th {
          background-color: #4CAF50;
          color: white;
      }
      </style>
    </head>
    
    <body>
    """+errdata.to_html(index = False,border = 0)+"""
    </body>
    
    </html>
    
    """
    if (len(errdata)>0) :
        myUtil.MailUtil.sendHtmlMailTo(receiver,subject,htmlContent)
        sql.execute('delete from cuspro_errorLog', e_oa)


except Exception as e:
    myUtil.JobLog.logger.error("存储过程错误检查：%s".format(e),exc_info=True)
    raise
