# _*_ coding: utf-8 _*_
# @Time    : 2017/9/28 17:10
# @Author  : Jimmy Yu
# @File    : fhdCheck.py
# @Remark  : 检测特批发货完成后，发货单状态没有被更新的问题
import myUtil.MailUtil
import myUtil.JobLog
from sqlalchemy import create_engine
import os
import pandas as pd
from pandas.io import sql

os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'
pd.set_option('display.max_colwidth', -1)
e_oa = create_engine('oracle://oadb:oracle@192.168.0.89:1521/OADB')

receiver = ['jimmyyu@fortune-co.com']
subject = "特批发货发货单状态异常"
try:
    errdata = pd.read_sql_query("""
    select doco from
(select distinct(t1.DOCO) from formtable_main_16 t1 
join workflow_requestbase t2 on t1.requestid = t2.requestid
left join proddta.f4211@jde_dblink t3 on t3.sdshpn = t1.doco
where t2.createdate >= to_char(sysdate-1,'YYYY-MM-DD') and t2.workflowid = 63
and t3.SDNXTR < 560 and t2.currentnodeid = 287) where not exists (select 1 from formtable_main_74 t4 where t4.fhdh = doco)
    """, e_oa)
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
    """ + errdata.to_html(index=False, border=0) + """
    </body>

    </html>

    """
    if (len(errdata) > 0):
        myUtil.MailUtil.sendHtmlMailTo(receiver, subject, htmlContent)


except Exception as e:
    myUtil.JobLog.logger.error("特批发货发货单检查：%s".format(e), exc_info=True)
    raise