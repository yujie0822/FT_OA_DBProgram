# _*_ coding: utf-8 _*_
# @Time    : 2017/8/8 13:02
# @Author  : Jimmy Yu
# @File    : DailyReport_So.py
import cx_Oracle,myUtil.MailUtil,myUtil.SqlEnv,myUtil.ReportLog,datetime,smtplib,math

now=datetime.datetime.now()
yesterday = now+datetime.timedelta(days=-1)
date_today = now.strftime('%Y-%m-%d')
time_now = now.strftime('%H:%M:%S')
date_yesterday = yesterday.strftime('%Y-%m-%d')
date_yesterdayList = [yesterday.month,yesterday.day]


oaConn = cx_Oracle.connect(myUtil.SqlEnv.MAIN_OA_CONNECT_STRING)
oaCursor = oaConn.cursor()
myUtil.ReportLog.logger.info("销售订单流程报表：OA Database Connected")

def createReportSql(wfid,dateFromStr,dateToStr):
    resultStr = "select t_req.requestid,t_req.currentnodeid from workflow_requestbase \
t_req left join workflow_nodebase t_node on t_node.id = t_req.currentnodeid \
where (t_req.workflowid = "+str(wfid)+" ) and t_req.createdate >= \'"+dateFromStr+"\'\
and t_req.createdate <= \'"+dateToStr+"\'"
    return resultStr

def createPersonalReportSql(userid,wfid1,wfid2,dateFromStr1,dateToStr1,dateFromStr2,dateToStr2):
    resultStr = """
(select (select count(t_ope.requestid) from workflow_currentoperator t_ope
left join workflow_requestbase t_req on t_req.requestid = t_ope.requestid
where t_ope.userid = {0} and t_ope.workflowid = {1} and t_req.createdate >= '{3}' and t_req.createdate <= '{4}')
as ACT,
(select count(t_ope.requestid) from workflow_currentoperator t_ope
left join workflow_requestbase t_req on t_req.requestid = t_ope.requestid
where t_ope.userid = {0} and ( t_ope.workflowid = {2} ) and t_req.createdate >= '{3}' and t_req.createdate <= '{4}')
as PAS from dual)
union
(select (select count(t_ope.requestid) from workflow_currentoperator t_ope
left join workflow_requestbase t_req on t_req.requestid = t_ope.requestid
where t_ope.userid = {0} and t_ope.workflowid = {1} and t_req.createdate >= '{5}' and t_req.createdate <= '{6}')
as ACT,
(select count(t_ope.requestid) from workflow_currentoperator t_ope
left join workflow_requestbase t_req on t_req.requestid = t_ope.requestid
where t_ope.userid = {0} and ( t_ope.workflowid = {2} ) and t_req.createdate >= '{5}' and t_req.createdate <= '{6}')
as PAS from dual)
""".format(userid,wfid1,wfid2,dateFromStr1,dateToStr1,dateFromStr2,dateToStr2)
    return resultStr
try:
    #Passive今日销售订单节点状态List
    oaCursor.execute(createReportSql("61 or t_req.workflowid = 522 ",date_yesterday,date_yesterday))
    pasTodayStatus = oaCursor.fetchall()
    pasTodayNodes = [x[1] for x in pasTodayStatus]
    pasCol = [0 for x in range(9)]
    #CS
    pasCol[0] = pasTodayNodes.count(261)+pasTodayNodes.count(1522)
    #销售
    pasCol[1] = pasTodayNodes.count(1081)+pasTodayNodes.count(1536)
    #法务
    pasCol[2] = pasTodayNodes.count(908)+pasTodayNodes.count(1535)
    #CS主管
    pasCol[3] = pasTodayNodes.count(262)+pasTodayNodes.count(1523)
    #区域经理
    pasCol[4] = pasTodayNodes.count(264)+pasTodayNodes.count(325)+pasTodayNodes.count(1525)+pasTodayNodes.count(1533)
    #销售副总
    pasCol[5] = pasTodayNodes.count(265)+pasTodayNodes.count(1526)
    #董事长
    pasCol[6] = pasTodayNodes.count(268)+pasTodayNodes.count(1529)
    # 流程结束
    pasCol[7] = pasTodayNodes.count(269)+pasTodayNodes.count(1530)
    #总计
    pasCol[8] = len(pasTodayNodes)-pasTodayNodes.count(426)-pasTodayNodes.count(1534)

    #Active今日销售订单节点状态List
    oaCursor.execute(createReportSql(62,date_yesterday,date_yesterday))
    actTodayStatus = oaCursor.fetchall()
    actTodayNodes = [x[1] for x in actTodayStatus]
    actCol = [0 for x in range(9)]
    #CS
    actCol[0] = actTodayNodes.count(270)
    #销售
    actCol[1] = actTodayNodes.count(1061)
    #法务
    actCol[2] = actTodayNodes.count(909)
    #CS主管
    actCol[3] = actTodayNodes.count(271)
    #区域经理
    actCol[4] = actTodayNodes.count(273)+actTodayNodes.count(322)
    #销售副总
    actCol[5] = actTodayNodes.count(274)
    #董事长
    actCol[6] = actTodayNodes.count(277)
    # 流程结束
    actCol[7] = actTodayNodes.count(278)
    #总计
    actCol[8] = len(actTodayNodes)-actTodayNodes.count(427)

    todaySumCol = [pasCol[x]+actCol[x] for x in range(9)]
    if(todaySumCol[8] == 0):
        percentCol = [0 for x in range(9)]
    else:
        percentCol = [int(round(float(todaySumCol[x])*100.0/float(todaySumCol[8]))) for x in range(9)]

    #Passive累计销售订单节点状态List
    oaCursor.execute(createReportSql("61 or t_req.workflowid = 522 ",'2017-07-01',date_yesterday))
    pasAllStatus = oaCursor.fetchall()
    pasAllNodes = [x[1] for x in pasAllStatus]
    pasCol_t = [0 for x in range(9)]
    #CS
    pasCol_t[0] = pasAllNodes.count(261)+pasAllNodes.count(1522)
    #销售
    pasCol_t[1] = pasAllNodes.count(1081)+pasAllNodes.count(1536)
    #法务
    pasCol_t[2] = pasAllNodes.count(908)+pasAllNodes.count(1535)
    #CS主管
    pasCol_t[3] = pasAllNodes.count(262)+pasAllNodes.count(1523)
    #区域经理
    pasCol_t[4] = pasAllNodes.count(264)+pasAllNodes.count(325)+pasAllNodes.count(1525)+pasAllNodes.count(1533)
    #销售副总
    pasCol_t[5] = pasAllNodes.count(265)+pasAllNodes.count(1526)
    #董事长
    pasCol_t[6] = pasAllNodes.count(268)+pasAllNodes.count(1529)
    # 流程结束
    pasCol_t[7] = pasAllNodes.count(269)+pasAllNodes.count(1530)
    #总计
    pasCol_t[8] = len(pasAllNodes)-pasAllNodes.count(426)-pasAllNodes.count(1534)

    #Active累计销售订单节点状态List
    oaCursor.execute(createReportSql(62,'2017-07-01',date_yesterday))
    actAllStatus = oaCursor.fetchall()
    actAllNodes = [x[1] for x in actAllStatus]
    actCol_t = [0 for x in range(9)]
    #CS
    actCol_t[0] = actAllNodes.count(270)
    #销售
    actCol_t[1] = actAllNodes.count(1061)
    #法务
    actCol_t[2] = actAllNodes.count(909)
    #CS主管
    actCol_t[3] = actAllNodes.count(271)
    #区域经理
    actCol_t[4] = actAllNodes.count(273)+actAllNodes.count(322)
    #销售副总
    actCol_t[5] = actAllNodes.count(274)
    #董事长
    actCol_t[6] = actAllNodes.count(277)
    # 流程结束
    actCol_t[7] = actAllNodes.count(278)
    #总计
    actCol_t[8] = len(actAllNodes)-actAllNodes.count(427)

    allSumCol = [pasCol_t[x]+actCol_t[x] for x in range(9)]
    percentCol_t = [int(round(float(allSumCol[x])*100.0/float(allSumCol[8]))) for x in range(9)]


    #合计数校验
    pasSum = 0
    for x in pasCol[:-1]:
        pasSum+=x
    if(pasSum!=pasCol[8]):
        myUtil.ReportLog.logger.warning("销售订单流程报表：PAS总和异常")
    else:
        myUtil.ReportLog.logger.debug("销售订单流程报表：PAS总和正常")
    actSum = 0
    for x in actCol[:-1]:
        actSum+=x
    if(actSum!=actCol[8]):
        myUtil.ReportLog.logger.warning("销售订单流程报表：ACT总和异常")
    else:
        myUtil.ReportLog.logger.debug("销售订单流程报表：ACT总和正常")

    oaCursor.execute(createPersonalReportSql(21,62,"61 or t_ope.workflowid = 522",date_yesterday,date_yesterday,'2017-07-10',date_yesterday))
    lawStatus = oaCursor.fetchall()
    law_total_today = lawStatus[0][0]+lawStatus[0][1]
    law_act_today = lawStatus[0][0]
    law_pas_today = lawStatus[0][1]

    law_total_all = lawStatus[1][0]+lawStatus[1][1]
    law_act_all = lawStatus[1][0]
    law_pas_all = lawStatus[1][1]


    if((todaySumCol[7]+todaySumCol[6]) == 0):
        law_total_today_p=0
    else:
        law_total_today_p=int(round(law_total_today*100.0/(todaySumCol[7]+todaySumCol[6])))

    if((actCol[7]+actCol[6]) == 0):
        law_act_today_p=0
    else:
        law_act_today_p=int(round(law_act_today*100.0/(actCol[7]+actCol[6])))

    if((pasCol[7]+pasCol[6]) == 0):
        law_pas_today_p=0
    else:
        law_pas_today_p = int(round(law_pas_today*100.0/(pasCol[7]+pasCol[6])))

    if((allSumCol[7]+allSumCol[6])==0):
        law_total_all_p=0
    else:
        law_total_all_p=int(round(law_total_all*100.0/(allSumCol[7]+allSumCol[6])))

    if((actCol_t[7]+actCol_t[6])==0):
        law_act_all_p=0
    else:
        law_act_all_p=int(round(law_act_all*100.0/(actCol_t[7]+actCol_t[6])))

    if((pasCol_t[7]+pasCol_t[6])==0):
        law_pas_all_p=0
    else:
        law_pas_all_p=int(round(law_pas_all*100.0/(pasCol_t[7]+pasCol_t[6])))


    #--------------------发送Email部分-------------------------
    receiver = ['ERPSUPPORT@fortune-co.com','jacksun@fortune-co.com']
    # receiver = ['jimmyyu@fortune-co.com']
    subject = str(date_yesterdayList[0])+'月'+str(date_yesterdayList[1])+'日销售订单审批流程试运行总结'
    htmlContent = """
<html>
<head>
  <style media="screen">
    .myHead{
      padding: 5px 5px 10px 10px;
    }
    .myText{
      padding: 5px 10px 0px 30px;
    }
    .myTableRow{
      padding: 10px 0px 5px 5px;
    }
    .myTable{
      border-collapse:collapse;
    }
    .myTable th,td{
      border: 1px solid black;
      padding:3px 7px 2px 7px;
    }

  </style>
</head>
"""+"""
<body>
  <div>
    <h2 class="myHead">OA流程节点统计日报表</h2>
  </div>
  <div class="myText">
    <p>以下分别列出{date[0]}月{date[1]}日当日及截止{date[0]}月{date[1]}日累计销售订单审批流程试运行的情况：</p>
    <p>{date[0]}月{date[1]}日当日：</p>
    <ol>
      <li>今天的有效订单一共{l3[8]}单（ACT共{l1[8]}单/PAS共{l2[8]}单），流程结束的一共{l3[7]}单（ACT有{l1[7]}单/PAS有{l2[7]}单），
        批到Lawrence的{law_total_today}单-总占比{law_total_today_p}%（ACT有{law_act_today}单-占比{law_act_today_p}%
        /PAS有{law_pas_today}单-占比{law_pas_today_p}%）</li>
      <li>流程各节点占比情况:
        <div class="myTableRow">
          <table class="myTable">
            <tr>
              <th>节点</th>
              <th>ACT</th>
              <th>PAS</th>
              <th>小计</th>
              <th>总占比</th>
            </tr>
            <tr>
              <td>CS</td>
              <td>{l1[0]}</td>
              <td>{l2[0]}</td>
              <td>{l3[0]}</td>
              <td>{l4[0]}%</td>
            </tr>
            <tr>
              <td>业务员</td>
              <td>{l1[1]}</td>
              <td>{l2[1]}</td>
              <td>{l3[1]}</td>
              <td>{l4[1]}%</td>
            </tr>
            <tr>
              <td>法务</td>
              <td>{l1[2]}</td>
              <td>{l2[2]}</td>
              <td>{l3[2]}</td>
              <td>{l4[2]}%</td>
            </tr>
            <tr>
              <td>CS主管</td>
              <td>{l1[3]}</td>
              <td>{l2[3]}</td>
              <td>{l3[3]}</td>
              <td>{l4[3]}%</td>
            </tr>
            <tr>
              <td>区域经理</td>
              <td>{l1[4]}</td>
              <td>{l2[4]}</td>
              <td>{l3[4]}</td>
              <td>{l4[4]}%</td>
            </tr>
            <tr>
              <td>销售副总</td>
              <td>{l1[5]}</td>
              <td>{l2[5]}</td>
              <td>{l3[5]}</td>
              <td>{l4[5]}%</td>
            </tr>
            <tr>
              <td>董事长</td>
              <td>{l1[6]}</td>
              <td>{l2[6]}</td>
              <td>{l3[6]}</td>
              <td>{l4[6]}%</td>
            </tr>
            <tr>
              <td>流程结束</td>
              <td>{l1[7]}</td>
              <td>{l2[7]}</td>
              <td>{l3[7]}</td>
              <td>{l4[7]}%</td>
            </tr>
            <tr>
              <th>总计</th>
              <td>{l1[8]}</td>
              <td>{l2[8]}</td>
              <td>{l3[8]}</td>
              <td>{l4[8]}%</td>
            </tr>
          </table>
        </div>
      </li>
    </ol>
    <p>截止{date[0]}月{date[1]}日累计：</p>
    <ol>
      <li>累计的有效订单一共{l3_t[8]}单（ACT共{l1_t[8]}单/PAS共{l2_t[8]}单），
        流程结束的一共{l3_t[7]}单（ACT有{l1_t[7]}单/PAS有{l2_t[7]}单），
        批到Lawrence的{law_total_all}单-总占比{law_total_all_p}%
        （ACT有{law_act_all}单-占比{law_act_all_p}%
        /PAS有{law_pas_all}单-占比{law_pas_all_p}%）</li>
        <li>流程各节点占比情况:
          <div class="myTableRow">
            <table class="myTable">
              <tr>
                <th>节点</th>
                <th>ACT</th>
                <th>PAS</th>
                <th>小计</th>
                <th>总占比</th>
              </tr>
              <tr>
                <td>CS</td>
                <td>{l1_t[0]}</td>
                <td>{l2_t[0]}</td>
                <td>{l3_t[0]}</td>
                <td>{l4_t[0]}%</td>
              </tr>
              <tr>
                <td>业务员</td>
                <td>{l1_t[1]}</td>
                <td>{l2_t[1]}</td>
                <td>{l3_t[1]}</td>
                <td>{l4_t[1]}%</td>
              </tr>
              <tr>
                <td>法务</td>
                <td>{l1_t[2]}</td>
                <td>{l2_t[2]}</td>
                <td>{l3_t[2]}</td>
                <td>{l4_t[2]}%</td>
              </tr>
              <tr>
                <td>CS主管</td>
                <td>{l1_t[3]}</td>
                <td>{l2_t[3]}</td>
                <td>{l3_t[3]}</td>
                <td>{l4_t[3]}%</td>
              </tr>
              <tr>
                <td>区域经理</td>
                <td>{l1_t[4]}</td>
                <td>{l2_t[4]}</td>
                <td>{l3_t[4]}</td>
                <td>{l4_t[4]}%</td>
              </tr>
              <tr>
                <td>销售副总</td>
                <td>{l1_t[5]}</td>
                <td>{l2_t[5]}</td>
                <td>{l3_t[5]}</td>
                <td>{l4_t[5]}%</td>
              </tr>
              <tr>
                <td>董事长</td>
                <td>{l1_t[6]}</td>
                <td>{l2_t[6]}</td>
                <td>{l3_t[6]}</td>
                <td>{l4_t[6]}%</td>
              </tr>
              <tr>
                <td>流程结束</td>
                <td>{l1_t[7]}</td>
                <td>{l2_t[7]}</td>
                <td>{l3_t[7]}</td>
                <td>{l4_t[7]}%</td>
              </tr>
              <tr>
                <th>总计</th>
                <td>{l1_t[8]}</td>
                <td>{l2_t[8]}</td>
                <td>{l3_t[8]}</td>
                <td>{l4_t[8]}%</td>
              </tr>
            </table>
          </div>
        </li>
    </ol>
  </div>
</body>

</html>

""".format(date=date_yesterdayList,l1=actCol,l2=pasCol,l3=todaySumCol,l4=percentCol,\
l1_t=actCol_t,l2_t=pasCol_t,l3_t=allSumCol,l4_t=percentCol_t,\
law_total_today=law_total_today,law_total_today_p=law_total_today_p,\
law_act_today=law_act_today,law_act_today_p=law_act_today_p,\
law_pas_today=law_pas_today,law_pas_today_p=law_pas_today_p,\
law_total_all=law_total_all,law_total_all_p=law_total_all_p,\
law_act_all=law_act_all,law_act_all_p=law_act_all_p,\
law_pas_all=law_pas_all,law_pas_all_p=law_pas_all_p,\
)

    myUtil.MailUtil.sendHtmlMailTo(receiver,subject,htmlContent)
    myUtil.ReportLog.logger.info("销售订单流程报表：邮件发送成功")

except Exception as e:
    myUtil.ReportLog.logger.error("销售订单流程报表：Error"+str(e),exc_info=True)
    raise

finally:
    oaCursor.close()
    oaConn.close()
    myUtil.ReportLog.logger.info("销售订单流程报表：OA Database Disconnected")
