#!/usr/bin/python
# -*- coding: UTF-8 -*-

import pymysql
from email import charset  # @UnusedImport
import xlwt
import xlrd
from xlutils.copy import copy

def setOutCell(outSheet, col, row, value):
    """ Change cell value without changing formatting. """
    def _getOutCell(outSheet, colIndex, rowIndex):
        """ HACK: Extract the internal xlwt cell representation. """
        row = outSheet._Worksheet__rows.get(rowIndex)
        if not row: return None
 
        cell = row._Row__cells.get(colIndex)
        return cell
 
    # HACK to retain cell style.
    previousCell = _getOutCell(outSheet, col, row)
    # END HACK, PART I
 
    outSheet.write(row, col, value)
 
    # HACK, PART II
    if previousCell:
        newCell = _getOutCell(outSheet, col, row)
        if newCell:
            newCell.xf_idx = previousCell.xf_idx

def importFunction1(row,col,s):
    for row in range(7,20):
        Sheet1.write(row,col,xlwt.Formula('Sheet3!'+s+str(row-5)))
        row = row + 1

def importData(sql,num):
    cur.execute(sql)
    #获取查询的所有记录
    results = cur.fetchall()
    # 获取MYSQL里面的数据字段名称
    fields = cur.description
    Sheet = newWb.get_sheet(num)
    #写入字段信息
    for field in range(0,len(fields)):
        print(field)
        Sheet.write(0,field,fields[field][0])
    row = 1
    col = 0
    #写入数据段信息
    for row in range(1,len(results)+1):
        for col in range(0,len(fields)):
            Sheet.write(row,col,u'%s'%results[row-1][col])      
        
#打开数据库连接  
db= pymysql.connect(host="localhost",user="root",  
    password="root",db="cti",port=3306,charset="utf8")

# 使用cursor()方法获取操作游标  
cur = db.cursor() 
  
# 编写sql 
sql1 = "SET @b_time = '%s'"
sql2 = "SET @e_time = '%s'"
sql3 = "SET @namespace = '%s'"
sql4 = "SET @_namespace = '%s'"
sql6 = "SET @tenant_id = '%s'"
sql7 = "SET @activity_id = '%s'"
sql8 = "select bpo.description,sum(rpt.N_Agent) as '座席数', sum(rpt.n_dial_ob) as '呼叫量',sum(rpt.n_talk_ob) as '接通量',sum(rpt.t_talk_ob) as '通话时长', sum(rpt.n_dial_po) as '自动外呼总量',sum(rpt.n_talk_po) as '自动外呼接通',sum(rpt.t_talk_po) as '自动外呼时长' FROM  ( select q.queue_id queueId,COUNT(DISTINCT t.first_agent_no) N_Agent,0 n_dial_ob,0 n_dial_nr_ob,0 n_talk_ob,0 n_talk_nr_ob,0 t_talk_ob, 0 n_talk_ob_a,0 n_talk_nr_ob_a,0 t_talk_ob_a, 0 n_dial_po,0 n_dial_nr_po,0 n_talk_po,0 t_talk_po, 0 n_talk_po_a,0 t_talk_po_a from fact_call t,          (SELECT id,CONCAT(login_name,@_namespace) an,org_id from t_user WHERE namespace=@namespace) u,         (SELECT queue_id,agent_id FROM t_agent_queue WHERE agent_id in (SELECT id FROM t_user WHERE tenant_id = @tenant_id)) q where t.tenant_id=@tenant_id AND t.begin_time BETWEEN @b_time AND @e_time  AND t.first_agent_no = u.an AND u.id = q.agent_id GROUP BY q.queue_id UNION ALL select q.queue_id queueId, 0 N_Agent,COUNT(t.call_id) n_dial_ob,COUNT(DISTINCT t.dnis) n_dial_nr_ob,0 n_talk_ob,0 n_talk_nr_ob,0 t_talk_ob, 0 n_talk_ob_a,0 n_talk_nr_ob_a,0 t_talk_ob_a, 0 n_dial_po,0 n_dial_nr_po,0 n_talk_po,0 t_talk_po, 0 n_talk_po_a,0 t_talk_po_a from fact_call t,          (SELECT id,CONCAT(login_name,@_namespace) an,org_id from t_user WHERE namespace=@namespace) u,         (SELECT queue_id,agent_id FROM t_agent_queue WHERE agent_id in (SELECT id FROM t_user WHERE tenant_id = @tenant_id)) q where t.tenant_id=@tenant_id AND t.begin_time BETWEEN @b_time AND @e_time   AND t.call_type=1      AND t.first_agent_no = u.an AND u.id = q.agent_id GROUP BY q.queue_id UNION ALL select q.queue_id queueId, 0 N_Agent,0 n_dial_ob,0 n_dial_nr_ob, COUNT(t.call_id) n_talk_ob,COUNT(DISTINCT t.dnis) n_talk_nr_ob,sum(t.talk_length) t_talk_ob, 0 n_talk_ob_a,0 n_talk_nr_ob_a,0 t_talk_ob_a,  0 n_dial_po,0 n_dial_nr_po,0 n_talk_po,0 t_talk_po, 0 n_talk_po_a,0 t_talk_po_a from fact_call t,          (SELECT id,CONCAT(login_name,@_namespace) an,org_id from t_user WHERE namespace=@namespace) u,         (SELECT queue_id,agent_id FROM t_agent_queue WHERE agent_id in (SELECT id FROM t_user WHERE tenant_id = @tenant_id)) q where t.tenant_id=@tenant_id AND t.begin_time BETWEEN @b_time AND @e_time   AND t.call_type=1   AND t.talk_length>0      AND t.first_agent_no = u.an AND u.id = q.agent_id GROUP BY q.queue_id UNION ALL select q.queue_id queueId, 0 N_Agent,0 n_dial_ob,0 n_dial_nr_ob,0 n_talk_ob,0 n_talk_nr_ob,0 t_talk_ob,  COUNT(t.call_id) n_talk_ob_a,COUNT(DISTINCT t.dnis) n_talk_nr_ob_a,sum(t.talk_length) t_talk_ob_a, 0 n_dial_po,0 n_dial_nr_po,0 n_talk_po,0 t_talk_po, 0 n_talk_po_a,0 t_talk_po_a from fact_call t,  (SELECT id,CONCAT(login_name,@_namespace) an,org_id from t_user WHERE namespace=@namespace) u,         (SELECT queue_id,agent_id FROM t_agent_queue WHERE agent_id in (SELECT id FROM t_user WHERE tenant_id = @tenant_id)) q     where t.tenant_id=@tenant_id AND t.begin_time BETWEEN @b_time AND @e_time   AND t.call_type=1   AND t.talk_length>30      AND t.first_agent_no = u.an AND u.id = q.agent_id GROUP BY q.queue_id UNION ALL SELECT q.id queueId,0 N_Agent,0 n_dial_ob,0 n_dial_nr_ob,0 n_talk_ob,0 n_talk_nr_ob,0 t_talk_ob, 0 n_talk_ob_a,0 n_talk_nr_ob_a,0 t_talk_ob_a, count(a.call_id) n_dial_po,0 n_dial_nr_po,0 n_talk_po,0 t_talk_po, 0 n_talk_po_a,0 t_talk_po_a FROM t_call_result a,(SELECT CONCAT(login_name, @_namespace) an,org_id from t_user WHERE namespace=@namespace) u, (SELECT `code`,id FROM t_queue where tenant_id = @tenant_id) q     WHERE a.activity_id = @activity_id    AND a.call_time BETWEEN @b_time AND @e_time     AND a.agent=u.an     AND a.skill = q.`code` GROUP BY q.id UNION ALL SELECT q.id queueId,0 N_Agent,0 n_dial_ob,0 n_dial_nr_ob,0 n_talk_ob,0 n_talk_nr_ob,0 t_talk_ob, 0 n_talk_ob_a,0 n_talk_nr_ob_a,0 t_talk_ob_a, 0 n_dial_po,count(DISTINCT a.rosterinfo_id) n_dial_nr_po,0 n_talk_po,0 t_talk_po, 0 n_talk_po_a,0 t_talk_po_a FROM t_call_result a, (SELECT CONCAT(login_name, @_namespace) an,org_id from t_user WHERE namespace=@namespace) u, (SELECT `code`,id FROM t_queue where tenant_id = @tenant_id) q     WHERE a.activity_id = @activity_id    AND a.call_time BETWEEN @b_time AND @e_time   AND a.agent=u.an     AND a.skill = q.`code` GROUP BY q.id UNION ALL SELECT q.id queueId,0 N_Agent,0 n_dial_ob,0 n_dial_nr_ob,0 n_talk_ob,0 n_talk_nr_ob,0 t_talk_ob, 0 n_talk_ob_a,0 n_talk_nr_ob_a,0 t_talk_ob_a, 0 n_dial_po,0 n_dial_nr_po,count(a.call_id) n_talk_po,sum(f.talk_length) t_talk_po, 0 n_talk_po_a,0 t_talk_po_a FROM t_call_result a, fact_call f,(SELECT CONCAT(login_name, @_namespace) an,org_id from t_user WHERE namespace=@namespace) u, (SELECT `code`,id FROM t_queue where tenant_id = @tenant_id) q     WHERE a.activity_id = @activity_id    AND a.call_time BETWEEN @b_time AND @e_time     AND a.result = 0   AND a.call_id = f.call_id   and f.call_type= 4    AND a.agent = u.an     AND a.skill = q.`code` GROUP BY q.id UNION ALL SELECT q.id queueId,0 N_Agent,0 n_dial_ob,0 n_dial_nr_ob,0 n_talk_ob,0 n_talk_nr_ob,0 t_talk_ob, 0 n_talk_ob_a,0 n_talk_nr_ob_a,0 t_talk_ob_a, 0 n_dial_po,0 n_dial_nr_po,0 n_talk_po,0 t_talk_po, count(a.call_id) n_talk_po_a,sum(f.talk_length) t_talk_po_a FROM t_call_result a, fact_call f,(SELECT CONCAT(login_name, @_namespace) an,org_id from t_user WHERE namespace=@namespace) u, (SELECT `code`,id FROM t_queue where tenant_id = @tenant_id) q     WHERE a.activity_id = @activity_id    AND a.call_time BETWEEN @b_time AND @e_time     AND a.result = 0   AND f.talk_length>30   AND a.call_id = f.call_id   and f.call_type= 4    AND a.agent = u.an     AND a.skill = q.`code` GROUP BY q.id ) rpt, (SELECT     t.id id,     t.description description FROM     t_queue t WHERE     tenant_id = @tenant_id) bpo WHERE rpt.queueId = bpo.id GROUP BY bpo.description ORDER BY bpo.description"  
sql9 = "select bpo.description,rpt.round,sum(rpt.n_dial_po) AS '自动外呼总量' from ( SELECT q.id queueId,a.round round, count(a.call_id) n_dial_po FROM t_call_result a,(SELECT CONCAT(login_name, @_namespace) an,org_id from t_user WHERE namespace=@namespace) u, (SELECT `code`,id FROM t_queue where tenant_id = @tenant_id) q     WHERE a.activity_id = @activity_id    AND a.call_time BETWEEN @b_time AND @e_time     AND a.agent=u.an     AND a.skill = q.`code` GROUP BY q.id,a.round ) rpt, (SELECT     t.id id,     t.description description FROM     t_queue t WHERE     tenant_id = @tenant_id) bpo WHERE rpt.queueId = bpo.id GROUP BY bpo.description,rpt.round ORDER BY bpo.description,rpt.round"
sql10 = "select bpo.description,rpt.round,sum(rpt.n_talk_po) as '接通量',sum(rpt.t_talk_po) as '通话时长' from ( SELECT q.id queueId,a.round round,count(a.call_id) n_talk_po,sum(f.talk_length) t_talk_po FROM t_call_result a, fact_call f,(SELECT CONCAT(login_name, @_namespace) an,org_id from t_user WHERE namespace=@namespace) u, (SELECT `code`,id FROM t_queue where tenant_id = @tenant_id) q     WHERE a.activity_id = @activity_id    AND a.call_time BETWEEN @b_time AND @e_time     AND a.result = 0   AND a.call_id = f.call_id   and f.call_type= 4    AND a.agent = u.an     AND a.skill = q.`code` GROUP BY q.id,a.round ) rpt, (SELECT     t.id id,     t.description description FROM     t_queue t WHERE     tenant_id = @tenant_id) bpo WHERE rpt.queueId = bpo.id GROUP BY bpo.description,rpt.round ORDER BY bpo.description,rpt.round"
try:  
    cur.execute(sql1 % ('2018-01-10'))    #执行sql语句  
    cur.execute(sql2 % ('2018-02-28'))
    cur.execute(sql3 % ('zln.cc'))
    cur.execute(sql4 % ('@zln.cc'))
    cur.execute(sql6 % ('7804b1dd-81cd-44a7-bdff-4061b0c18a6d'))
    cur.execute(sql7 % ('f047c4640b1611e88bfa0050569505de'))
    
    #open existed xls file  
    oldWb = xlrd.open_workbook(r'C:/Users/89232/Desktop/readout.xls',formatting_info=True)  
    newWb = copy(oldWb) 
    
    Sheet1 = newWb.get_sheet(0)
    #Sheet1写入函数(从Sheet3)
    importFunction1(7, 1, 'A')
    importFunction1(7, 2, 'B')
    importFunction1(7, 7, 'C')
    importFunction1(7, 9, 'D')
    importFunction1(7, 12, 'E')
    importFunction1(7, 34, 'H')
    
    #向Sheet3导入数据
    importData(sql8, 1)
      
    #向Sheet4导入数据
    importData(sql9, 2)       
    #importData(sql10, 2)
    
    
    cur.execute(sql10)
    #获取查询的所有记录
    results = cur.fetchall()
    # 获取MYSQL里面的数据字段名称
    fields = cur.description
    Sheet = newWb.get_sheet(2)
    #写入字段信息
    
    for field in range(2,len(fields)):
        print(fields[field][0])
        Sheet.write(0,field+1,fields[field][0])
    
    #写入数据段信息
    for row in range(1,len(results)+1):
        for col in range(2,len(fields)):
            Sheet.write(row,col+1,u'%s'%results[row-1][col]) 
    
           
           
    print("write new values ok") 
    newWb.save('C:/Users/89232/Desktop/readout3.xls')
    print("save with same name ok")
      
except Exception as e:  
    raise e  
finally:  
    db.close()  #关闭连接    