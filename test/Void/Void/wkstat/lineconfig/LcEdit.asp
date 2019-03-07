<!--#include file="../includes/db.asp"-->
<%
if request.cookies("leader")="" then
response.redirect "../default.asp"
end if

LineID=request("LineID")
sql="SELECT * FROM v_wmaster where LineID="&LineID
call openDB()
rs.open sql,conn,1,1
LineID             = rs("LineID")
Line               = rs("Line")
Model              = rs("机种")
WoGroup            = rs("工程")
WorkStation        = rs("工位")
OriginWorker       = rs("工号")
Wkname             = rs("姓名")
Changedate         = rs("担当日")
Wkstatus           = rs("有效")
stat               = trim(rs("版本"))
call closeDB()

Select Case Request("OK") 
      Case "确定"
       if  request("originworker") = "999999" or request("originworker") = "" or len(request("originworker"))<> 6 then
        response.write "空格未填完成或填写错误[<a href='javascript:history.back()'>返回</a>]" 
        response.end
       end if 
       dtae =date()
       dim  statuspn
       statuspn=year(now())&right("0"&month(now()),2)&right("0"&day(now()),2)
       'sta = stat&"a"&year(now())&right("0"&month(now()),2)&right("0"&day(now()),2)
       'sql = "Update wt_lineconfig set wkstatus = 'N',statuspno='"&sta&"' where LineId = "&LineID
       sql = "Update wt_lineconfig set wkstatus = 'N' where LineId = "&LineID
       sql1= "Insert into wt_lineconfig(line,model,wogroup,workstation,originworker,changedate,wkstatus,statuspno,leader,truename,intime) values ('" _
              &request("line")&"','" _
              &request("model")&"','" _
              &request("wogroup")&"','" _
              &request("workstation")&"','" _
              &request("originworker")&"','" _
              &dtae&"','Y','"&statuspn&"','"&Request.Cookies("leader")&"','"&Request.ServerVariables("REMOTE_HOST")&"','"&NOW()&"')" 
      call openDB()
	 conn.BeginTrans
	 conn.execute(sql)
	 conn.execute(sql1)
	 conn.CommitTrans
      call closeDB()
      Response.Redirect "lcbrowse_ld1.asp"
 Case "缩位"
       sql2 = "Update wt_lineconfig set wkstatus = 'N' where LineId = "&LineID
       call openDB()
	 conn.execute(sql2)
       call closeDB()
      Response.Redirect "lcbrowse_ld1.asp"
End Select



%>
<html>
<head>
<title>工位人员管理</title>
</head>

<body>
<form action ="LcEdit.asp"  method="POST">
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=11>
<b>原工位作业者变更</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="8%" height="25">Line</td>
<td width="8%">机种</td>
<td width="8%">工程</td>
<td width="4%">工位</td>
<td width="8%">工号</td>
<td width="8%">姓名</td>
<td width="8%">有效</td>
<td width="8%">担当日</td>
<td width="30%" align="center">变更安排</td>
</tr>
<tr bgcolor="#FCFCFC" align="center">
<input type="hidden" value="<%=lineid %>" name="lineid">
<td width="8%" height="25"><input type="hidden" value="<%=line%>" name="line"><%=Line%></td>
<td width="8%"><input type="hidden" value="<%=Model%>" name="Model"><%=Model%></td>
<td width="4%"><input type="hidden" value="<%=WoGroup%>" name="WoGroup"><%=WoGroup%></td>
<td width="8%"><input type="hidden" value="<%=WorkStation%>" name="WorkStation"><%=WorkStation%></td>
<td width="8%"><%=OriginWorker%></td>
<td width="8%"><%=Wkname%></td>
<td width="8%"><%=Wkstatus%></td>
<td width="8%"><%=Changedate%></td>
<td width="8%">变更为担当者工号
<input name="originworker" type="text" class="smallInput"  size="17">
<input type="submit" value="确定" name="OK" class="smallinput">
<input type="submit" value="缩位" name="OK" class="smallinput">
</td>
</tr>
</table>
</form>
</body>
</html>