<!--替位记录追加-->
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
wkname             = rs("姓名")
Changedate         = rs("担当日")
Wkstatus           = rs("有效")
call closeDB()

Select Case Request("OK") 
      Case "确定" 
        if request("CReason") =""  or request("CardID") = "" or len(request("CardID"))<> 6 then
        response.write "空格未填完成[<a href='javascript:history.back()'>返回</a>]" 
        response.end
        end if  
        if  request("CardID") <> "999999" and request("WKcontent") ="" then
        response.write "请选择作业内容[<a href='javascript:history.back()'>返回</a>]" 
        response.end
        end if 
       dim  pn
       pn=year(request("dtae"))&right("0"&month(request("dtae")),2)
       sql= "Insert into wt_wdairy(dtae,pno,lineid,creason,cardid,wkcontent,leader,truename,intime) values ('" _
              &request("dtae")&"','" _
              &pn&"','" _
              &request("lineid")&"','" _
              &request("Creason")&"','" _
              &request("CardID")&"','" _
              &request("WKcontent")&"','" _
              &request.cookies("leader")&"','"&Request.ServerVariables("REMOTE_HOST")&"','"&NOW()&"')" 
      call openDB()
      conn.execute(sql)
      call closeDB()
      aa = request.cookies("c_Sc")
      if request("Creason") = "支援" then
      Response.Redirect "lcbrowse.asp?cat=0&sc="&request("CardID")
      else 
      Response.Redirect "lcbrowse.asp?cat=0&sc="&aa
      end if
End Select
%>
<html>
<head>
<title>工位人员管理</title>
<script>
function aa(){
var a = document.getElementById("cardid").value
av = a
}</script>
</head>

<body>
<form action="LcAdd.asp" method="POST" >
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=11>
<b>替位变更</b>
</td></tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="8%" height="25">Line</td>
<td width="8%">机种</td>
<td width="8%">工程</td>
<td width="4%">工位</td>
<td width="8%">工号</td>
<td width="8%">姓名</td>
<td width="8%">有效</td>
<td width="8%">担当日</td>
</tr>
<tr bgcolor="#FCFCFC" align="center">
<input type="hidden" value="<%=Lineid %>" name="lineid">
<td width="8%" height="22"><%=Line %></td>
<td width="8%"><%=Model %></td>
<td width="4%"><%=WoGroup %></td>
<td width="8%"><%=WorkStation %></td>
<td width="8%"><%=OriginWorker %></td>
<td width="8%"><%=Wkname %></td>
<td width="8%"><%=Wkstatus %></td>
<td width="8%"><%=Changedate %></td>
</tr><tr bgcolor="#FCFCFC">
<td width="8%" height="25">被替位原因</td>
<td width="8%" colspan = 7 >
<input type="radio" name="CReason" value="年假">年假&nbsp;&nbsp; 
<input type="radio" name="CReason" value="缺勤">缺勤&nbsp;&nbsp;
<input type="radio" name="CReason" value="病假">病假&nbsp;&nbsp;
<input type="radio" name="CReason" value="产假">产假&nbsp;&nbsp;
<input type="radio" name="CReason" value="婚假">婚假&nbsp;&nbsp;
<input type="radio" name="CReason" value="丧假">丧假&nbsp;&nbsp;
<input type="radio" name="CReason" value="生育假">生育假&nbsp;&nbsp;
<input type="radio" name="CReason" value="离职">离职&nbsp;&nbsp;<br>
<input type="radio" name="CReason" value="预离">预离&nbsp;&nbsp;
<input type="radio" name="CReason" value="预加减位">预加减位&nbsp;&nbsp;
<input type="radio" name="CReason" value="支援">支援（当日有出勤，到其他工位替位或学位）&nbsp;&nbsp;
<input type="radio" name="CReason" value="确认期">确认期（做了原作业者变更，但仍在确认期内）&nbsp;&nbsp;

</td>
</tr><tr  bgcolor="#FCFCFC">
<td  height="22">日期</td>
<td colspan = 7>
<input class="notline" name="Dtae"  size="10" value="<% =date() %>"
</td>
</tr><tr  bgcolor="#FCFCFC">
<td  height="22">安排为</td>
<td colspan =7 >
担当者工号：<input name="Cardid" type="text" class="smallInput"  size="17" id="cardid" onkeyup="aa()" onblur="aa()">
</td>
</tr><tr  bgcolor="#FCFCFC">
<td>作业内容</td>
<td colspan =7  height="22">
<input type="radio" name="WKcontent" value="学位">学位&nbsp;&nbsp;
<input type="radio" name="WKcontent" value="替位">替位&nbsp;&nbsp;
</td></tr>
</tr><tr  bgcolor="#FCFCFC">
<td  height="22" colspan =8 align ="center"><input type="submit" value="确定" name="OK" class="smallinput"></td>
</tr><tr  bgcolor="#FCFCFC">
<td  height="22" colspan =8 >
说明：<br>
1.由于之前用替位，学位等字眼，在记录上无法区分是前面的人员还是后面的人员，所以，在表格上做了明确。<br>
2.如果工位的人员当日有出勤，但到了别工位进行替位或学位了，这时，统一用支援进行表达，而安排的人员的作业，就用学位或替位来表示。<br>
3.如果支援人员的工位不需要再替位的话，用999999表示的话，自然地不需要选择学位及替位了。<br>
4.预离：工位担当预定离职，安排员工进入学位的情况。<br>
5.预加减位：LINE打算加减位，在两个工位中插入学位的情况。 <br>
6.确认期是指己经作了原作业者变更，但还在确认期需要继续确认的人员。
</td>
</tr>



<!--
<tr><td colspan =8>
<a href="lc_incheck.asp?cat=0&sc="&av  target ="container1">检证人员</a>
</td></tr>
<tr><td colspan =8>
<iframe src='' name='container1' width='100%' marginwidth='1' height='100' marginheight='1' frameborder='0' scrolling='none'></iframe>
</td></tr>
 --> 
</table>
</form>
</body>
</html>