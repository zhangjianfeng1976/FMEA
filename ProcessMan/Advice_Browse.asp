<!--#include file="../includes/db.asp"-->
<%
Select Case Request("OK") 
       Case "搜索" 	                
	     Response.Redirect "Advice_Browse.asp?Cat=0&Sc="&Request("Search")
       Case "全机种搜索" 	                
	     Response.Redirect "Advice_Browse.asp?Cat=1&Sc="&Request("Search")
       Case "已对应项目" 	                
	     Response.Redirect "Advice_Browse.asp?Cat=3"
       Case "未改善项目一览" 	                
	     Response.Redirect "Advice_BrInput.asp"
       Case "＜＜＜"
             if request("page")= "" then
               Response.Redirect "Process_Browse.asp?page=1"
             else 	                
	       Response.Redirect "Process_Browse.asp?page="&request("page")
             end if
       Case "修订"

             Sev = int(request("Sev"))
             Occ = int(request("Occ"))
             Det = int(request("Det"))
             RPN = round((Sev * Occ * Det)^(1/3),1)
                 
            Sql_chg ="Update t_Fmea_Advice set "
            Sql_chg = Sql_chg & "AdvContent = '"& request("AdvContent") & "',"
            Sql_chg = Sql_chg & "Rspnser = '"& request("Rspnser") & "',"
            Sql_chg = Sql_chg & "FshDate = '"& request("FshDate") & "',"
            Sql_chg = Sql_chg & "ActionContent = '"& request("ActionContent") & "',"
            Sql_chg = Sql_chg & "ResultS = '"& request("Sev") & "',"
            Sql_chg = Sql_chg & "ResultO = '"& request("Occ") & "',"
            Sql_chg = Sql_chg & "ResultD = '"& request("Det") & "',"
            Sql_chg = Sql_chg & "ResultRPN = '"& RPN & "'"
            Sql_chg = Sql_chg & " Where AdID = '"& request("AdID") & "'"
            call openDB()
            conn.execute(Sql_chg)
            call closeDB()
            Response.Redirect "Process_Browse.asp"
       
       Case "升版"
             
            sql = "Insert into t_Fmea_Advice_history "
            sql = sql & "(AdID,ItmID,PjID,Pjkey,AdvContent,Rspnser,FshDate,"
            sql = sql & "ActionContent,ResultS,ResultO,ResultD,"
            sql = sql & "ResultRPN,AdvStatus,InUser,InTime)"
            sql = sql & " Select AdID,ItmID,PjID,Pjkey,AdvContent,Rspnser,FshDate,"
            sql = sql & "ActionContent,ResultS,ResultO,ResultD,"
            sql = sql & "ResultRPN,'N',InUser,InTime"
            sql = sql & " from t_Fmea_Advice Where AdID = '"& request("AdID") & "'"
             
            Sev = int(request("Sev"))
            Occ = int(request("Occ"))
            Det = int(request("Det"))
            RPN = round((Sev * Occ * Det)^(1/3),1)
                 
            SqlStr_c ="Update t_Fmea_Advice set "
            SqlStr_c = SqlStr_c & "AdvContent = '"& request("AdvContent") & "',"
            SqlStr_c = SqlStr_c & "Rspnser = '"& request("Rspnser") & "',"
            SqlStr_c = SqlStr_c & "FshDate = '"& request("FshDate") & "',"
            SqlStr_c = SqlStr_c & "ActionContent = '"& request("ActionContent") & "',"
            SqlStr_c = SqlStr_c & "ResultS = '"& request("Sev") & "',"
            SqlStr_c = SqlStr_c & "ResultO = '"& request("Occ") & "',"
            SqlStr_c = SqlStr_c & "ResultD = '"& request("Det") & "',"
            SqlStr_c = SqlStr_c & "ResultRPN = '"& RPN & "'"
            SqlStr_c = SqlStr_c & " Where AdID = '"& request("AdID") & "'"
            
            call openDB()
            conn.BeginTrans
            conn.execute(sql)
            conn.execute(SqlStr_c)           
            conn.CommitTrans
            call closeDB()
	
            Response.Redirect "Process_Browse.asp"

         Case "废止"

            sql_cl = "Insert into t_Fmea_Advice_history "
            sql_cl = sql_cl & "(AdID,ItmID,PjID,Pjkey,AdvContent,Rspnser,FshDate,"
            sql_cl = sql_cl & "ActionContent,ResultS,ResultO,ResultD,"
            sql_cl = sql_cl & "ResultRPN,AdvStatus,InUser,InTime)"
            sql_cl = sql_cl & " Select AdID,ItmID,PjID,Pjkey,AdvContent,Rspnser,FshDate,"
            sql_cl = sql_cl & "ActionContent,ResultS,ResultO,ResultD,"
            sql_cl = sql_cl & "ResultRPN,'CL',InUser,InTime"
            sql_cl = sql_cl & " from t_Fmea_Advice Where AdID = '"& request("AdID") & "'"
             
            SqlStr_d = "delete t_Fmea_Advice where AdID ="&request("AdID") 

            call openDB()
            conn.BeginTrans
            conn.execute(sql_cl)
            conn.execute(SqlStr_d)           
            conn.CommitTrans
            call closeDB()
	
            Response.Redirect "Process_Browse.asp"		             
End Select

pgs =2500
select case request("cat")
       case "" 
         ttl = request.cookies("fmea.Pjkey")
         sql="SELECT * FROM v_Fmea_total where Pjkey = '"&request.cookies("fmea.Pjkey")&"' order by SubPjName,PageNo,WorkQueue"
         pgs = 10
       case 0
          ttl = request.cookies("fmea.Pjkey")
          sql = "SELECT * FROM v_Fmea_total where "
          sql = sql & "(PartsNo like '%" & request("sc") & "%' or "
          sql = sql & "PartsName like '%" & request("sc") & "%' or "
          sql = sql & "AdvContent like '%" & request("sc") & "%' or "
          sql = sql & "Rspnser like '%" & request("sc") & "%' or "
          sql = sql & "ActionContent like '%" & request("sc") & "%'"
          sql = sql & ") and Pjkey = '"&request.cookies("fmea.Pjkey")&"' and  AdID is not null"
       case 1
          ttl = "全部"
          sql = "SELECT * FROM v_Fmea_total where "
          sql = sql & "(PartsNo like '%" & request("sc") & "%' or "
          sql = sql & "PartsName like '%" & request("sc") & "%' or "
          sql = sql & "AdvContent like '%" & request("sc") & "%' or "
          sql = sql & "Rspnser like '%" & request("sc") & "%' or "
          sql = sql & "ActionContent like '%" & request("sc") & "%'"
          sql = sql & ") order by ItmID"
       case 2
         ttl = request.cookies("fmea.Pjkey")
         sql="SELECT * FROM v_Fmea_total WHERE ItmID = "& request("ItmID")  
       case 3
         ttl = request.cookies("fmea.Pjkey")
         sql="SELECT * FROM v_Fmea_total WHERE Pjkey = '"&request.cookies("fmea.Pjkey")&"' and  AdID is not null order by SubPjName,PageNo,WorkQueue"     
end select

call openDB()

rs.open sql,conn,1,1
if not rs.eof then
rs.pagesize =pgs
pagecount=rs.PageCount 
page=int(request.QueryString ("page"))
if page<=0 then page=1
if request.QueryString("page")="" then page=1
rs.AbsolutePage=page 
end if
%>

<html>
<head>
<script language="javascript" type="text/javascript">
<!--
function jump() 
{
 if(event.keyCode==13)event.keyCode=9
}
-->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>机种改善措施项目列表</title>
</head>

<body bgcolor ="#FCFCFC">
<form action="Advice_Browse.asp">
<input type="hidden" value="<% = request("page") %>" name="page">
<div style ="float:right ;width：200px;hight:200px;">
<input type="submit" value="未改善项目一览" name="OK" class="smallInput">
<input name="Search" type="text" class="smallInput" size="40">
<input type="submit" value="搜索" name="OK" class="smallInput">
<input type="submit" value="全机种搜索" name="OK" class="smallInput">
<input type="submit" value="已对应项目" name="OK" class="smallInput">
<input type="submit" value="＜＜＜" name="OK" class="smallInput">
&nbsp;&nbsp;
</div>

<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr><td width="75%" colspan=14 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
[<a href="../desktop.asp">回到首页</a>][<a href="Process_Browse.asp">回上一级</a>][<a href="javascript:history.back()">返回</a>]
</td><td  width="50%" align="right">
<% call clspage() %>   
</table>
</td></tr>

<tr bgcolor="#E6F7FF"><td height="30" align="center" colspan=14><b><% = ttl %>改善措施列表</b></td></tr>
 
<tr bgcolor="#FCFCFC" align="center">
<td width="2%"  rowspan ="2">顺序</td>
<td width="8%" height="20" rowspan ="2">工程名称</td>
<td width="4%"  rowspan ="2">标准书页码</td>
<td width="6%"  rowspan ="2">部品编号</td>
<td width="8%"  rowspan ="2">部品名称</td>
<td width="15%" rowspan ="2">建议措施</td>
<td width="5%"  rowspan ="2">责任者</td>
<td width="8%"  rowspan ="2">目标完成日</td>
<td width="18%" colspan ="5">措施结果</td>
<%
if request.cookies("fmea.power") > "1" then
   response.write "<td width='12%' rowspan ='2'>操作</td>"
end if
%>
</tr><tr bgcolor="#FCFCFC" align="center">
<td width="10%">采用措施</td>
<td width="2%">S</td>
<td width="2%">O</td>
<td width="2%">D</td>
<td width="2%">RI</td>
</tr>

<% 
if request("act") = "del" then
        SqlStr_d = "delete t_Fmea_Advice where AdID ="&request("AdID")     
        conn.execute(SqlStr_d)
        call closeDB()
        Response.Redirect "Process_Browse.asp"
elseif request("Act") = "Edit" then
        SqlStr_e = "SELECT * FROM t_Fmea_Advice where AdID ="&request("AdID")
        Set rs_ed = Server.CreateObject("ADODB.Recordset")        
        rs_ed.open SqlStr_e,conn,1,1
        response.write"<tr bgcolor='#FCFCFC'>"
        response.write"<input name='AdID' type='hidden' class='notline' size='5' value ='"&rs_ed("AdID")&"'>"
        response.write"<input name='ItmID' type='hidden' class='notline' size='5' value ='"&rs_ed("ItmID")&"'>"
        response.write"<td bgcolor='#F0F0F0'></td><td bgcolor='#F0F0F0'></td>"
        response.write"<td bgcolor='#F0F0F0'></td><td bgcolor='#F0F0F0'></td>"
        response.write"<td bgcolor='#F0F0F0'></td>"
	    response.write"<td><input name='AdvContent' class='notline' size='40' value ='"&rs_ed("AdvContent")&"'></td>"
        response.write"<td><input name='Rspnser' class='notline' size='10' value ='"&rs_ed("Rspnser")&"'></td>"
	    response.write"<td><input name='FshDate' class='notline' size='15' value ='"&rs_ed("FshDate")&"'></td>"
        response.write"<td><input name='ActionContent' class='notline' size='25' value ='"&rs_ed("ActionContent")&"'></td>"
        response.write"<td><input name='Sev' class='notline' size='3' value ='"&rs_ed("ResultS")&"'></td>"
        response.write"<td><input name='Occ' class='notline' size='3' value ='"&rs_ed("ResultO")&"'></td>"
        response.write"<td><input name='Det' class='notline' size='3' value ='"&rs_ed("ResultD")&"'></td>"
        response.write"<td bgcolor='#F0F0F0'></td><td bgcolor='#F0F0F0'></td></tr>"
        response.write"<tr bgcolor='#FCFCFC'><td height = '30' colspan = '14' align='center'>"
        response.write"<input name='OK' value='修订' type='submit' class='smallInput'>&nbsp;"
        response.write"<input name='OK' value='升版' type='submit' class='smallInput'>&nbsp;" 
        response.write"<input name='OK' value='废止' type='submit' class='smallInput'>&nbsp;"
        response.write"[<a href='javascript:history.back()'>返回</a>]" 
        response.write"</td></tr>"  
        response.write"<tr bgcolor='#FCFCFC'><td height = '30' colspan = '14' >"
        response.write"【修正】是不会留下原来的记录的，所以，只在录入错误时使用，要慎用。<br>"
        response.write"【升版】是会留下原来的记录的，可以在修订记录上查看，所以，是在文件修订时使用。修订完成后请在【机种列表】上进行升版。<br>"
        response.write"【废止】与升版一样，删除相关的内容，但会留下原来的记录的，这时，修订完成后也需要在【机种列表】上进行升版。<br>"
        response.write"</td></tr>"  
        call closeDB()
else

call MainBody()

 %>

</table>
</form>
<% call closeDB() %>
<% end if %>
</body>
</html>

<%
Sub MainBody()
If rs.eof then
	response.write "<tr><td colspan=6>"
	response.write "没有记录。"
	response.write "</td></tr>"
end if

n=1 
while not rs.eof and n<= rs.pagesize
    AdID = rs("AdID")
	ItmID = rs("ItmID")
	response.write"<tr bgcolor='#FCFCFC'>"
    response.write"<td>"&rs("WorkQueue")&"</td>"
    response.write"<td>"&rs("SubPjName")&"</td>"	
    response.write"<td>"&rs("PageNo")&"</td>"
    response.write"<td>"&rs("PartsNo")&"</td>"
    response.write"<td>"&rs("PartsName")&"</td>"	
    response.write"<td>"&rs("AdvContent")&"</td>"	
	response.write"<td>"&rs("Rspnser")&"</td>"
    response.write"<td>"&rs("FshDate")&"</td>"
    response.write"<td>"&rs("ActionContent")&"</td>"
    response.write"<td>"&rs("ResultS")&"</td>"
    response.write"<td>"&rs("ResultO")&"</td>"
    response.write"<td>"&rs("ResultD")&"</td>"
    response.write"<td>"&rs("ResultRPN")&"</td>" 
    if request.cookies("fmea.power") > "1" then
       response.write"<td>&nbsp;"
        if isnull(rs("AdID")) then
           response.write"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        else
           response.write"<a href=Advice_Browse.asp?Act=Edit&AdID="&AdID&">修订</a>&nbsp;&nbsp;"
        end if
        if isnull(rs("ItmID_histo")) then
            response.write"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        else
            response.write"<a href=../history/His_Advice_Browse.asp?itmID="&itmID&">"&"修订履历</a>&nbsp;&nbsp"
        end if
        vRPN = trim(rs("RPN"))
        if vRPN > 2 and isnull(rs("AdID")) Then
           response.write"<a href=Advice_add.asp?ItmID="&ItmID&"&PjID="&rs("PjID")&"&PjKey="&rs("PjKey")&">措施填入</a>&nbsp;&nbsp;"
        else
           response.write"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        end if
        response.write"</td>"
    end if
	response.write"</tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            '没有上一页，但是有下一页'
Response.Write "首页&nbsp;上页&nbsp;"
Response.Write "<a href=Advice_Browse.asp?page="&page+1&">下页</a>&nbsp;"
Response.Write "<a href=Advice_Browse.asp?page="&pagecount&">尾页</a>"

elseif page=pagecount and not page=1  then         '没有下一页，但是有上一页' 
Response.Write "<a href=Advice_Browse.asp?page=1>首页</a>&nbsp;&nbsp;"
Response.Write "<a href=Advice_Browse.asp?page="&page-1&">上页</a>&nbsp;&nbsp;下页&nbsp;尾页"

elseif page<1 or page>pagecount then '没有任何记录' 
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

elseif page=1 and page=pagecount then '没有上一页，没有下一页'
Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

else
Response.Write "<a href=Advice_Browse.asp?page=1>首页</a>&nbsp;"
Response.Write "<a href=Advice_Browse.asp?page="&page-1&">上页</a>&nbsp;"
Response.Write "<a href=Advice_Browse.asp?page="&page+1&">下页</a>&nbsp;"
Response.Write "<a href=Advice_Browse.asp?page="&pagecount&">尾页</a>"
end if 
Response.Write  "&nbsp; 第"&page&"页&nbsp;共"&pagecount&"页 计"&rs.recordcount&"条记录"
end sub
%>