<!--#include file="../includes/db.asp"-->
<%
if (request.cookies("fmea.Pjkey") ="" or  request.cookies("fmea.Pjkey") <> request("Pjkey")) and  request("PjKey") <>"" then  
response.cookies("fmea.Pjkey") = request("Pjkey")
end if 

if (request.cookies("fmea.PjID") ="" or  request.cookies("fmea.PjID") <> request("PjID")) and  request("PjID") <>"" then  
response.cookies("fmea.PjID") = request("PjID")
end if 

Select Case Request("OK") 
       Case "全机种搜索" 	                
	     Response.Redirect "Process_Browse.asp?Cat=0&Sc="&Request("Search")	

       Case "搜索" 	                
	     Response.Redirect "Process_Browse.asp?Cat=1&Sc="&Request("Search")
       
       Case "需对应项目" 	                
	     Response.Redirect "Process_Browse.asp?Cat=2"
    
       Case "全系列搜索" 	                
	     Response.Redirect "Process_Browse.asp?Cat=3&Sc="&Request("Search")
     
       Case "评价录入"
             Response.Redirect "Process_Add.asp"
 
       Case "措施录入"
             Response.Redirect "Advice_BrInput.asp"  

       Case "对应措施列表"
             if request("page") ="" then
               Response.Redirect "Advice_Browse.asp?page=1"
             else
               Response.Redirect "Advice_Browse.asp?page="&request("page")      
             end if
       Case "修正"
             Sev = int(request("Sev"))
             Occ = int(request("Occ"))
             Det = int(request("Det"))
             RPN = round((Sev * Occ * Det)^(1/3),1)
                 
            SqlStr_c ="Update t_Fmea_Item set "
            SqlStr_c = SqlStr_c & "WorkQueue = '"& request("WorkQueue") & "',"
            SqlStr_c = SqlStr_c & "SubPjName = '"& request("SubPjName") & "',"
            SqlStr_c = SqlStr_c & "PageNo = '"& request("PageNo") & "',"         
            SqlStr_c = SqlStr_c & "PartsNo = '"& request("PartsNo") & "',"
            SqlStr_c = SqlStr_c & "PartsName = '"& request("PartsName") & "',"
            SqlStr_c = SqlStr_c & "PartsType = '"& request("PartsType") & "',"
            SqlStr_c = SqlStr_c & "WorkCata = '"& request("WorkCata") & "',"
            SqlStr_c = SqlStr_c & "WorkContent = '"& request("WorkContent") & "',"
            SqlStr_c = SqlStr_c & "NGMode = '"& request("NGMode") & "',"
            SqlStr_c = SqlStr_c & "NGEffect = '"& request("NGEffect") & "',"
            SqlStr_c = SqlStr_c & "ProDesign = '"& request("ProDesign") & "',"
            SqlStr_c = SqlStr_c & "Sev = '"& request("Sev") & "',"
            SqlStr_c = SqlStr_c & "Occ = '"& request("Occ") & "',"
            SqlStr_c = SqlStr_c & "Det = '"& request("Det") & "',"
            SqlStr_c = SqlStr_c & "RPN = '"& RPN & "'"
            SqlStr_c = SqlStr_c & " Where ItmID = '"& request("ItmID") & "'"
            call openDB()
            conn.execute(SqlStr_c)
            call closeDB()
            Response.Redirect "Process_Browse.asp?Cat=2"

         Case "升版" 
            
            sql = "Insert into t_Fmea_Item_history "
            sql = sql & "(ItmID,PjID,Pjkey,SubPjName,PageNo,WorkQueue,PartsNo,PartsName,PartsType,WorkCata,"
            sql = sql & "WorkContent,NGMode,NGEffect,ProDesign,Sev,Occ,"
            sql = sql & "Det,RPN,ItmStatus,InUser,InTime)"
            sql = sql & " Select ItmID, PjID ,Pjkey,SubPjName,PageNo,WorkQueue,PartsNo,"
            sql = sql & "PartsName,PartsType,WorkCata,WorkContent,NGMode,NGEffect,"
            sql = sql & "ProDesign,Sev,Occ,Det,RPN,'N',InUser,InTime"
            sql = sql & " from t_Fmea_Item Where ItmID = '"& request("ItmID") & "'"
            
            Sev = int(request("Sev"))
            Occ = int(request("Occ"))
            Det = int(request("Det"))
            RPN = round((Sev * Occ * Det)^(1/3),1)
            SqlStr_c ="Update t_Fmea_Item set "
            SqlStr_c = SqlStr_c & "WorkQueue = '"& request("WorkQueue") & "',"
            SqlStr_c = SqlStr_c & "SubPjName = '"& request("SubPjName") & "',"
            SqlStr_c = SqlStr_c & "PageNo = '"& request("PageNo") & "'," 
            SqlStr_c = SqlStr_c & "PartsNo = '"& request("PartsNo") & "',"
            SqlStr_c = SqlStr_c & "PartsName = '"& request("PartsName") & "',"
            SqlStr_c = SqlStr_c & "PartsType = '"& request("PartsType") & "',"
            SqlStr_c = SqlStr_c & "WorkCata = '"& request("WorkCata") & "',"
            SqlStr_c = SqlStr_c & "WorkContent = '"& request("WorkContent") & "',"
            SqlStr_c = SqlStr_c & "NGMode = '"& request("NGMode") & "',"
            SqlStr_c = SqlStr_c & "NGEffect = '"& request("NGEffect") & "',"
            SqlStr_c = SqlStr_c & "ProDesign = '"& request("ProDesign") & "',"
            SqlStr_c = SqlStr_c & "Sev = '"& request("Sev") & "',"
            SqlStr_c = SqlStr_c & "Occ = '"& request("Occ") & "',"
            SqlStr_c = SqlStr_c & "Det = '"& request("Det") & "',"
            SqlStr_c = SqlStr_c & "RPN = '"& RPN & "',"
            SqlStr_c = SqlStr_c & "InUser = '"& request.cookies("fmea.truename") & "',"
            SqlStr_c = SqlStr_c & "InTime = '"& now() & "'"
            SqlStr_c = SqlStr_c & " Where ItmID = '"& request("ItmID") & "'"
            
            call openDB()
            conn.BeginTrans
            conn.execute(sql)
            conn.execute(SqlStr_c)           
            conn.CommitTrans
            call closeDB()
            Response.Redirect "Process_Browse.asp"
          
           Case "废止" 
            Response.Write "你将要废止数据，相关记录会显示作业废止，要继续吗？（但不能用下载工具下载）<Br><Br>"
            Response.Write "【<a href='Process_del.asp?cat=1&ItmID="&request("ItmID")&"'>没问题，下手吧！</a>】"
            Response.Write "&nbsp;&nbsp;&nbsp;【<a href='javascript:history.go(-2)'>等等，我再确认一下</a>】" 
            Response.End         
                          
           Case "削除" 
            Response.Write "<font size='4' color ='ff0000'>你正在进行一项极度危险的操作，将要删除数据，记录会成为历史尘埃，真的要继续吗？</font><Br><Br>"
            Response.Write "【<a href='Process_del.asp?cat=2&ItmID="&request("ItmID")&"'>己决定了的啦，你再墨迹墨迹，我砸掉你</a>】"
            Response.Write "&nbsp;&nbsp;&nbsp;【<a href='javascript:history.go(-2)'>这样子，我还是再确认一下吧</a>】" 
            Response.End
            
          Case "GO"
             Response.Redirect "Process_Browse.asp?page="&Request("pagename")
End Select

pgs =2500
select case request("cat")
       case "" 
         ttl = request.cookies("fmea.PjKey")
         sql="SELECT * FROM t_Fmea_Item where PjID = "&request.cookies("fmea.PjID")
         pgs = 10
       case 0
         ttl ="全机种"
         sql="SELECT * FROM t_Fmea_Item where (SubPjName like '%"&request("Sc")& _
           "%' Or PageNo like '%"&request("Sc")& _
           "%' Or PartsType like '%"&request("Sc")& _
           "%' Or PartsNo like '%"&request("Sc")& _
           "%' Or PartsName like '%"&request("Sc")& _
           "%' Or WorkCata like '%"&request("Sc")& _
           "%' Or WorkContent like '%"&request("Sc")& _
           "%' Or NGMode like '%"&request("Sc")& _
           "%' Or NGEffect like '%"&request("Sc")& _
           "%' Or ProDesign like '%"&request("Sc")& _
           "%') ORDER BY SubPjName,PageNo" 
       case 1
         ttl = request.cookies("fmea.PjKey")
         sql="SELECT * FROM t_Fmea_Item where (SubPjName like '%"&request("Sc")& _
           "%' Or PageNo like '%"&request("Sc")& _
           "%' Or PartsNo like '%"&request("Sc")& _
           "%' Or PartsType like '%"&request("Sc")& _
           "%' Or PartsName like '%"&request("Sc")& _
           "%' Or WorkCata like '%"&request("Sc")& _
           "%' Or WorkContent like '%"&request("Sc")& _
           "%' Or NGMode like '%"&request("Sc")& _
           "%' Or NGEffect like '%"&request("Sc")& _
           "%' Or ProDesign like '%"&request("Sc")& _
           "%')  and PjID = '"&Request.cookies("fmea.PjID")&"' ORDER BY SubPjName,PageNo" 
           
        case 2
         ttl = request.cookies("fmea.PjKey")
         sql="SELECT * FROM t_Fmea_Item where  RPN > 2  and PjID = '"&Request.cookies("fmea.PjID")&"' ORDER BY SubPjName,PageNo" 

        case 3
         ttl = request.cookies("fmea.Serial")
         sql="SELECT * FROM t_Fmea_Item where (SubPjName like '%"&request("Sc")& _
           "%' Or PageNo like '%"&request("Sc")& _
           "%' Or PartsNo like '%"&request("Sc")& _
           "%' Or PartsType like '%"&request("Sc")& _
           "%' Or PartsName like '%"&request("Sc")& _
           "%' Or WorkCata like '%"&request("Sc")& _
           "%' Or WorkContent like '%"&request("Sc")& _
           "%' Or NGMode like '%"&request("Sc")& _
           "%' Or NGEffect like '%"&request("Sc")& _
           "%' Or ProDesign like '%"&request("Sc")& _
           "%') and  Pjkey like '"&Request.cookies("fmea.Serial")&"%' ORDER BY SubPjName,PageNo"         
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

<Script LANGUAGE = "JAVASCRIPT" type="text/javascript">
<!--Java Script Start
var winconf = "menubar=no,toolbar=yes,location=no,directories=no,status=no,scrollbar=no,resizeable=no,";
function about(filename,w,h) 
{
about = window.open(filename,"Copyright",winconf+"width="+w+",height="+h);
about.focus();
}
//Java Script End-->
</SCRIPT> 


<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
.STYLE2 {color: #0000FF}
-->
</style>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>机种管理项目一览</title>
</head>

<body bgcolor ="#FCFCFC">
<form action="Process_Browse.asp">
<input type="hidden" value="<% = request("page") %>" name="page">
<div style ="float:right ;width：200px;hight:200px;">
<%
if request.cookies("fmea.power") > "1" then
  response.write "<input type='submit' value='评价录入' name='OK' class='smallInput'>&nbsp;"
  response.write "<input type='submit' value='措施录入' name='OK' class='smallInput'>&nbsp;"
end if
%>
<input name="Search" type="text" class="smallInput" size="40">
<input type="submit" value="搜索" name="OK" class="smallInput">
<input type="submit" value="全系列搜索" name="OK" class="smallInput">
<input type="submit" value="全机种搜索" name="OK" class="smallInput">
<input type="submit" value="需对应项目" name="OK" class="smallInput">
<input type="submit" value="对应措施列表" name="OK" class="smallInput">
&nbsp;&nbsp;
</div>

<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align="center" colspan=18><b>
<% = ttl %>
作业内容评价列表</b>
</td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="2%">顺序</td>
<td width="5%" height="20">工程名称</td>
<td width="3%">标准书页码</td>
<td width="7%">部品编号</td>
<td width="10%">部品名称及种类</td>
<td width="8%">作业分类与内容</td>
<td width="10%">潜在缺陷模式</td>
<td width="10%">潜在缺陷后果</td>
<td width="7%">现行工程设计及过程控制</td>
<td width="2%">S</td>
<td width="2%">O</td>
<td width="2%">D</td>
<td width="2%">RI</td>
<td width="2%">措施</td>
<%
if request.cookies("fmea.power") > "1" then
   response.write "<td width='10%'>操作</td>"
end if
%>
</tr>

<% 
if request("act") = "del" then
        SqlStr_d = "delete t_Fmea_Item where ItmID ="&request("ItmID")     
        conn.execute(SqlStr_d)
        call closeDB()
        Response.Redirect "Process_Browse.asp"
elseif request("Act") = "Edit" then
        SqlStr_e = "SELECT * FROM t_Fmea_Item where ItmID ="&request("ItmID")
        Set rs_ed = Server.CreateObject("ADODB.Recordset")        
        rs_ed.open SqlStr_e,conn,1,1
        response.write"<tr bgcolor='#FCFCFC'>"
        response.write"<input name='ItmID' type='hidden' value ='"&rs_ed("ItmID")&"'>"
        response.write"<td><input name='WorkQueue' class='notline' size='3' value ='"&rs_ed("WorkQueue")&"'></td>"
        response.write"<td><input name='SubPjName' class='notline' size='9' value ='"&rs_ed("SubPjName")&"'></td>"
	response.write"<td><input name='PageNo' class='notline' size='9' value ='"&rs_ed("PageNo")&"'></td>"      
	response.write"<td><input name='PartsNo' class='notline' size='15' value ='"&rs_ed("PartsNo")&"'></td>"    
        response.write"<td><div><select name='PartsType'>"
        response.write"<option value ='"&rs_ed("PartsType")&"'>"&rs_ed("PartsType")&"</option>"  
        response.write"<option value ='板金'>板金</option>"
        response.write"<option value ='塑胶'>塑胶</option>"
        response.write"<option value ='轮轴'>轮轴</option>"
        response.write"<option value ='玻璃'>玻璃</option>"
        response.write"<option value ='弹簧'>弹簧</option>"
        response.write"<option value ='齿轮'>齿轮</option>"
        response.write"<option value ='皮带'>皮带</option>"
        response.write"<option value ='感光鼓'>感光鼓</option>"
        response.write"<option value ='组立部件'>组立部件</option>"
        response.write"<option value ='电线'> 电线</option>"
        response.write"<option value ='基板'> 基板</option>"
        response.write"<option value ='马达'> 马达</option>"
        response.write"<option value ='感应器'>感应器</option>"
        response.write"<option value ='离合器'>离合器</option>"
        response.write"<option value ='轴承'>轴承</option>"
        response.write"<option value ='海棉'>海棉</option>"
        response.write"<option value ='胶片'>胶片</option>"
        response.write"<option value ='卡环'>卡环</option>"
        response.write"<option value ='标签'>标签</option>"
        response.write"<option value ='螺丝'>螺丝</option>"
        response.write"<option value ='油'>油</option>"
        response.write"<option value ='耗材'>耗材</option>"
        response.write"</select></div><div>"
        response.write"<input name='PartsName' class='notline' size='20' value ='"&rs_ed("PartsName")&"'></div></td>"
        response.write"<td><div><select name='WorkCata'>"
        response.write"<option value ='"&rs_ed("WorkCata")&"'>"&rs_ed("WorkCata")&"</option>"  
        response.write"<option value ='打螺丝'>打螺丝</option>"
        response.write"<option value ='装入'>装入</option>"
        response.write"<option value ='张贴'>张贴</option>"
        response.write"<option value ='结线'>结线</option>"
        response.write"<option value ='线处理'>线处理</option>"
        response.write"<option value ='涂布'>涂布</option>"
        response.write"<option value ='清扫'>清扫</option>"
        response.write"<option value ='设定'>设定</option>"
        response.write"<option value ='检查'>检查</option>"
        response.write"</select></div>"
        response.write"<input name='WorkContent' class='notline' size='17' value ='"&rs_ed("WorkContent")&"'></td>"
        response.write"<td><input name='NGMode' class='notline' size='22' value ='"&rs_ed("NGMode")&"'></td>"
	response.write"<td><input name='NGEffect' class='notline' size='20' value ='"&rs_ed("NGEffect")&"'></td>"
        response.write"<td><input name='ProDesign' class='notline' size='15' value ='"&rs_ed("ProDesign")&"'></td>"
        response.write"<td><select name ='Sev' class= 'notline'>"
        response.write"<option value ='"&rs_ed("Sev")&"'>"&rs_ed("Sev")&"</option>"
        response.write"<option value ='1'>1</option>"
        response.write"<option value ='2'>2</option>"
        response.write"<option value ='3'>3</option>"
        response.write"<option value ='4'>4</option>"
        response.write"</select></td>"
        response.write"<td><select name ='Occ' class= 'notline'>"
        response.write"<option value ='"&rs_ed("Occ")&"'>"&rs_ed("Occ")&"</option>"
        response.write"<option value ='1'>1</option>"
        response.write"<option value ='2'>2</option>"
        response.write"<option value ='3'>3</option>"
        response.write"<option value ='4'>4</option>"

        response.write"</select></td>"
        response.write"<td><select name ='Det' class= 'notline'>"
        response.write"<option value ='"&rs_ed("Det")&"'>"&rs_ed("Det")&"</option>"
        response.write"<option value ='1'>1</option>"
        response.write"<option value ='2'>2</option>"
        response.write"<option value ='3'>3</option>"
        response.write"<option value ='4'>4</option>"

        response.write"</select></td>"
        response.write"<td></td><td></td>"
        response.write"<td><input name='OK' value='修正' type='submit' class='smallInput'>&nbsp;" 
        response.write"<input name='OK' value='升版' type='submit' class='smallInput'>&nbsp;" 
        response.write"<input name='OK' value='废止' type='submit' class='smallInput'>&nbsp;"
        response.write"<input name='OK' value='削除' type='submit' class='smallInput'>&nbsp;"
        response.write"</td></tr>"   
        call closeDB()
else

call MainBody() 

%>

<tr>
<td width="75%" colspan=18 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
[<a href="../desktop.asp">回到首页</a>][<a href="../PjManage/Pj_Browse.asp">回上一级</a>] [<a href="javascript:history.back()">返回</a>] 
</td><td  width="50%" align="right">
 <% call clspage() %> 
</table>
</td>
</tr>
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
	ItmID = rs("ItmID")
	response.write"<tr bgcolor='#FCFCFC'>"
        
        if rs("ItmStatus") = "N" then
        response.write"<td bgcolor ='FFFF00'>"&rs("WorkQueue")&"</td>"
        response.write"<td bgcolor ='FFFF00'>"&rs("SubPjName")&"</td>"
        response.write"<td bgcolor ='FFFF00'><div>(作业废止)</div>"&rs("PageNo")&"</td>"
        else
        response.write"<td>"&rs("WorkQueue")&"</td>"
        response.write"<td>"&rs("SubPjName")&"</td>"
        response.write"<td>"&rs("PageNo")&"</td>"
        end if
	response.write"<td><a href = http://km-square.km.local/kmc-portal/mp/pgm/ItemSearch_c/Pages/ITEM_Search_P_c.aspx?ITEMNO="&rs("PartsNo")&"&p_QASTATUS=%20 target ='_blank'>"&rs("PartsNo")&"</a></td>"
        response.write"<td><div>("&rs("PartsType")&")</div>"&rs("PartsName")&"</td>"
        response.write"<td><div>("&rs("WorkCata")&")</div><div>"&rs("WorkContent")&"</div></td>"
        response.write"<td>"&rs("NGMode")&"</td>"
	response.write"<td>"&rs("NGEffect")&"</td>"
        response.write"<td>"&rs("ProDesign")&"</td>"
        response.write"<td align ='center'>"&rs("Sev")&"</td>"
        response.write"<td align ='center'>"&rs("Occ")&"</td>"
        response.write"<td align ='center'>"&rs("Det")&"</td>"
        vRPN = trim(rs("RPN"))
        if vRPN > 2 Then
        response.write"<td bgcolor = '#DDFFAA'><span class='STYLE1'>&nbsp;"& vRPN &"</span></td>"    
        response.write"<td bgcolor = '#DDFFAA'><a href=Advice_Browse.asp?cat=2&ItmID="&ItmID&">对应</a></td>" 
           if request.cookies("fmea.power") > "1" then       
             response.write"<td>&nbsp;<a href=Process_Browse.asp?Act=Edit&ItmID="&ItmID&">修订</a>&nbsp;"
             response.write"<a href=../history/His_Process_Browse.asp?ItmID="&ItmID&">"&"修订履历</a>&nbsp;"
             response.write"<a href=Advice_add.asp?ItmID="&ItmID&">"&"措施填入</a>&nbsp;</td>"
           end if
        response.write"</tr>"
        else
        response.write"<td>"& vRPN &"</td>" 
        response.write"<td></td>"
           if request.cookies("fmea.power") > "1" then      
             response.write"<td>&nbsp;<a href=Process_Browse.asp?Act=Edit&ItmID="&ItmID&">修订</a>&nbsp;"
             response.write"<a href=../history/His_Process_Browse.asp?ItmID="&ItmID&">"&"修订履历</a>&nbsp;</td>"
           end if
        response.write"</tr>"
        end if
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            '没有上一页，但是有下一页'
	Response.Write "首页&nbsp;上页&nbsp;"
	Response.Write "<a href=Process_Browse.asp?page="&page+1&">下页</a>&nbsp;"
	Response.Write "<a href=Process_Browse.asp?page="&pagecount&">尾页</a>"

elseif page=pagecount and not page=1  then         '没有下一页，但是有上一页' 
	Response.Write "<a href=Process_Browse.asp?page=1>首页</a>&nbsp;&nbsp;"
	Response.Write "<a href=Process_Browse.asp?page="&page-1&">上页</a>&nbsp;&nbsp;下页&nbsp;尾页"

elseif page<1 or page>pagecount then '没有任何记录' 
	Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

elseif page=1 and page=pagecount then '没有上一页，没有下一页'
	Response.Write "首页&nbsp;上页&nbsp;下页&nbsp;尾页"

else
	Response.Write "<a href=Process_Browse.asp?page=1>首页</a>&nbsp;"
	Response.Write "<a href=Process_Browse.asp?page="&page-1&">上页</a>&nbsp;"
	Response.Write "<a href=Process_Browse.asp?page="&page+1&">下页</a>&nbsp;"
	Response.Write "<a href=Process_Browse.asp?page="&pagecount&">尾页</a>"
end if 

Response.Write "&nbsp;第<input type='Text' name='pagename' size='2' class='smallInput'>页"
Response.Write "<input type='submit' value='GO' name='ok' class='smallInput'>"
Response.Write  "&nbsp; 第"&page&"页&nbsp;共"&pagecount&"页 计"&rs.recordcount&"条记录"
end sub
%>

