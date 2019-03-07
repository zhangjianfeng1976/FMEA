<!--#include file="../includes/db.asp"-->
<%
if (request.cookies("fmea.Pjkey") ="" or  request.cookies("fmea.Pjkey") <> request("Pjkey")) and  request("PjKey") <>"" then  
response.cookies("fmea.Pjkey") = request("Pjkey")
end if 

if (request.cookies("fmea.PjID") ="" or  request.cookies("fmea.PjID") <> request("PjID")) and  request("PjID") <>"" then  
response.cookies("fmea.PjID") = request("PjID")
end if 

Select Case Request("OK") 
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

            Response.Redirect "Process_Browse_search.asp?cat=0&sc="&request("sc")

         Case "修订(升版)" 
            
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

            Response.Redirect "Process_Browse_search.asp?cat=0&sc="&request("sc")
          
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
            
End Select

pgs =2500
select case request("cat")
       case "" 
         ttl = request.cookies("fmea.PjKey")
         sql="SELECT * FROM v_Fmea_Item where PjID = "&request.cookies("fmea.PjID")&" order by SubPjName,PageNo,WorkQueue"
       case 0
         ttl ="全机种"
         sql="SELECT * FROM v_Fmea_Item where (SubPjName like '%"&request("Sc")& _
           "%' Or PageNo like '%"&request("Sc")& _
           "%' Or PartsType like '%"&request("Sc")& _
           "%' Or PartsNo like '%"&request("Sc")& _
           "%' Or PartsName like '%"&request("Sc")& _
           "%' Or WorkCata like '%"&request("Sc")& _
           "%' Or WorkContent like '%"&request("Sc")& _
           "%' Or NGMode like '%"&request("Sc")& _
           "%' Or NGEffect like '%"&request("Sc")& _
           "%' Or ProDesign like '%"&request("Sc")& _
           "%') ORDER BY SubPjName,PageNo,WorkQueue" 
        
        case 1
         ttl = request.cookies("fmea.PjKey")
         sql="SELECT * FROM v_Fmea_Item where (SubPjName like '%"&request("Sc")& _
           "%' Or PageNo like '%"&request("Sc")& _
           "%' Or PartsNo like '%"&request("Sc")& _
           "%' Or PartsType like '%"&request("Sc")& _
           "%' Or PartsName like '%"&request("Sc")& _
           "%' Or WorkCata like '%"&request("Sc")& _
           "%' Or WorkContent like '%"&request("Sc")& _
           "%' Or NGMode like '%"&request("Sc")& _
           "%' Or NGEffect like '%"&request("Sc")& _
           "%' Or ProDesign like '%"&request("Sc")& _
           "%')  and PjID = '"&Request.cookies("fmea.PjID")&"' ORDER BY SubPjName,PageNo,WorkQueue"         
        
        case 3
         ttl = request.cookies("fmea.Serial")
         sql="SELECT * FROM v_Fmea_Item where (SubPjName like '%"&request("Sc")& _
           "%' Or PageNo like '%"&request("Sc")& _
           "%' Or PartsNo like '%"&request("Sc")& _
           "%' Or PartsType like '%"&request("Sc")& _
           "%' Or PartsName like '%"&request("Sc")& _
           "%' Or WorkCata like '%"&request("Sc")& _
           "%' Or WorkContent like '%"&request("Sc")& _
           "%' Or NGMode like '%"&request("Sc")& _
           "%' Or NGEffect like '%"&request("Sc")& _
           "%' Or ProDesign like '%"&request("Sc")& _
           "%') and  Pjkey like '"&Request.cookies("fmea.Serial")&"%' ORDER BY SubPjName,PageNo,WorkQueue"   
       
     
        case 5
         ttl = request.cookies("fmea.PjKey")
         sql="SELECT * FROM v_Fmea_Item where ItmID ="&request("ItmID")&" ORDER BY SubPjName,PageNo,WorkQueue" 
      
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

var winconf = "menubar=no,toolbar=no,location=no,directories=no,status=no,scrollbar=no,resizeable=no,";
function about(filename,w,h) 
{
about = window.open(filename,"Copyright",winconf+"left=0,width="+w+",height="+h);
about.focus();
}
-->
</script> 
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
<form action="Process_Browse_search.asp">
<% 
pg = request("page")
if pg  ="" then  pg = 1
response.write "<input type='hidden' value='"& pg  &"' name='page'>"
%>
<input type='hidden' value="<%=request("Sc") %>" name='sc'>
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">

<tr><td width="75%" colspan=18 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
[<a href="../desktop.asp">回到首页</a>]
[<a href="javascript:history.back()">返回</a>] 
</td><td  width="50%" align="right">

</table>
</td></tr>

<tr bgcolor="#E6F7FF"><td height="30" align="center" colspan=18><b><% = ttl %>作业内容评价列表</b></td></tr> 

<tr bgcolor="#FCFCFC" align="center">
<td width="2%" height="20">顺序</td>
<td width="5%">工程名称</td>
<td width="3%">标准书页码</td>
<td width="7%">部品编号</td>
<td width="8%">部品名称及种类</td>
<td width="8%">作业分类与内容</td>
<td width="10%">潜在缺陷模式</td>
<td width="10%">潜在缺陷后果</td>
<td width="7%">现行工程设计及过程控制</td>
<td width="2%">S</td>
<td width="2%">O</td>
<td width="2%">D</td>
<td width="2%">RI</td>
<td width="4%">措施</td>
<%
if request.cookies("fmea.power") > "1" then
   response.write "<td width='4%'>操作</td>"
end if
%>
</tr>

<% 
if request("act") = "del" then
        SqlStr_d = "delete t_Fmea_Item where ItmID ="&request("ItmID")     
        conn.execute(SqlStr_d)
        call closeDB()
        Response.Redirect "Process_Browse_search.asp"
elseif request("Act") = "Edit" then
        SqlStr_e = "SELECT * FROM t_Fmea_Item where ItmID ="&request("ItmID")
        Set rs_ed = Server.CreateObject("ADODB.Recordset")        
        rs_ed.open SqlStr_e,conn,1,1
        response.write"<tr bgcolor='#FCFCFC'>"
        response.write"<input name='ItmID' type='hidden' value ='"&rs_ed("ItmID")&"'>"
        response.write"<td><input name='WorkQueue' class='notline' size='3' value ='"&rs_ed("WorkQueue")&"'></td>"
        'response.write"<td><input name='SubPjName' class='notline' size='9' value ='"&rs_ed("SubPjName")&"'></td>"
        response.write"<td><textarea name='SubPjName' cols='10' rows='4' wrap='virtual' class='notline' >"&rs_ed("SubPjName")&"</textarea></td>"
	response.write"<td><input name='PageNo' class='notline' size='9' value ='"&rs_ed("PageNo")&"'></td>"      
	response.write"<td><input name='PartsNo' class='notline' size='15' value ='"&rs_ed("PartsNo")&"'></td>"    
        response.write"<td><div><select name='PartsType'>"
        response.write"<option value ='"&rs_ed("PartsType")&"'>"&rs_ed("PartsType")&"</option>"  
        response.write"<option value ='框架件'>框架件</option>"
        response.write"<option value ='驱动件'>驱动件</option>"
        response.write"<option value ='光学件'>光学件</option>"
        response.write"<option value ='电气件'>电气件</option>"
        response.write"<option value ='引导件'>引导件</option>"
        response.write"<option value ='导通件'>导通件</option>"
        response.write"<option value ='组立部件'>组立部件</option>"
        response.write"<option value ='弹簧'>弹簧</option>"
        response.write"<option value ='轴承'>轴承</option>"
        response.write"<option value ='海棉'>海棉</option>"
        response.write"<option value ='胶片'>胶片</option>"
        response.write"<option value ='卡环'>卡环</option>"
        response.write"<option value ='标签'>标签</option>"
        response.write"<option value ='螺丝'>螺丝</option>"
        response.write"<option value ='润滑材'>润滑材</option>"
        response.write"<option value ='耗材'>耗材</option>"
        response.write"<option value ='其他'>其他</option>"        
        response.write"</select></div><div>"
        response.write"<textarea name='PartsName' cols='15' rows='2' wrap='virtual' class='notline' >"&rs_ed("PartsName")&"</textarea></div></td>"
        'response.write"<input name='PartsName' class='notline' size='20' value ='"&rs_ed("PartsName")&"'></div></td>"
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
        'response.write"<input name='WorkContent' class='notline' size='17' value ='"&rs_ed("WorkContent")&"'></td>"
        response.write"<textarea name='WorkContent' cols='15' rows='2' wrap='virtual' class='notline' >"&rs_ed("WorkContent")&"</textarea></td>"
        response.write"<td><textarea name='NGMode' cols='22' rows='4' wrap='virtual' class='notline' >"&rs_ed("NGMode")&"</textarea></td>"
        response.write"<td><textarea name='NGEffect' cols='22' rows='4' wrap='virtual' class='notline' >"&rs_ed("NGEffect")&"</textarea></td>"
        response.write"<td><textarea name='ProDesign' cols='18' rows='4' wrap='virtual' class='notline' >"&rs_ed("ProDesign")&"</textarea></td>"
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

        response.write"<td bgcolor='#F0F0F0'></td><td bgcolor='#F0F0F0'></td><td bgcolor='#F0F0F0'></td></tr>"
        response.write"<tr bgcolor='#FCFCFC'><td colspan =15  align='center' height =35>"
        response.write"<input name='OK' value='修正' type='submit' class='smallInput'>&nbsp;" 
        response.write"<input name='OK' value='修订(升版)' type='submit' class='smallInput'>&nbsp;" 
        response.write"<input name='OK' value='废止' type='submit' class='smallInput'>&nbsp;"
        response.write"<input name='OK' value='削除' type='submit' class='smallInput'>&nbsp;"
        response.write"[<a href='javascript:history.back()'>返回</a>]" 
        response.write"</td></tr>" 
        response.write"<tr bgcolor='#FCFCFC'><td colspan =15>"
        response.write""
        response.write"【修正】是不会留下原来的记录的，所以，只在录入错误时使用，要慎用。<br>"
        response.write"【修订(升版)】是会留下原来的记录的，可以在修订记录上查看，所以，是在文件修订时使用。修订完成后请在【机种列表】上进行升版。<br>"
        response.write"【废止】与升版一样，会留下原来的记录的，而又想在原文件上保留标记，这时，修订完成后也需要在【机种列表】上进行升版。<br>"
        response.write"【削除】是将记录删除，这种情况有需要的时候才使用，必须确认再确认的。要慎用。最好用【修正】或【废止】。"
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

    if isnull(rs("PartsOK")) then
       response.write"<td><a href = http://km-square.km.local/kmc-portal/mp/pgm/ItemSearch_c/Pages/ITEM_Search_P_c.aspx?ITEMNO="
       response.write rs("PartsNo")&"&p_QASTATUS=%20 target ='_blank'>"&rs("PartsNo")&"</a></td>"
    else
       response.write"<td bgcolor ='FF9955'>"
       response.write"<a href = '../His_Question/PartsFamily.asp?cat=0&sc="&trim(rs("PartsNo"))&"'>★</a>"
       response.write"<a href = http://km-square.km.local/kmc-portal/mp/pgm/ItemSearch_c/Pages/ITEM_Search_P_c.aspx?ITEMNO="
       response.write rs("PartsNo")&"&p_QASTATUS=%20 target ='_blank'>"&rs("PartsNo")&"</a></td>"
    end if

    response.write"<td><div>("&rs("PartsType")&")</div><div>"&rs("PartsName")&"</div></td>"
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
       response.write"<td bgcolor = '#DDFFAA'>"
       response.write"<div><a href=Advice_Browse.asp?cat=2&ItmID="&ItmID&">对应</a>&nbsp;"
       response.write"<a href=Process_Browse_search.asp?cat=0&sc="&request("sc")&" onClick=about('Advice_Read.asp?cat=2&ItmID="&ItmID&"',850,250)>查看</a></div>"
       if isnull(rs("ItmID_his")) then
       else
       response.write"<div><a href=../history/His_Process_Browse_search.asp?ItmID="&ItmID&">修订履历</a></div>"
       end if
       if isnull(rs("ItmID_fb")) then
       else
         response.write"<div><a href=Process_Browse_search.asp?cat=0&sc="&request("sc")&" onClick=about('../His_Question/FBRead.asp?cat=3&ItmID="&ItmID&"',600,600)>过去问题</a></div>"
       end if
       response.write"</td>"        
    else
       response.write"<td>"& vRPN &"</td>" 
       response.write"<td>"
       if isnull(rs("ItmID_his")) then
       else
       response.write"<div><a href=../history/His_Process_Browse_search.asp?ItmID="&ItmID&">修订履历</a></div>"
       end if
       if isnull(rs("ItmID_fb")) then
       else
         response.write"<div><a href=Process_Browse_search.asp?cat=0&sc="&request("sc")&" onClick=about('../His_Question/FBRead.asp?cat=3&ItmID="&ItmID&"',600,600)>过去问题</a></div>"
       end if
       response.write"</td>"
    end if 

    if request.cookies("fmea.power") > "1" then    
       response.write"<td>"
       response.write"<div height =20><a href=Process_Browse_search.asp?Act=Edit&ItmID="&ItmID&">修订</a>&nbsp;"
       WQ = trim(rs("WorkQueue"))
       WQ = WQ + 0.01
       response.write"<a href=Process_Add.asp?cat=1&WQ="& WQ &">插入</a></div>"
       response.write"<div><a href=Process_Browse_search.asp?cat=0&sc="&request("sc")&" onClick=about('Advice_Add.asp?cat=1&ItmID="&ItmID&"',800,250)>措施填入</a></div>"
       response.write"<div><a href=../His_Question/FailBase.asp?Action=ADD&ItmID="&ItmID&">历史填入</a></div>"
       response.write"</td>"
    end if

    response.write"</tr>"        
    rs.MoveNext
n=n+1
wend
End Sub

%>

