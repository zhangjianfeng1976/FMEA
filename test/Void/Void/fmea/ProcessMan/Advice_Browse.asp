<!--#include file="../includes/db.asp"-->
<%
Select Case Request("OK") 
       Case "����" 	                
	     Response.Redirect "Advice_Browse.asp?Cat=0&Sc="&Request("Search")
       Case "ȫ��������" 	                
	     Response.Redirect "Advice_Browse.asp?Cat=1&Sc="&Request("Search")
        Case "�޶�"
             Sev = int(request("Sev"))
             Occ = int(request("Occ"))
             Det = int(request("Det"))
             RPN = Sev * Occ * Det
                 
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
            conn.execute(SqlStr_c)
            call closeDB()			             
End Select

pgs =2500
select case request("cat")
       case "" 
         ttl ="ȫ��"
         sql="SELECT * FROM v_Fmea_ttl where AdID is not null"
         pgs = 25
       case 0
          ttl =""
          sql = "SELECT * FROM v_Fmea_ttl where "
          sql = sql & "(AdvContent like '%" & request("sc") & "%' or "
          sql = sql & "Rspnser like '%" & request("sc") & "%' or "
          sql = sql & "ActionContent like '%" & request("sc") & "%'"
          sql = sql & ") and (AdID is not null)"
          pgs = 2500
       case 1
          ttl = "ȫ��"
          sql = "SELECT * FROM v_Fmea_ttl where "
          sql = sql & "(AdvContent like '%" & request("sc") & "%' or "
          sql = sql & "Rspnser like '%" & request("sc") & "%' or "
          sql = sql & "ActionContent like '%" & request("sc") & "%'"
          sql = sql & ") order by ItmID"
          pgs = 2500
       case 2
         sql="SELECT * FROM v_Fmea_ttl WHERE ItmID = "& request("ItmID")       
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
<title>���ָ��ƴ�ʩ��Ŀһ��</title>
</head>

<body bgcolor ="#FCFCFC">
<form action="Advice_Browse.asp">
<div style ="float:right ;width��200px;hight:200px;">
<!-- <input type='submit' value="��ʩ¼��" name="OK" class="smallInput"> -->
<input name="Search" type="text" class="smallInput" size="40">
<input type="submit" value="����" name="OK" class="smallInput">
<input type="submit" value="ȫ��������" name="OK" class="smallInput">
&nbsp;&nbsp;
</div>

<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align="center" colspan=13><b>
<% = ttl %>
���ƴ�ʩһ��</b>
</td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="8%" height="20" rowspan ="2">FMEA��ż��汾</td>
<td width="8%"  rowspan ="2">��ҵ��׼��ҳ�뼰����</td>
<td width="15%" rowspan ="2">�����ʩ</td>
<td width="5%"  rowspan ="2">������</td>
<td width="8%"  rowspan ="2">Ŀ�������</td>
<td width="18%" colspan ="5">��ʩ���</td>
<%
if request.cookies("fmea.power") = "0" then
   response.write "<td width='10%' rowspan ='2'>����</td>"
end if
%>
</tr><tr bgcolor="#FCFCFC" align="center">
<td width="10%">���ô�ʩ</td>
<td width="2%">S</td>
<td width="2%">O</td>
<td width="2%">D</td>
<td width="2%">RPN</td>
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
        response.write"<input name='AdID' type='hidden'class='notline' size='5' value ='"&rs_ed("AdID")&"'>"
        response.write"<input name='ItmID' type='hidden'class='notline' size='5' value ='"&rs_ed("ItmID")&"'>"
        response.write"<td></td><td></td>"
	response.write"<td><input name='AdvContent' class='notline' size='40' value ='"&rs_ed("AdvContent")&"'></td>"
        response.write"<td><input name='Rspnser' class='notline' size='10' value ='"&rs_ed("Rspnser")&"'></td>"
	response.write"<td><input name='FshDate' class='notline' size='15' value ='"&rs_ed("FshDate")&"'></td>"
        response.write"<td><input name='ActionContent' class='notline' size='25' value ='"&rs_ed("ActionContent")&"'></td>"
        response.write"<td><input name='Sev' class='notline' size='3' value ='"&rs_ed("ResultS")&"'></td>"
        response.write"<td><input name='Occ' class='notline' size='3' value ='"&rs_ed("ResultO")&"'></td>"
        response.write"<td><input name='Det' class='notline' size='3' value ='"&rs_ed("ResultD")&"'></td>"
        response.write"<td></td>"
        response.write"<td><input name='OK' value='�޶�' type='submit' class='smallInput'></td>" 
        response.write"</tr>"   
        call closeDB()
else

call MainBody()

 %>

<tr>
<td width="75%" colspan=13 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
<% call clspage() %>
</td><td  width="50%" align="right">
[<a href="javascript:history.back()">����</a>]   
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
	response.write "û�м�¼��"
	response.write "</td></tr>"
end if

n=1 
while not rs.eof and n<= rs.pagesize
        AdID = rs("AdID")
	ItmID = rs("ItmID")
	response.write"<tr bgcolor='#FCFCFC'>"
        response.write"<td>"&rs("PjKey")&"</td>"	
        response.write"<td>"&rs("PageNo")&"-"&rs("WorkQueue")&"</td>"	
        response.write"<td>"&rs("AdvContent")&"</td>"	
	response.write"<td>"&rs("Rspnser")&"</td>"
        response.write"<td>"&rs("FshDate")&"</td>"
        response.write"<td>"&rs("ActionContent")&"</td>"
        response.write"<td>"&rs("ResultS")&"</td>"
        response.write"<td>"&rs("ResultO")&"</td>"
        response.write"<td>"&rs("ResultD")&"</td>"
        response.write"<td>"&rs("ResultRPN")&"</td>" 
        if request.cookies("fmea.power") = "0" then
        response.write"<td>&nbsp;<a href=Advice_Browse.asp?Act=Edit&AdID="&AdID&">"&"����</a>&nbsp;&nbsp"
        response.write"<a href=Advice_Browse.asp?Act=VerUp&AdID="&AdID&">"&"�޶�</a>&nbsp;&nbsp"
        response.write"<a href=Advice_History.asp?ItmID="&ItmID&">"&"�޶�����</a>&nbsp;&nbsp</td>"
        end if
	response.write"</tr>"
	rs.MoveNext
n=n+1
wend
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
Response.Write "<a href=Advice_Browse.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=Advice_Browse.asp?page="&pagecount&">βҳ</a>"

elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
Response.Write "<a href=Advice_Browse.asp?page=1>��ҳ</a>&nbsp;&nbsp;"
Response.Write "<a href=Advice_Browse.asp?page="&page-1&">��ҳ</a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
Response.Write "<a href=Advice_Browse.asp?page=1>��ҳ</a>&nbsp;"
Response.Write "<a href=Advice_Browse.asp?page="&page-1&">��ҳ</a>&nbsp;"
Response.Write "<a href=Advice_Browse.asp?page="&page+1&">��ҳ</a>&nbsp;"
Response.Write "<a href=Advice_Browse.asp?page="&pagecount&">βҳ</a>"
end if 
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����¼"
end sub


%>

