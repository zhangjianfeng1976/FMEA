<!--#include file="../includes/db.asp"-->
<% 
Select case request("OK")
    case "��ѯ" 	                
         Select case Request("cata")
            case "ģ����ѯ"
                Response.Redirect "SafeParts.asp?Cat=0&Sc="&Request("Search")								
          End select         
    case "����" 
        Response.Redirect "SafeParts.asp?Action=ADD"
    case "GO"
        Response.Redirect "SafeParts.asp?page="&Request("pagename")	
    case "ȷ��"
        sql="INSERT INTO t_SafeParts(PartsNo,PartsName,InUser,InTime) " _
             &"VALUES ('" _
             &trim(request("PartsNo"))&"', '"_
			 &trim(request("PartsName"))&"', '"_
			 &Request.cookies("fmea.truename")&"', '"_
			 &Now()&"')"
		call openDB()
		conn.execute(sql)
	    call closeDB()
        response.Redirect "SafeParts.asp"
     case "�޶�"		 
        sql="Update t_SafeParts set PartsNo = '"&request("PartsNo") _
              &"',PartsName = '"&request("PartsName") _
              &"',InUser = '"&Request.cookies("fmea.truename") _
              &"',InTime = '"&Now()&"' where Id = "&request("ID")
        call openDB()
        conn.execute(sql)
        call closeDB()
        response.Redirect "SafeParts.asp"		 		  	  
 End Select

   
'ɾ��______________________________________________________________________'
ID=request("ID")
Select Case Request("Action")
    Case "Delete" 
        Call openDB()
        sql="SELECT InUser FROM t_SafeParts where ID ="&ID
        rs.open sql,conn,1,1
		if trim(rs("InUser"))<>trim(Request.cookies("fmea.truename")) then
			response.Write "�㲻��ɾ�����˵ļ�¼!![<a href='javascript:history.back()'>����</a>]"
			response.End()
			Call closeDB()
		else
			Call closeDB()
			sql="Delete t_SafeParts Where ID ="&ID       
			Call OpenDB()
			conn.execute(sql)
			Call closeDB()
		end if
	   
    Case "Edit" 
        Call openDB()
		sql="SELECT  * FROM  t_SafeParts where ID ="&ID
        rs.open sql,conn,1,1
		if trim(rs("InUser"))<>trim(Request.cookies("fmea.truename")) then
			response.Write "�㲻���޶����˵ļ�¼!![<a href='javascript:history.back()'>����</a>]"
			response.End()
			Call closeDB()
		else
			ID             = rs("ID")
            PartsNo        = rs("PartsNo")
            PartsName      = rs("PartsName")
            InUser         = rs("InUser")
            InTime         = rs("InTime")
            call closeDB()    
		end if	   
End Select

pgx=500
select case request("cat")
    case ""
        pgx=30
        sql="SELECT * FROM t_SafeParts order by ID desc"
    case 0
        sql="SELECT * FROM t_SafeParts where PartsNo like '%"&request("Sc")& _
		    "%' Or PartsName like '%"&request("Sc")& _
			"%' Or InUser like '%"&request("Sc")& _
            "%' ORDER BY Id DESC" 	   				      
end select

Call openDB()
rs.open sql,conn,1,1
If not rs.eof then
    rs.pagesize = pgx
    pagecount=rs.PageCount 
    page=int(request.QueryString ("page"))
    if page<=0 then page=1
    if request.QueryString("page")="" then page=1
    rs.AbsolutePage=page 
End if
 %>  
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Includes/Style.css" rel="stylesheet" type="text/css">
<SCRIPT language=JavaScript type=text/JavaScript>
//˫����������Ļ�Ĵ���
var currentpos,timer;
function initialize()
{
timer=setInterval ("scrollwindow ()",1);
}
function sc()
{
clearInterval(timer);
}
function scrollwindow()
{
currentpos=document.body.scrollTop;
window.scroll(0,++currentpos);
if (currentpos !=document.body.scrollTop)
sc();
}
document.onmousedown=sc
document.ondblclick=initialize
</SCRIPT>
<style type="text/css">
<!--
.STYLE2 {
	font-size: 20;
	color: #FFFFFF;
	font-family: "��Բ";
}
-->
</style>
</head>

<body>
<form action="SafeParts.asp">
<table width="50%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE" align="center" >
<tr><td align="right" colspan="7" bgcolor="#E6F7F0">
<div style ="float:left;">
<div style ="float:right ;">
<select name='cata' class="smallinput">
<option>ģ����ѯ</option>
</select>
<input name="Search" type="text" class="smallInput" size="18" value ="<% = request("sc") %>">
<input type="submit" value="��ѯ" name="OK" class="smallInput">
<% if request.cookies("fmea.power") >  "1" then 
response.write"<input type='submit' value='����' name='OK' class='smallInput'>"
end if
%>
[<a href="javascript:history.back()">����</a>][<a href="../ProcessMan/Process_Browse_total.asp?cat=7">��ȫ��Ҫ��Ʒ��������</a>]
</div>
<div style ="float:left;"><font face="��������" size="4" color="#000000">��ȫ��Ҫ��ƷLIST</font></div>
</div>  
</td>
</tr>
<tr align="center" bgcolor="#E6F7FF">
<td width="10%" height="20">��Ʒ���</td>
<td width="20%">��Ʒ����</td>
<td width="8%">������</td>
<td width="10%">����ʱ��</td>
<%
if request.cookies("fmea.power") > "1" then
	response.write "<td width='10%'>����</td>"
end if
%>
</tr>
<% 
select case request("Action")
	case "ADD"
		response.write"<tr bgcolor='#fcfcfc'>"
		response.write"<td><input name='PartsNo' class='notline' size='30' tabindex ='1'></td>"
		response.write"<td><input name='PartsName' class='notline' size='75' tabindex ='2'></td>"
		response.write"<td colspan = '3'><input type='submit' value='ȷ��' name='OK' class='smallInput' tabindex ='3'></td>"
		response.write"</tr>"  
	case "Edit" 
		response.write"<input name='ID' type='hidden'  value = '"&ID&"'>"
		response.write"<tr bgcolor='#fcfcfc'>"
		response.write"<td><input name='PartsNo' class='notline' size='30' value = '"&PartsNo&"'></td>"
		response.write"<td><input name='PartsName' class='notline' size='75' value = '"&PartsName&"'></td>"
		response.write"<td colspan = '3'><input type='submit' value='�޶�' name='OK' class='smallInput'></td>"
		response.write"</tr>"
case else  
	call mainbody
end select

 %>
 
<tr><td align="left" colspan="5" bgcolor="#E6F7FF">
<div>
<div style = "float:left";>
[<a href="../desktop.asp">�ص���ҳ</a>] [<a href="javascript:history.back()">����</a>] 
</div><div style = "float:right";>
<% call clspage() %>
</div>
</td>

</tr>
</table>
<% call closeDB() %>
</form>
</body>
</html>

<%
Sub Mainbody()

if rs.eof then
  response.write "<tr><td width=""75%"" colspan=12>"
  response.write "û�����ݡ�"
  response.write "</td></tr>"
end if

n=1 
While Not rs.eof and n<=rs.pagesize 
  response.write"<tr bgcolor='#FCFCFC'>"
  response.write"<td>&nbsp&nbsp<a href ='../ProcessMan/Process_Browse_total.asp?Cat=0&Sc="&trim(rs("PartsNo"))&"'>"&rs("PartsNo")&"</a></td>"
  response.write"<td>&nbsp&nbsp"&rs("PartsName")&"</td>"
  response.write"<td>&nbsp&nbsp"&rs("InUser")&"</td>"
  response.write"<td>&nbsp&nbsp"&rs("InTime")&"</td>"
  if request.cookies("fmea.power") > "1" then  
	response.write"<td>&nbsp&nbsp<a href=SafeParts.asp?Action=Delete&id="&rs("id")&">ɾ��</a>&nbsp;&nbsp;"
	response.write"<a href=SafeParts.asp?Action=Edit&id="&rs("id")&">�޶�</a></td>"
  end if
  response.write"</tr>"
  n=n+1 
  rs.movenext '��ʾҳ�������' 
Wend 

End Sub

Sub clspage()

if page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
	Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
	Response.Write "<a href=SafeParts.asp?page="&page+1&">��ҳ<a>&nbsp;"
	Response.Write "<a href=SafeParts.asp?page="&pagecount&">βҳ<a>"
elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
	Response.Write "<a href=SafeParts.asp?page=1>��ҳ<a>&nbsp;&nbsp;"
	Response.Write "<a href=SafeParts.asp?page="&page-1&">��ҳ<a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"
elseif page<1 or page>pagecount then 'û���κμ�¼' 
	Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"
elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
	Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"
else
	Response.Write "<a href=SafeParts.asp?page=1>��ҳ<a>&nbsp;"
	Response.Write "<a href=SafeParts.asp?page="&page-1&">��ҳ<a>&nbsp;"
	Response.Write "<a href=SafeParts.asp?page="&page+1&">��ҳ<a>&nbsp;"
	Response.Write "<a href=SafeParts.asp?page="&pagecount&">βҳ<a>"
end if

Response.Write "&nbsp;��<input type='Text' name='pagename' size='2' class='smallInput'>ҳ"
Response.Write "<input type='submit' value='GO' name='ok' class='smallInput'>"
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����¼"

end sub
%>

 
