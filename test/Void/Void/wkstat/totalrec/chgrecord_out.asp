<!--#include file="../includes/db.asp"-->
<%
Select Case Request("OK") 
       Case "��ѯ" 	                
            select case Request("cata")
              case "�߼���ѯ"
              case "ģ����ѯ"
	         Response.Redirect "chgrecord_out.asp?Cat=0&Sc="&Request("Search")			   		
          end select            
End Select


pgs =2500
select case request("cat")
       case "" 
       sql="SELECT top 300 * FROM v_wdairy_out order by ���� DESC"
       pgs =20
       case 0
       sql="SELECT  * FROM v_wdairy_out where ԭ��ҵ�߹��� like '%"&request("Sc")& _
           "%' Or ��λ�߹��� like '%"&request("Sc")& _
           "%' Or LINE like '%"&request("Sc")& _
           "%' Or ���� like '%"&request("Sc")& _
           "%' Or ���� like '%"&request("Sc")& _
           "%' Or ��λ like '%"&request("Sc")& _
           "%' ORDER BY ����"                      
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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Includes/Style.css" rel="stylesheet" type="text/css">
<title></title>
</head>

<body>
<form action="chgrecord_out.asp">
<table width="100%" border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
<tr>
<td bgcolor="#e6f7ff" colspan ='11'>
<table width="100%"><tr><td align='left'>
<b><font face="������κ" size="3" color="#CC6600">��������Ա��λ��¼</font></b>
</td><td align='right'>
<select name='cata' class="smallinput">
<option>ģ����ѯ</option>
<option>�����ڲ�ѯ</option>
</select>
<input name="Search" type="text" class="smallInput" size="19">
<input type="submit" value="��ѯ" name="OK" class="smallInput">
</td></tr>
</table>
</td>
</tr> 

<tr bgcolor="#FCFCFC" align="center">
<td width="8%" height="20">����</td>
<td width="8%">Line</td>
<td width="8%">����</td>
<td width="8%">����</td>
<td width="8%">��λ</td>
<td width="8%">ԭ��ҵ�߹���</td>
<td width="8%">����</td>
<td width="8%">ԭ��</td>
<td width="8%">��λ�߹���</td>
<td width="8%">��λ</td>
<td width="8%">ְ��</td>
</tr>

<% Call dairy %>

<tr>
<td width="75%" colspan=11 height="30" bgcolor="#EFEFEF" align='center'>
<table width="100%">
<tr><td width="50%" align="left">
<% call clspage() %>
</td><td  width="50%" align="right">
[<a href="javascript:history.back()">����</a>]   
<% Call CloseDB() %>
</table>
</td>
</tr>

</table>
</form>
</body>
</html>


<%
Sub dairy()
   if rs.eof then
      response.write "<tr><td width=75% colspan=12>"
      response.write "��û�м�¼��"
      response.write "</td></tr>"
   end if

      n=1 
      while not rs.eof and n<= rs.pagesize
      line = rs("line")
      response.write "<tr>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("����")&"</td>"
      response.write "<td width='9%' height='16' bgcolor='#FCFCFC'>" 
      response.write line
      response.write "</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("����")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("����")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("��λ")&"</td>" 
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("ԭ��ҵ�߹���")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("����")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("ԭ��")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("��λ�߹���")&"</td>"
      response.write "<td bgcolor='#e6F6F6' align ='center'>"&rs("��λ")&"</td>"
      response.write "<td bgcolor='#FCFCFC' align ='center'>"&rs("ְ��")&"</td>"
      response.write "</tr>"
      rs.MoveNext
      n=n+1
      wend
   
End Sub

Sub clspage()
If page=1 and not page=pagecount  then            'û����һҳ����������һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;"
Response.Write "<a href=chgrecord_out.asp?page="&page+1&">��ҳ<a>&nbsp;"
Response.Write "<a href=chgrecord_out.asp?page="&pagecount&">βҳ<a>"

elseif page=pagecount and not page=1  then         'û����һҳ����������һҳ' 
Response.Write "<a href=chgrecord_out.asp?page=1>��ҳ<a>&nbsp;&nbsp;"
Response.Write "<a href=chgrecord_out.asp?page="&page-1&">��ҳ<a>&nbsp;&nbsp;��ҳ&nbsp;βҳ"

elseif page<1 or page>pagecount then 'û���κμ�¼' 
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

elseif page=1 and page=pagecount then 'û����һҳ��û����һҳ'
Response.Write "��ҳ&nbsp;��ҳ&nbsp;��ҳ&nbsp;βҳ"

else
Response.Write "<a href=chgrecord_out.asp?page=1>��ҳ<a>&nbsp;"
Response.Write "<a href=chgrecord_out.asp?page="&page-1&">��ҳ<a>&nbsp;"
Response.Write "<a href=chgrecord_out.asp?page="&page+1&">��ҳ<a>&nbsp;"
Response.Write "<a href=chgrecord_out.asp?page="&pagecount&">βҳ<a>"
end if 
Response.Write  "&nbsp; ��"&page&"ҳ&nbsp;��"&pagecount&"ҳ ��"&rs.recordcount&"����¼"
end sub

%>