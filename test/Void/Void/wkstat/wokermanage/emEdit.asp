<!--#include file="../includes/db.asp"-->
<%
if request.cookies("leader")="" then
response.redirect "../default.asp"
end if

ID=request("ID")
sql="SELECT * FROM wt_employbase where ID="&ID
call openDB()
rs.open sql,conn,1,1
ID               = rs("ID")
CardID           = rs("CardID")
Name             = rs("Name")
Vocation         = rs("Vocation")
call closeDB()

Select Case Request("OK") 
      Case "ȷ��" 
        if request("CReason") =""   then
        response.write "�ո�δ�����[<a href='javascript:history.back()'>����</a>]" 
        response.end
       end if  
       dim  pn
       pn=year(request("dtae"))&right("0"&month(request("dtae")),2)
       sql= "Insert into wt_outwdairy(dtae,pno,cardid,creason,leader,truename,intime) values ('" _
              &request("dtae")&"','" _
              &pn&"','" _
              &request("CardID")&"','" _
              &request("Creason")&"','" _
              &request.cookies("leader")&"','" _
              &Request.ServerVariables("REMOTE_HOST")&"','"&NOW()&"')" 
      call openDB()
      conn.execute(sql)
      call closeDB()
      response.redirect "../chgrecord.asp"
End Select
%>
<html>
<head>
<title>��λ��Ա����</title>
</head>

<body>
<form action="emEdit.asp" method="POST" >
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=11>
<b>��Ա��������</b>[<%=request.cookies("leader") %>]</td>
</tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="4%" height="25">����</td>
<td width="4%">����</td>
<td width="4%">ְ��</td>
<td width="28%" colspan =4 >��������</td>
</tr>
<tr bgcolor="#FCFCFC" align="center">
<input type="hidden" value="<%=id %>" name="id">
<input type="hidden" value="<%=cardid %>" name="cardid">
<td width="4%" height="22"><%=cardid %></td>
<td width="4%"><%=name %></td>
<td width="4%"><%=vocation %></td>
<td width="20%">
<input type="radio" name="CReason" value="���">���&nbsp;&nbsp; 
<input type="radio" name="CReason" value="ȱ��">ȱ��&nbsp;&nbsp;
<input type="radio" name="CReason" value="����">����&nbsp;&nbsp;
<input type="radio" name="CReason" value="����">����&nbsp;&nbsp;
<input type="radio" name="CReason" value="���">���&nbsp;&nbsp;
<input type="radio" name="CReason" value="ɥ��">ɥ��&nbsp;&nbsp;
<input type="radio" name="CReason" value="������">������&nbsp;&nbsp;
<input type="radio" name="CReason" value="��ְ">��ְ&nbsp;&nbsp;
</td>
<td width="4%">
<input class="notline" name="Dtae"  size="10" value="<% =date() %>"
</td>
<td width="4%" >
<input type="submit" value="ȷ��" name="OK" class="smallinput"></td>
</tr>
  
</table>
</form>
</body>
</html>