<!--��λ��¼׷��-->
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
Model              = rs("����")
WoGroup            = rs("����")
WorkStation        = rs("��λ")
OriginWorker       = rs("����")
wkname             = rs("����")
Changedate         = rs("������")
Wkstatus           = rs("��Ч")
call closeDB()

Select Case Request("OK") 
      Case "ȷ��" 
        if request("CReason") =""  or request("CardID") = "" or len(request("CardID"))<> 6 then
        response.write "�ո�δ�����[<a href='javascript:history.back()'>����</a>]" 
        response.end
        end if  
        if  request("CardID") <> "999999" and request("WKcontent") ="" then
        response.write "��ѡ����ҵ����[<a href='javascript:history.back()'>����</a>]" 
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
      if request("Creason") = "֧Ԯ" then
      Response.Redirect "lcbrowse.asp?cat=0&sc="&request("CardID")
      else 
      Response.Redirect "lcbrowse.asp?cat=0&sc="&aa
      end if
End Select
%>
<html>
<head>
<title>��λ��Ա����</title>
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
<b>��λ���</b>
</td></tr> 
<tr bgcolor="#FCFCFC" align="center">
<td width="8%" height="25">Line</td>
<td width="8%">����</td>
<td width="8%">����</td>
<td width="4%">��λ</td>
<td width="8%">����</td>
<td width="8%">����</td>
<td width="8%">��Ч</td>
<td width="8%">������</td>
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
<td width="8%" height="25">����λԭ��</td>
<td width="8%" colspan = 7 >
<input type="radio" name="CReason" value="���">���&nbsp;&nbsp; 
<input type="radio" name="CReason" value="ȱ��">ȱ��&nbsp;&nbsp;
<input type="radio" name="CReason" value="����">����&nbsp;&nbsp;
<input type="radio" name="CReason" value="����">����&nbsp;&nbsp;
<input type="radio" name="CReason" value="���">���&nbsp;&nbsp;
<input type="radio" name="CReason" value="ɥ��">ɥ��&nbsp;&nbsp;
<input type="radio" name="CReason" value="������">������&nbsp;&nbsp;
<input type="radio" name="CReason" value="��ְ">��ְ&nbsp;&nbsp;<br>
<input type="radio" name="CReason" value="Ԥ��">Ԥ��&nbsp;&nbsp;
<input type="radio" name="CReason" value="Ԥ�Ӽ�λ">Ԥ�Ӽ�λ&nbsp;&nbsp;
<input type="radio" name="CReason" value="֧Ԯ">֧Ԯ�������г��ڣ���������λ��λ��ѧλ��&nbsp;&nbsp;
<input type="radio" name="CReason" value="ȷ����">ȷ���ڣ�����ԭ��ҵ�߱����������ȷ�����ڣ�&nbsp;&nbsp;

</td>
</tr><tr  bgcolor="#FCFCFC">
<td  height="22">����</td>
<td colspan = 7>
<input class="notline" name="Dtae"  size="10" value="<% =date() %>"
</td>
</tr><tr  bgcolor="#FCFCFC">
<td  height="22">����Ϊ</td>
<td colspan =7 >
�����߹��ţ�<input name="Cardid" type="text" class="smallInput"  size="17" id="cardid" onkeyup="aa()" onblur="aa()">
</td>
</tr><tr  bgcolor="#FCFCFC">
<td>��ҵ����</td>
<td colspan =7  height="22">
<input type="radio" name="WKcontent" value="ѧλ">ѧλ&nbsp;&nbsp;
<input type="radio" name="WKcontent" value="��λ">��λ&nbsp;&nbsp;
</td></tr>
</tr><tr  bgcolor="#FCFCFC">
<td  height="22" colspan =8 align ="center"><input type="submit" value="ȷ��" name="OK" class="smallinput"></td>
</tr><tr  bgcolor="#FCFCFC">
<td  height="22" colspan =8 >
˵����<br>
1.����֮ǰ����λ��ѧλ�����ۣ��ڼ�¼���޷�������ǰ�����Ա���Ǻ������Ա�����ԣ��ڱ����������ȷ��<br>
2.�����λ����Ա�����г��ڣ������˱�λ������λ��ѧλ�ˣ���ʱ��ͳһ��֧Ԯ���б������ŵ���Ա����ҵ������ѧλ����λ����ʾ��<br>
3.���֧Ԯ��Ա�Ĺ�λ����Ҫ����λ�Ļ�����999999��ʾ�Ļ�����Ȼ�ز���Ҫѡ��ѧλ����λ�ˡ�<br>
4.Ԥ�룺��λ����Ԥ����ְ������Ա������ѧλ�������<br>
5.Ԥ�Ӽ�λ��LINE����Ӽ�λ����������λ�в���ѧλ������� <br>
6.ȷ������ָ��������ԭ��ҵ�߱����������ȷ������Ҫ����ȷ�ϵ���Ա��
</td>
</tr>



<!--
<tr><td colspan =8>
<a href="lc_incheck.asp?cat=0&sc="&av  target ="container1">��֤��Ա</a>
</td></tr>
<tr><td colspan =8>
<iframe src='' name='container1' width='100%' marginwidth='1' height='100' marginheight='1' frameborder='0' scrolling='none'></iframe>
</td></tr>
 --> 
</table>
</form>
</body>
</html>