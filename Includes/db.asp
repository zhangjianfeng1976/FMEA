<!--#include file="Style.css"-->

<%
 ArWkThm = array("1-1-1","1-2-1","1-3-1", _
 "2-1-1","2-2-1","2-3-1","2-4-1", _
 "3-1-1","3-1-2","3-1-3","3-2-1","3-3-1","3-3-2","3-4-1","3-5-1","3-6-1","3-7-1", "3-8-1","3-9-1","3-9-2","3-10-1", _
 "3-11-1","3-11-2","3-12-1","3-12-2", "3-13-1", _
 "4-1-1","4-1-2","4-1-3","4-2-2","4-3-1","4-3-2","4-4-1","4-4-2", _
 "5-1-1","5-2-1","5-3-1", _
 "6-1-1","6-1-2","6-2-1","6-3-1", _
 "7-1-1","7-2-1","7-3-1","7-4-1","7-5-1","7-6-1","7-7-1", _
 "8-1-1","8-2-1","8-3-1","8-4-1", _
 "9-1-1", "9-2-1","9-3-1", _
 "10-1-1","10-2-1","10-3-1","10-4-1","10-5-1","10-6-1","10-6-2","10-7-1", _
 "11-1-1","11-2-1",  _
 "12-1-1","12-2-1","12-2-2","12-2-3","12-2-4", _ 
 "13-1-1", _ 
 "14-1-1","14-2-1")
 
 ArWkCt  = array("����˿","װ��","����","����","�ߴ���","Ϳ��","��ɨ","�趨","���","ͬ��","-")
 
 ArPt    = array("��ܼ�","������","��ѧ��","������","������","��ͨ��","��������", _
  "����","���","����","��Ƭ","����","��ǩ","��˿","ͬ��Ʒ","�󻬲�","�Ĳ�","����" ,"-")
%>

<%
'�������ݿ�����ĸ�������'
dim sql
dim rs
dim conn
'�����ݿ�'
sub openDB()
	set conn=server.createobject("ADODB.Connection") 
	conn.open"Driver={SQL Server};"_ 
	&"Server=cn-gongguan;" _
	&"Database=fmea;" _
	&"Uid=zhang;" _
	&"Pwd=sheet"
	set rs=server.createobject("ADODB.Recordset")
	
end sub

'�ر����ݿ�'
sub closeDB()
	If IsObject(conn) Then
		if not(conn is nothing) then
			set rs=nothing
			
			conn.close
			set conn=nothing
		end if
	End If
end sub

'�����֤'
sub insureID()
	if session("Power")="" then
		call closeDB()
        response.write "<script language='JavaScript' type='text/javascript'></script>"
		response.redirect "../../reload.htm"
	end if
end sub

'������û��ȡ��Ӧ�е�Ȩ��ʱ����ʾ������Ϣ'
sub noRight()
	response.write "<div height=50 align=center valign=center>��û�н��д˲�����Ȩ��,������ȡ��.</div>"
	call closeDB()
	response.end
end sub

'ϵͳ��������ʱ����ʾ������Ϣ'
sub trigErr()
        call closeDB()
	    response.write "<div height=50 align=center valign=center>�д�����,������ȡ��."
        response.write "[<a href='javascript:history.back()'>����</a>]"	
        response.write "</div>"
        response.end
end sub

'���ӿͻ��˵��ı����ݴ��뵽���ݿ��ʱ������Ʋ�ţ�ÿ��Insert��Update֮ǰ����Ҫ���滻��Ʋ��'
function replacePrime(strItem)
	if strItem="" then
		call trigErr()
	end if
	replacePrime=replace(strItem,"'","`")
end function

'�����ݿ��ȡ�ı����ڷ��͵��ͻ���֮ǰӦ�ý���Ʋ���滻����'
function replaceBack(strItem)
	if strItem="" then
		call trigErr()
	end if
	replaceBack=replace(strItem,"#Rep_PRIME_lace#","'")
end function

%>

