<!--#include file="Style.css"-->
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

<%
function tablex(total,table_x,table_y,all_width,all_height,line_no)
'��������(���ݵ����飬�����꣬�����꣬ͼ��Ŀ�ȣ�ͼ��ĸ߶�,��������)
'��ASP��������ͼ����2��������ͼ
'���ߣ�����(Passwordgm) QQ:25968152 MSN:passwordgm@sina.com Email:passwordgm@sina.com
'���˷ǳ�Ը���ASP,VML,FLASH�İ�������HTTP://topclouds.126.com���н�����̽��
'�汾1.0 ����޸����� 2003-8-11
'�ǳ���л��ʹ���������������ʹ�ú�ת��ʱ������Ȩ��Ϣ�����Ƕ����߹�������õ����ء�

line_color="#69f"
left_width=50
total_no=ubound(total,1)

temp1=0
for i=1 to total_no
	for j=1 to line_no
		if temp1<total(i,j) then temp1=total(i,j)
	next
next
temp1=int(temp1)
if temp1>9 then
	temp2=mid(cstr(temp1),2,1)
	if temp2>4 then 
		temp3=(int(temp1/(10^(len(cstr(temp1))-1)))+1)*10^(len(cstr(temp1))-1)
	else
		temp3=(int(temp1/(10^(len(cstr(temp1))-1)))+0.5)*10^(len(cstr(temp1))-1)
	end if
else
	if temp1>4 then temp3=10 else temp3=5
end if
temp4=temp3
response.write "<v:rect id='_x0000_s1027' alt='' style='position:absolute;left:"&table_x+left_width&"px;top:"&table_y&"px;width:"&all_width&"px;height:"&all_height&"px;z-index:1' fillcolor='#9cf' stroked='f'><v:fill rotate='t' angle='15' focus='100%' type='gradient'/></v:rect>"
for i=0 to all_height-1 step all_height/5
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:1' from='"&table_x+left_width+length&"px,"&table_y+all_height-length-i&"px' to='"&table_x+all_width+left_width&"px,"&table_y+all_height-length-i&"px' strokecolor='"&line_color&"'/>"
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:1' from='"&table_x+(left_width-15)&"px,"&table_y+i&"px' to='"&table_x+left_width&"px,"&table_y+i&"px'/>"
	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x&"px;top:"&table_y+i&"px;width:"&left_width&"px;height:18px;z-index:4'>"
	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='right'>"&temp4&"</td></tr></table></v:textbox></v:shape>"
	temp4=temp4-temp3/5
next
response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:1' from='"&table_x+left_width&"px,"&table_y+all_height&"px' to='"&table_x+all_width+left_width&"px,"&table_y+all_height&"px'/>"
response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:1' from='"&table_x+left_width&"px,"&table_y&"px' to='"&table_x+left_width&"px,"&table_y+all_height&"px'/>"

dim line_code
redim line_code(line_no,5)
for i=1 to line_no
	line_temp=split(total(0,i),",")
	line_code(i,1)=line_temp(0)
	line_code(i,2)=line_temp(1)
	line_code(i,3)=line_temp(2)
	line_code(i,4)=line_temp(3)
	line_code(i,5)=line_temp(4)
next
for j=1 to line_no
	for i=1 to total_no-1
			x1=table_x+left_width+all_width*(i-1)/total_no
			y1=table_y+(temp3-total(i,j))*(all_height/temp3)
			x2=table_x+left_width+all_width*i/total_no
			y2=table_y+(temp3-total(i+1,j))*(all_height/temp3)
			response.write "<v:line id=""_x0000_s1025"" alt="""" style='position:absolute;left:0;text-align:left;top:0;z-index:1' from="""&x1&"px,"&y1&"px"" to="""&x2&"px,"&y2&"px"" coordsize=""21600,21600"" strokecolor="""&line_code(j,1)&""" strokeweight="""&line_code(j,2)&""">"
		select case line_code(j,3)
		case 1
		case 2
			response.write "<v:stroke dashstyle='1 1'/>"
		case 3
			response.write "<v:stroke dashstyle='dash'/>"
		case 4
			response.write "<v:stroke dashstyle='dashDot'/>"
		case 5
			response.write "<v:stroke dashstyle='longDash'/>"
		case 6
			response.write "<v:stroke dashstyle='longDashDot'/>"
		case 7
			response.write "<v:stroke dashstyle='longDashDotDot'/>"
		case else
		end select
		response.write "</v:line>"&CHR(13)
		select case line_code(j,4)
		case 1
		case 2
			response.write "<v:rect id=""_x0000_s1027"" style='position:absolute;left:"&x1-2&"px;top:"&y1-2&"px;width:4px;height:4px; z-index:5' fillcolor="""&line_code(j,1)&""" strokecolor="""&line_code(j,1)&"""/>"&CHR(13)
		case 3
			response.write "<v:oval id=""_x0000_s1026"" style='position:absolute;left:"&x1-2&"px;top:"&y1-2&"px;width:4px;height:4px;z-index:4' fillcolor="""&line_code(j,1)&""" strokecolor="""&line_code(j,1)&"""/>"&CHR(13)
		end select
	next
		select case line_code(j,4)
		case 1
		case 2
			response.write "<v:rect id=""_x0000_s1027"" style='position:absolute;left:"&x2-2&"px;top:"&y2-2&"px;width:4px;height:4px; z-index:5' fillcolor="""&line_code(j,1)&""" strokecolor="""&line_code(j,1)&"""/>"&CHR(13)
		case 3
			response.write "<v:oval id=""_x0000_s1026"" style='position:absolute;left:"&x2-2&"px;top:"&y2-2&"px;width:4px;height:4px;z-index:4' fillcolor="""&line_code(j,1)&""" strokecolor="""&line_code(j,1)&"""/>"&CHR(13)
		end select
next
	
for i=1 to total_no
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:1' from='"&table_x+left_width+all_width*(i-1)/total_no&"px,"&table_y+all_height&"px' to='"&table_x+left_width+all_width*(i-1)/total_no&"px,"&table_y+all_height+15&"px'/>"
	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width+all_width*(i-1)/total_no&"px;top:"&table_y+all_height&"px;width:"&all_width/total_no&"px;height:18px;z-index:4'>"
	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='left'>"&total(i,0)&"</td></tr></table></v:textbox></v:shape>"
next

tb_height=15
response.write "<v:rect id='_x0000_s1025' style='position:absolute;left:"&table_x+all_width+15&"px;top:"&table_y&"px;width:90px;height:"&line_no*tb_height+10&"px;z-index:4'/>"
for i=1 to line_no
	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+all_width+25&"px;top:"&table_y+5+(i-1)*tb_height&"px;width:100px;height:"&tb_height&"px;z-index:4'>"
	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='1' cellpadding='0' width='100%' height='100%'><tr><td align='left'>"&line_code(i,5)&"</td></tr></table></v:textbox></v:shape>"
	response.write "<v:rect id='_x0000_s1040' alt='' style='position:absolute;left:"&table_x+all_width+90&"px;top:"&table_y+5+(i-1)*tb_height+4&"px;width:5px;height:5px;z-index:4' fillcolor='"&line_code(i,1)&"'><v:fill color2='"&line_code(i,1)&"' rotate='t' focus='100%' type='gradient'/></v:rect>"
next

end function

function table2(total,table_x,table_y,all_width,all_height,line_no)
'��������(���ݵ����飬�����꣬�����꣬ͼ��Ŀ�ȣ�ͼ��ĸ߶�,��������)
'��ASP��������ͼ����2��������ͼ
'���ߣ�����(Passwordgm) QQ:25968152 MSN:passwordgm@sina.com Email:passwordgm@sina.com
'���˷ǳ�Ը���ASP,VML,FLASH�İ�������HTTP://topclouds.126.com���н�����̽��
'�汾1.0 ����޸����� 2003-8-11
'�ǳ���л��ʹ���������������ʹ�ú�ת��ʱ������Ȩ��Ϣ�����Ƕ����߹�������õ����ء�

line_color="#69f"
left_width=50
total_no=ubound(total,1)

temp1=0
for i=1 to total_no
	for j=1 to line_no
		if temp1<total(i,j) then temp1=total(i,j)
	next
next
temp1=int(temp1)
if temp1>9 then
	temp2=mid(cstr(temp1),2,1)
	if temp2>4 then 
		temp3=(int(temp1/(10^(len(cstr(temp1))-1)))+1)*10^(len(cstr(temp1))-1)
	else
		temp3=(int(temp1/(10^(len(cstr(temp1))-1)))+0.5)*10^(len(cstr(temp1))-1)
	end if
else
	if temp1>4 then temp3=10 else temp3=5
end if
temp4=temp3
response.write "<v:rect id='_x0000_s1027' alt='' style='position:absolute;left:"&table_x+left_width&"px;top:"&table_y&"px;width:"&all_width&"px;height:"&all_height&"px;z-index:-1' fillcolor='#9cf' stroked='f'><v:fill rotate='t' angle='-45' focus='100%' type='gradient'/></v:rect>"
for i=0 to all_height-1 step all_height/5
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+length&"px,"&table_y+all_height-length-i&"px' to='"&table_x+all_width+left_width&"px,"&table_y+all_height-length-i&"px' strokecolor='"&line_color&"'/>"
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+(left_width-15)&"px,"&table_y+i&"px' to='"&table_x+left_width&"px,"&table_y+i&"px'/>"
	response.write ""
	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x&"px;top:"&table_y+i&"px;width:"&left_width&"px;height:18px;z-index:1'>"
	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='right'>"&temp4&"</td></tr></table></v:textbox></v:shape>"
	temp4=temp4-temp3/5
next
response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width&"px,"&table_y+all_height&"px' to='"&table_x+all_width+left_width&"px,"&table_y+all_height&"px'/>"
response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width&"px,"&table_y&"px' to='"&table_x+left_width&"px,"&table_y+all_height&"px'/>"

dim line_code
redim line_code(line_no,5)
for i=1 to line_no
	line_temp=split(total(0,i),",")
	line_code(i,1)=line_temp(0)
	line_code(i,2)=line_temp(1)
	line_code(i,3)=line_temp(2)
	line_code(i,4)=line_temp(3)
	line_code(i,5)=line_temp(4)
next
for j=1 to line_no
	for i=1 to total_no-1
			x1=table_x+left_width+all_width*(i-1)/total_no
			y1=table_y+(temp3-total(i,j))*(all_height/temp3)
			x2=table_x+left_width+all_width*i/total_no
			y2=table_y+(temp3-total(i+1,j))*(all_height/temp3)
			response.write "<v:line id=""_x0000_s1025"" alt="""" style='position:absolute;left:0;text-align:left;top:0;z-index:1' from="""&x1&"px,"&y1&"px"" to="""&x2&"px,"&y2&"px"" coordsize=""21600,21600"" strokecolor="""&line_code(j,1)&""" strokeweight="""&line_code(j,2)&""">"
		select case line_code(j,3)
		case 1
		case 2
			response.write "<v:stroke dashstyle='1 1'/>"
		case 3
			response.write "<v:stroke dashstyle='dash'/>"
		case 4
			response.write "<v:stroke dashstyle='dashDot'/>"
		case 5
			response.write "<v:stroke dashstyle='longDash'/>"
		case 6
			response.write "<v:stroke dashstyle='longDashDot'/>"
		case 7
			response.write "<v:stroke dashstyle='longDashDotDot'/>"
		case else
		end select
		response.write "</v:line>"&CHR(13)
		select case line_code(j,4)
		case 1
		case 2
			response.write "<v:rect id=""_x0000_s1027"" style='position:absolute;left:"&x1-2&"px;top:"&y1-2&"px;width:4px;height:4px; z-index:2' fillcolor="""&line_code(j,1)&""" strokecolor="""&line_code(j,1)&"""/>"&CHR(13)
		case 3
			response.write "<v:oval id=""_x0000_s1026"" style='position:absolute;left:"&x1-2&"px;top:"&y1-2&"px;width:4px;height:4px;z-index:1' fillcolor="""&line_code(j,1)&""" strokecolor="""&line_code(j,1)&"""/>"&CHR(13)
		end select
	next
		select case line_code(j,4)
		case 1
		case 2
			response.write "<v:rect id=""_x0000_s1027"" style='position:absolute;left:"&x2-2&"px;top:"&y2-2&"px;width:4px;height:4px; z-index:2' fillcolor="""&line_code(j,1)&""" strokecolor="""&line_code(j,1)&"""/>"&CHR(13)
		case 3
			response.write "<v:oval id=""_x0000_s1026"" style='position:absolute;left:"&x2-2&"px;top:"&y2-2&"px;width:4px;height:4px;z-index:1' fillcolor="""&line_code(j,1)&""" strokecolor="""&line_code(j,1)&"""/>"&CHR(13)
		end select
next
	
for i=1 to total_no
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+all_width*(i-1)/total_no&"px,"&table_y+all_height&"px' to='"&table_x+left_width+all_width*(i-1)/total_no&"px,"&table_y+all_height+15&"px'/>"
	response.write ""
	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width+all_width*(i-1)/total_no&"px;top:"&table_y+all_height&"px;width:"&all_width/total_no&"px;height:18px;z-index:1'>"
	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='left'>"&total(i,0)&"</td></tr></table></v:textbox></v:shape>"
next

tb_height=15
response.write "<v:rect id='_x0000_s1025' style='position:absolute;left:"&table_x+all_width+15&"px;top:"&table_y&"px;width:90px;height:"&line_no*tb_height+10&"px;z-index:1'/>"
for i=1 to line_no
	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+all_width+25&"px;top:"&table_y+5+(i-1)*tb_height&"px;width:100px;height:"&tb_height&"px;z-index:1'>"
	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='1' cellpadding='0' width='100%' height='100%'><tr><td align='left'>"&line_code(i,5)&"</td></tr></table></v:textbox></v:shape>"
	response.write "<v:rect id='_x0000_s1040' alt='' style='position:absolute;left:"&table_x+all_width+90&"px;top:"&table_y+5+(i-1)*tb_height+4&"px;width:5px;height:5px;z-index:1' fillcolor='"&line_code(i,1)&"'><v:fill color2='"&line_code(i,1)&"' rotate='t' focus='100%' type='gradient'/></v:rect>"
next

end function
%>