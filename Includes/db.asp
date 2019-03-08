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
 
 ArWkCt  = array("打螺丝","装入","张贴","结线","线处理","涂布","清扫","设定","检查","同捆","-")
 
 ArPt    = array("框架件","驱动件","光学件","电气件","引导件","导通件","组立部件", _
  "弹簧","轴承","海绵","胶片","卡环","标签","螺丝","同捆品","润滑材","耗材","其他" ,"-")
%>

<%
'定义数据库操作的各个变量'
dim sql
dim rs
dim conn
'打开数据库'
sub openDB()
	set conn=server.createobject("ADODB.Connection") 
	conn.open"Driver={SQL Server};"_ 
	&"Server=cn-gongguan;" _
	&"Database=fmea;" _
	&"Uid=zhang;" _
	&"Pwd=sheet"
	set rs=server.createobject("ADODB.Recordset")
	
end sub

'关闭数据库'
sub closeDB()
	If IsObject(conn) Then
		if not(conn is nothing) then
			set rs=nothing
			
			conn.close
			set conn=nothing
		end if
	End If
end sub

'身份验证'
sub insureID()
	if session("Power")="" then
		call closeDB()
        response.write "<script language='JavaScript' type='text/javascript'></script>"
		response.redirect "../../reload.htm"
	end if
end sub

'当操作没有取得应有的权力时，显示错误信息'
sub noRight()
	response.write "<div height=50 align=center valign=center>您没有进行此操作的权限,操作被取消.</div>"
	call closeDB()
	response.end
end sub

'系统产生错误时，显示错误信息'
sub trigErr()
        call closeDB()
	    response.write "<div height=50 align=center valign=center>有错误发生,操作被取消."
        response.write "[<a href='javascript:history.back()'>返回</a>]"	
        response.write "</div>"
        response.end
end sub

'将从客户端的文本数据存入到数据库的时候不能有撇号，每次Insert和Update之前，都要先替换掉撇号'
function replacePrime(strItem)
	if strItem="" then
		call trigErr()
	end if
	replacePrime=replace(strItem,"'","`")
end function

'从数据库读取文本后，在发送到客户端之前应该将把撇号替换回来'
function replaceBack(strItem)
	if strItem="" then
		call trigErr()
	end if
	replaceBack=replace(strItem,"#Rep_PRIME_lace#","'")
end function

%>

