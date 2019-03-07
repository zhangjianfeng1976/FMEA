<!--#include file="../includes/db.asp"-->
<%
sql="select * from v_1_wogroup P Pivot (SUM(数量) FOR 工程 IN ([前],[调整],[仕上],[包装],[检查],[ISU],[LSU],[DLP],[DRUM],[FUSER],[IH],[中转],[加工],[MK],[RPS],[车手],[组立配膳],[KIT配膳],[码头],[日货],[组立工程外],[KIT工程外],[提前加工])) AS pvt "

'sql="select * from v_1_wogroup "
call openDB()
rs.open sql,conn,1,1
if not rs.eof then
rs.pagesize = 10000
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
<title></title>
</head>

<body>
<table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#5A8BCE">
<tr bgcolor="#E6F7FF">
<td height="30" align='center' bgcolor="#E6F7FF" colspan=27>
<b>LINE机种别(工程配置表)各工程数量</b></td>
</tr> 
<tr bgcolor="#FCFCFC" align="center" heigh ="25">
<td width="2%">No.</td>
<td width="5%">LINE</td>
<td width="10%">机种</td>
<td width="2%">前</td>
<td width="2%">调整</td>
<td width="2%">仕上</td>
<td width="2%">包装</td>
<td width="2%">检查</td>
<td width="2%">加工</td>
<td width="2%">提前加工</td>
<td width="2%">ISU</td>
<td width="2%">LSU</td>
<td width="2%">DLP</td>
<td width="2%">DRUM</td>
<td width="2%">FUSER</td>
<td width="2%">IH</td>
<td width="2%">中转</td>
<td width="2%">MK</td>
<td width="2%">RPS</td>
<td width="2%">车手</td>
<td width="2%">组立配膳</td>
<td width="2%">KIT配膳</td>
<td width="2%">码头</td>
<td width="2%">日货</td>
<td width="2%">组立工程外</td>
<td width="2%">KIT工程外</td>
<td width="4%" bgcolor='#FCFCDD'>合计</td>
</tr>

<% call MainBody() %>


</table>
<% Call CloseDB() %>
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
   	response.write"<tr bgcolor='#FCFCFC'>"
        response.write"<td>"&n&"</td>"
        response.write"<td>"&rs("line")&"</td>"
        response.write"<td>"&rs("机种")&"</td>"		
	response.write"<td>"&rs("前")&"</td>"
        i1 = rs("前")
        if isnull(i1) then 
           i1 = 0
         end if
        a1=a1+ i1
        response.write"<td>"&rs("调整")&"</td>"
        i2 = rs("调整")
        if isnull(i2) then 
           i2 = 0
        end if
        a2=a2+ i2    
        response.write"<td>"&rs("仕上")&"</td>"
        i3 = rs("仕上")
        if isnull(i3) then 
           i3 = 0
         end if
         a3=a3+ i3 
        response.write"<td>"&rs("包装")&"</td>"   
        i4 = rs("包装")
        if isnull(i4) then 
           i4 = 0
         end if
         a4=a4+ i4  
        response.write"<td>"&rs("检查")&"</td>"
        i5 = rs("检查")
        if isnull(i5) then 
           i5 = 0
         end if
         a5=a5+ i5 
        response.write"<td>"&rs("加工")&"</td>"
        i6 = rs("加工")
        if isnull(i6) then 
           i6 = 0
         end if
        a6=a6 + i6 
        response.write"<td>"&rs("提前加工")&"</td>"
        i7 = rs("提前加工")
        if isnull(i7) then 
           i7 = 0
         end if
         a7 =a7 + i7 
        response.write"<td>"&rs("ISU")&"</td>"
        i8 = rs("ISU")
        if isnull(i8) then 
           i8 = 0
        end if
         a8 =a8 + i8
        response.write"<td>"&rs("LSU")&"</td>"
        i9 = rs("LSU")
        if isnull(i9) then 
           i9 = 0
        end if
         a9 =a9 + i9
        response.write"<td>"&rs("DLP")&"</td>"
        i10 = rs("DLP")
        if isnull(i10) then 
           i10 = 0
        end if
         a10 =a10 + i10
        response.write"<td>"&rs("DRUM")&"</td>"
        i11 = rs("DRUM")
        if isnull(i11) then 
           i11 = 0
        end if
         a11 =a11 + i11
        response.write"<td>"&rs("FUSER")&"</td>"
        i12 = rs("FUSER")
        if isnull(i12) then 
           i12 = 0
        end if
         a12 =a12 + i12
        response.write"<td>"&rs("IH")&"</td>"
        i13 = rs("IH")
        if isnull(i13) then 
           i13 = 0
        end if
         a13 =a13 + i13
        response.write"<td>"&rs("中转")&"</td>"
        i14 = rs("中转")
        if isnull(i14) then 
           i14 = 0
        end if
         a14 =a14 + i14
         response.write"<td>"&rs("MK")&"</td>"
        i15 = rs("MK")
        if isnull(i15) then 
           i15 = 0
        end if
         a15 =a15 + i15
        response.write"<td>"&rs("RPS")&"</td>"
        i16 = rs("RPS")
        if isnull(i16) then 
           i16 = 0
        end if
         a16 =a16 + i16
        response.write"<td>"&rs("车手")&"</td>"
        i17 = rs("车手")
        if isnull(i17) then 
           i17 = 0
        end if
         a17 =a17 + i17
        response.write"<td>"&rs("组立配膳")&"</td>"
        i18 = rs("组立配膳")
        if isnull(i18) then 
           i18 = 0
        end if
         a18 =a18 + i18
        response.write"<td>"&rs("KIT配膳")&"</td>"
        i19 = rs("KIT配膳")
        if isnull(i19) then 
           i19 = 0
        end if
         a19 =a19 + i19
        response.write"<td>"&rs("码头")&"</td>"
        i20 = rs("码头")
        if isnull(i20) then 
           i20 = 0
        end if
         a20 =a20 + i20
        response.write"<td>"&rs("日货")&"</td>"
        i21 = rs("日货")
        if isnull(i21) then 
           i21 = 0
        end if
         a21 =a21 + i21
        response.write"<td>"&rs("组立工程外")&"</td>"
        i22 = rs("组立工程外")
        if isnull(i22) then 
           i22 = 0
        end if
         a22 =a22 + i22
        response.write"<td>"&rs("KIT工程外")&"</td>"
        i23 = rs("KIT工程外")
        if isnull(i23) then 
           i23 = 0
        end if
         a23 =a23 + i23
        response.write"<td bgcolor='#FCFCDD'>"&i1+i2+i3+i4+i5+i6+i7+i8+i9+i10+i11+i12+i13+i14+i15+i16+i17+i18+i19+i20+i21+i22+i23&"</td>"
    	response.write"</tr>"
	rs.MoveNext
n=n+1
wend


        response.write"<tr  bgcolor='#F59595'>"
        response.write"<td></td>"	
        response.write"<td>[<a href='javascript:history.back()'>返回</a>]</td>"	
  	response.write"<td>合计</td>"      
        response.write"<td>"&a1&"</td>"
        response.write"<td>"&a2&"</td>"
        response.write"<td>"&a3&"</td>"
        response.write"<td>"&a4&"</td>"
        response.write"<td>"&a5&"</td>"
        response.write"<td>"&a6&"</td>"
        response.write"<td>"&a7&"</td>"
        response.write"<td>"&a8&"</td>"
        response.write"<td>"&a9&"</td>"
        response.write"<td>"&a10&"</td>"
        response.write"<td>"&a11&"</td>"
        response.write"<td>"&a12&"</td>"
        response.write"<td>"&a13&"</td>"
        response.write"<td>"&a14&"</td>"
        response.write"<td>"&a15&"</td>"
        response.write"<td>"&a16&"</td>"
        response.write"<td>"&a17&"</td>"
        response.write"<td>"&a18&"</td>"
        response.write"<td>"&a19&"</td>"
        response.write"<td>"&a20&"</td>"
        response.write"<td>"&a21&"</td>"
        response.write"<td>"&a22&"</td>"
        response.write"<td>"&a23&"</td>"
        response.write"<td>"&a1+a2+a3+a4+a5+a6+a7+a8+a9+a10+a11+a12+a13+a14+a15+a16+a17+a18+a19+a20+a21+a22+a23&"</td>"
	response.write"</tr>"

End Sub
%>
