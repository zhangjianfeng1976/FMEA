<!--��λ��¼׷��-->
<!--#include file="../includes/db.asp"-->
<%
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
       dim  pn
       pn=year(request("dtae"))&right("0"&month(request("dtae")),2)
       sql= "Insert into wt_wdairy(dtae,pno,lineid,creason,cardid,truename,intime) values ('" _
              &request("dtae")&"','" _
              &pn&"','" _
              &request("lineid")&"','" _
              &request("Creason")&"','" _
              &request("CardID")&"','"&Request.ServerVariables("REMOTE_HOST")&"','"&NOW()&"')" 
      call openDB()
      conn.execute(sql)
      call closeDB()
      response.redirect "../chgrecord.asp"
End Select
%>

<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<style type="text/css">
<!--
.style1 {color: #FF0000}
.mouseOut
{
 font-size:12px;
 background: #708090;
 color: #FFFAFA;
}
.mouseOver
{
 font-size:12px;
 background: #FFFAFA;
 color: #000000;
}
-->
</style>
     <script type="text/javascript" language="javascript">
        var xmlHttp;
        var completeDiv;
        var inputField;
        var nameTable;
        var nameTableBody;
        var flag=false;
        function createXMLHttpRequest() {
            if (window.ActiveXObject) {
                xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
            }
            else if (window.XMLHttpRequest) {
                xmlHttp = new XMLHttpRequest();                
            }
        }
        
        function setflag(){
            flag = true;
        }
        
        function DisSelect()
        {
            if(flag==false)
            document.getElementById("popup").style.display="none";
        }
        function initVars() {
            inputField = document.getElementById("cardid");            
            nameTable = document.getElementById("name_table");
            completeDiv = document.getElementById("popup");
            nameTableBody = document.getElementById("name_table_body");
            document.getElementById("popup").style.display="block";
        }
        function findNames() {
            initVars();
            if (inputField.value.length > 0)
    {
                createXMLHttpRequest();
                var url = "lc_incheck.asp?cat=0&sc=" + inputField.value;  
                xmlHttp.open("GET", url, true);
                xmlHttp.onreadystatechange = callback;
                xmlHttp.send(null);
            }
    else
   {
                clearNames();
            }
        }
        function callback() {
            if (xmlHttp.readyState == 4) {
                if (xmlHttp.status == 200)
                {
     try
                    {
                        var name = xmlHttp.responseXML.getElementsByTagName("name")[0].firstChild.data;
                    }
                    catch(e)
                    {
                        document.getElementById("popup").style.display="none";
                        clearNames();
                    }
                    setNames(xmlHttp.responseXML.getElementsByTagName("content"));
                } 
                else if (xmlHttp.status == 204)
                {
                    clearNames();
                }
            }
        }
        
        function setNames(the_names) {            
            clearNames();
            var size = the_names.length;
            setOffsets();
            var row,cell,spans;
            for (var i = 0; i < size; i++) {
                //var nextNode = the_names[i].firstChild.data;
    var e = the_names[i];
    //ȡ���ӽڵ����ݣ�����װ��Ϊ����
    var nextNode=e.getElementsByTagName("name")[0].firstChild.data;
    //����tr��td��spanԪ��
                row  =document.createElement("tr");
                cell =document.createElement("td");
    //spans=document.createElement("span");
    //����cell����
                cell.onmouseout = function() {this.className='mouseOver'; flag=false;};
                cell.onmouseover = function() {this.className='mouseOut'; flag=true;};
                cell.setAttribute("bgcolor","#FFFAFA");
                cell.setAttribute("border","0");
                //cell.setAttribute("onmouseover","setflag()");
                cell.onclick = function() { populateName(this); };
    //��nextNode��ӵ�td
    var txtName = document.createTextNode(nextNode);
    cell.appendChild(txtName);
    //װ���������ݵ�spanԪ��
    row.appendChild(cell);
                nameTableBody.appendChild(row);
            }
        }
        function setOffsets() {
            var end = inputField.offsetWidth;
            var left = calculateOffsetLeft(inputField);
            var top = calculateOffsetTop(inputField) + inputField.offsetHeight;
            completeDiv.style.border = "black 1px solid";
            completeDiv.style.left = left + "px";
            completeDiv.style.top = top + "px";
            nameTable.style.width="400px";
        }
        
        function calculateOffsetLeft(field) {
          return calculateOffset(field, "offsetLeft");
        }
        function calculateOffsetTop(field) {
          return calculateOffset(field, "offsetTop");
        }
        function calculateOffset(field, attr) {
          var offset = 0;
          while(field) {
            offset += field[attr]; 
            field = field.offsetParent;
          }
          return offset;
        }
        function populateName(cell) {
    //������ݵ�webҳ�棬cell---->td����
            inputField.value = cell.firstChild.nodeValue;
            clearNames();
        }
        //����б�����
        function clearNames() {
            var ind = nameTableBody.childNodes.length;
            for (var i = ind - 1; i >= 0 ; i--) {
                 nameTableBody.removeChild(nameTableBody.childNodes[i]);
            }
            completeDiv.style.border = "none";
        }
</script>
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
<td width="8%" height="25">��λԭ��</td>
<td width="8%" colspan = 4 >
<input type="radio" name="CReason" value="���">���&nbsp;&nbsp; 
<input type="radio" name="CReason" value="ȱ��">ȱ��&nbsp;&nbsp;
<input type="radio" name="CReason" value="����">����&nbsp;&nbsp;
<input type="radio" name="CReason" value="����">����&nbsp;&nbsp;
<input type="radio" name="CReason" value="��ְ">��ְ&nbsp;&nbsp;
<input type="radio" name="CReason" value="ѧλ">ѧλ&nbsp;&nbsp;
<input type="radio" name="CReason" value="��λ">��λ</td>
<td width="8%">
<input class="notline" name="Dtae"  size="10" value="<% =date() %>"
</td>
<td width="8%" colspan =3 >
���Ϊ:�����߹���<input name="Cardid" type="text" class="smallInput"  size="17" id="cardid" onKeyUp="findNames();">
<input type="submit" value="ȷ��" name="OK" class="smallinput"></td>
</tr>

  
</table>
</form>
</body>
</html>