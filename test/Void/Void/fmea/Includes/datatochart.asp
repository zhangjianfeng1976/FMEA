<%
'������ͼ�λ����'
'dataArray��ά����'
'virtualFilePath���ͼ���ļ���(����·��)'
'nType��ʾ����'
Dim initType
Sub ExportPicture(dataArray,virtualFilePath,nType)
Dim excelapp ' As New excel.Application'
Dim excelwbk ' As excel.Workbook'
Dim excelcht ' As excel.Chart'
Dim excelsht 'As excel.Worksheet'
Dim idx,idy,ftype,usedData,totalcount,count:count = 1
On Error Resume Next

Set excelapp = Server.CreateObject("Excel.Application")
Set excelwbk = excelapp.Workbooks.Add()
Set excelcht = excelwbk.Charts.Add()
Set excelsht = excelwbk.Worksheets.Add()
If UCase(Right(virtualFilePath,4)) = ".JPG" Or UCase(Right(virtualFilePath,4)) = ".JPEG" Then
ftype = "jpg"
Else
ftype = "gif"
End If
initType = nType
For idx=LBound(dataArray,1) To UBound(dataArray,1)
For idy=LBound(dataArray,2) To UBound(dataArray,2)
excelsht.Cells(idx+1,idy+1) = dataArray(idx,idy)
Next
Next

Set usedData = excelsht.usedRange
excelcht.SeriesCollection.Add usedData

excelcht.HasLegend = True
excelcht.HasTitle = True
'excelcht.ChartTitle.Caption = "����Ա���ֲ�ͼ"'
excelcht.ApplyCustomType nType
excelcht.Export Server.Mappath(virtualFilePath), ftype
excelsht.Close False
excelwbk.Close False
Set usedData = Nothing
Set excelcht = Nothing
Set excelwbk = Nothing
Set excelapp = Nothing
End Sub
%>
<Select name="sel" Onchange="changePict()">
<Option value="51">��ά����ͼ</Option><!--xlColumnClustered
<Option value="52">xlColumnStacked</Option>
<Option value="53">xlColumnStacked100</Option>-->
<Option value="54">��ά��״ͼ</Option><!--xl3DColumnClustered
<Option value="55">xl3DColumnStacked</Option>
<Option value="56">xl3DColumnStacked100</Option>-->
<Option value="57">��ά����ͼ</Option><!--xlBarClustered
<Option value="58">xlBarStacked</Option>
<Option value="59">xlBarStacked100</Option>-->
<Option value="60">��ά��״ͼ</Option><!--xl3DBarClustered
<Option value="61">xl3DBarStacked</Option>
<Option value="62">xl3DBarStacked100</Option>-->
<Option value="63">����ͼ</Option><!--xlLineStacked
<Option value="64">xlLineStacked100</Option>
<Option value="65">xlLineMarkers</Option>
<Option value="66">xlLineMarkersStacked</Option>
<Option value="67">xlLineMarkersStacked100</Option>
<Option value="68">xlPieOfPie</Option>
<Option value="69">xlPieExploded</Option>
<Option value="70">xl3DPieExploded</Option>
<Option value="71">xlBarOfPie</Option>-->
<Option value="72">����ͼ</Option><!--xlXYScatterSmooth
<Option value="73">xlXYScatterSmoothNoMarkers</Option>
<Option value="74">xlXYScatterLines</Option>
<Option value="75">xlXYScatterLinesNoMarkers</Option>-->
<Option value="76">�������ͼ</Option><!--xlAreaStacked
<Option value="77">xlAreaStacked100</Option>
<Option value="78">xl3DAreaStacked</Option>
<Option value="79">xl3DAreaStacked100</Option>
<Option value="80">xlDoughnutExploded</Option>
<Option value="81">xlRadarMarkers</Option>
<Option value="82">xlRadarFilled</Option>
<Option value="83">xlSurface</Option>
<Option value="84">xlSurfaceWireframe</Option>
<Option value="85">xlSurfaceTopView</Option>
<Option value="86">xlSurfaceTopViewWireframe</Option>
<Option value="15">xlBubble</Option>
<Option value="87">xlBubble3DEffect</Option>
<Option value="88">xlStockHLC</Option>
<Option value="89">xlStockOHLC</Option>
<Option value="90">xlStockVHLC</Option>
<Option value="91">xlStockVOHLC</Option>-->
<Option value="92">����Բ��ͼ</Option><!--xlCylinderColClustered
<Option value="93">xlCylinderColStacked</Option>
<Option value="94">xlCylinderColStacked100</Option>-->
<Option value="95">����Բ��ͼ</Option><!--xlCylinderBarClustered
<Option value="96">xlCylinderBarStacked</Option>
<Option value="97">xlCylinderBarStacked100</Option>
<Option value="98">xlCylinderCol</Option>
<Option value="99">xlConeColClustered</Option>
<Option value="100">xlConeColStacked</Option>
<Option value="101">xlConeColStacked100</Option>
<Option value="102">xlConeBarClustered</Option>
<Option value="103">xlConeBarStacked</Option>
<Option value="104">xlConeBarStacked100</Option>
<Option value="105">xlConeCol</Option>
<Option value="106">xlPyramidColClustered</Option>
<Option value="107">xlPyramidColStacked</Option>
<Option value="108">xlPyramidColStacked100</Option>
<Option value="109">xlPyramidBarClustered</Option>
<Option value="110">xlPyramidBarStacked</Option>
<Option value="111">xlPyramidBarStacked100</Option>
<Option value="112">xlPyramidCol</Option>
<Option value="-4100">xl3DColumn</Option>
<Option value="4">xlLine</Option>
<Option value="-4101">xl3DLine</Option>-->
<Option value="-4102">��ͼ</Option><!--xl3DPie-->
<Option value="5">����ͼ</Option><!--xlPie
<Option value="-4169">xlXYScatter</Option>
<Option value="-4098">xl3DArea</Option>
<Option value="1">xlArea</Option>-->
<Option value="-4120">Բ��ͼ</Option><!--xlDoughnut-->
<Option value="-4151">�״�ͼ</Option><!--xlRadar-->
</Select>
<Script language=javascript>
function initMenu(formobj)
{
var nType="<%=initType%>";
var i;
for(i=0;i<formobj.sel.options.length;i++)
{
if(formobj.sel.options[i].value==nType)
{
formobj.sel.options[i].selected=true;
break;
}
}
}
</Script>
