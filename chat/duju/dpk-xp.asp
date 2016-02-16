<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<head>
<title>洗牌桌</title>
<META http-equiv="Content-Type" content="text/html; charset=big5">
<STYLE>
A{ TEXT-DECORATION: none;color:#ffffff}
.td1 {background-color:#FFFFFF; }
TD {FONT-SIZE: 9pt;cursor:hand; }
.tb2 { font-family: "新細明體", "serif"; font-size:13; line-height:100%;}
</STYLE>
</head>
<body bgcolor="#0099FF" text="#FFFFFF">
<%
namef=Trim(Request.QueryString("name"))
nickname=info(0)
if namef<>"" and namef<> nickname then
  response.write "<script Language='Javascript'>alert('不要看別人的牌!');location.href = 'javascript:history.go(-1)';</script>"
  response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="dpk.mdb"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'如果你的服務器支持jet4.0，請使用下載的連接方法，提高程序性能
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 
sql="select * from dpk where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)
if rs.eof and rs.bof then
rs.close
conn.close
response.write "<script Language='Javascript'>alert('您現在沒有參與牌局，請向你的對手使用\""/打撲克 1000 \""命令要求開局。');location.href = 'javascript:history.go(-1)';</script>"
response.end
else
dpk=rs("pk")
dpkto=rs("uto")
rs.close
conn.close
end if
dpk2=split(dpk,"|")
dpknumb=ubound(dpk2)
%>
<table border="0" align="center" cellpadding="3" cellspacing="5">
<%
nw=1
dim i
dim writeArr(15)

for i= 0 to dpknumb-1 
nw=paixupk(dpk2(i))
writeArr(nw)= writeArr(nw) & "|" & showpk(dpk2(i))
next 

writePPP=""
for j=1 to 15
writePPP=writePPP & writeArr(j)
next

dpk3=split(mid(writePPP,2),"|")

writestr=""
for i= 0 to dpknumb-1 step 3
writestr=writestr & "<tr>"
writestr=writestr & dpk3(i)
if i+1<dpknumb then writestr=writestr & dpk3(i+1)
if i+2<dpknumb then writestr=writestr & dpk3(i+2)
writestr=writestr & "</tr>" 
next
response.write writestr

%>
</table><BR>
<table borderColorDark='#E1E1E1' class="tb2" border="1" align="center" cellspacing="2" bordercolor="#993300" height="30" cellpadding="4">
  <tr> 
    <td onClick=javascript:s('') onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">重 選</font></td>
  </tr>
<tr> 
<td onClick=javascript:s('不要了') onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">讓 牌</font></td>
</tr>
<tr>
<td onClick=javascript:s('認輸了') onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">認  輸</font></td>
</tr>
<tr> 
<td onclick=javascript:window.open("dpkhelp.htm") height="17" onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">幫 助</font></td>
</tr>
</table>
<br>
<div align="center"><font size="2">□</font><font size="2"><%=dpknumb%>張牌□</font><font size="2"><br>
  □請點擊撲克□</font><br>
<script language="JavaScript">
parent.m.location.href="about:blank"

function s(list){
var sa=parent.f2.document.af.sytemp.value;
if(list=="認輸了"){parent.sw('[大家]');}else{parent.sw('[<%=dpkto%>]');}
if(sa==""||sa.charAt(0)!="/"||sa.charAt(1)!="發"||sa.charAt(2)!="牌"||sa.charAt(5)=="公"||list.charAt(0)=="公"||sa.charAt(4)!=" "||sa.charAt(5)=="不"||list.charAt(0)=="不"||list==""||sa=="/發牌$ "){sa="/發牌$ "}
else{sa=sa+"\+"}
parent.f2.document.af.sytemp.value=sa+list;
parent.f2.document.af.sytemp.focus();
//window.location.reload();
}
function mover(tb){
tb.borderColorLight='#993300';tb.borderColorDark='#E1E1E1';}
function mdown(tb,tttb){
TimerID=0;
tb.borderColorLight='#f6f6f6';tb.borderColorDark='#000000';
}
function mup(tb){
tb.borderColorLight='#993300';tb.borderColorDark='#E1E1E1';}
function mout(tb){
tb.borderColorLight='#993300';tb.borderColorDark='#E1E1E1';}
var loadtime=500;
</script>
<%
function showpk(pk)
dim wpk
wpk=replace(pk,"紅","<img src=dpk/1.GIF><br><font color=#FF0000>")
wpk=replace(wpk,"黑","<img src=dpk/2.GIF><br><font color=#000000>")
wpk=replace(wpk,"梅","<img src=dpk/3.GIF><br><font color=#000000>")
wpk=replace(wpk,"方","<img src=dpk/4.GIF><br><font color=#FF0000>")
wpk=replace(wpk,"小王","<img src=dpk/xw.gif><br><br>")
wpk=replace(wpk,"大王","<img src=dpk/dw.gif><br><br>")
showpk="<td class=td1 onClick=javascript:s('" & pk & "') align=center> " & wpk & "</font></td>"
end function

function paixupk(pk)
dim wpk
wpk=replace(pk,"紅","")
wpk=replace(wpk,"黑","")
wpk=replace(wpk,"梅","")
wpk=replace(wpk,"方","")
wpk=replace(wpk,"A","1")
wpk=replace(wpk,"J","11")
wpk=replace(wpk,"Q","12")
wpk=replace(wpk,"K","13")
wpk=replace(wpk,"小王","14")
wpk=replace(wpk,"大王","15")
paixupk=wpk
end function
%>
</body></html>


