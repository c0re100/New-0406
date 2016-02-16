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
<title>洗麻將桌</title>
<META http-equiv="Content-Type" content="text/html; charset=big5">
<STYLE>
A{ TEXT-DECORATION: none;color:#ffffff}
.td1 {background-color:#FFFFFF; }
TD {FONT-SIZE: 9pt;cursor:hand; }
.tb2 { font-family: "新細明體", "serif"; font-size:13; line-height:100%;}
.input2 { height:20;font-size:9pt;background-color:FF9900;color:FFFFFF;border: 1 double }
.clock { height:18;font-size:9pt;background-color:green;color:yellow;border: 1 double }
</STYLE>
</head>
<body bgcolor="#0099FF" text="#FFFFFF">
<br>
  <%
namef=Trim(Request.QueryString("name"))
nickname=info(0)
if namef<>"" and namef<> nickname then
  response.write "<script Language='Javascript'>alert('不要看別人的牌!');location.href = 'javascript:history.go(-1)';</script>"
  response.end
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="DMJ.mdb"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'如果你的服務器支持jet4.0，請使用下載的連接方法，提高程序性能
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr

sql="select * from dmj where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)
if rs.eof and rs.bof then
rs.close
conn.close
set rs=nothing
set conn=nothing
response.write "<script Language='Javascript'>alert('您現在沒有參與牌局，請向你的對手使用\""/打麻將 1000\""命令要求開局。');location.href = 'javascript:history.go(-1)';</script>"
response.end
else
myMj=rs("myMj")
mjto=rs("uto")
rs.close
conn.close
end if
myMj2=split(myMj,"|")
myMjnumb=ubound(myMj2)
%>
</div>
<table border="0" align="center" cellpadding="0" cellspacing="2">
<%
myMj2=split(myMj,"|")
myMjnumb=ubound(myMj2)
dim paixuArr(3)

dim pArr_w(9)
dim pArr_s(9)
dim pArr_t(9)

for i= 0 to myMjnumb-1 
Rp=right(myMj2(i),1)
Lp=left(myMj2(i),1)
if Rp="萬" then 
pArr_w(Lp)=pArr_w(Lp) & "|" & myMj2(i)
elseif Rp="索" then 
pArr_s(Lp)=pArr_s(Lp) & "|" & myMj2(i)
elseif Rp="筒" then 
pArr_t(Lp)=pArr_t(Lp) & "|" & myMj2(i)
else 
paixuArr(3)=paixuArr(3) & "|" & myMj2(i)
end if
next

for i=1 to 9
paixuArr(0)=paixuArr(0) & pArr_w(i)
next
for i=1 to 9
paixuArr(1)=paixuArr(1) & pArr_s(i)
next
for i=1 to 9
paixuArr(2)=paixuArr(2) & pArr_t(i)
next

paixu=paixuArr(0) & paixuArr(1) & paixuArr(2) & paixuArr(3)
myMj3=split(mid(paixu,2),"|")

writestr=""
for i= 0 to myMjnumb-1 step 3
writestr=writestr & "<tr>"
writestr=writestr & showMj(myMj3(i))
if i+1<myMjnumb then writestr=writestr & showMj(myMj3(i+1))
if i+2<myMjnumb then writestr=writestr & showMj(myMj3(i+2))
writestr=writestr & "</tr>" 
next
response.write writestr
%>
</table>
<table id="mjhu" name="mjhu" height="30" border="0" align="center" cellpadding="4" cellspacing="2" bordercolor="#993300" borderColorDark='#E1E1E1' class="tb2" style="display:none">
  <tr> 
    <td align="right" ><strong>對</strong> <input name="txt_mjhu" type="text" readonly size="5" class="clock"></td>
  </tr>
  <tr> 
    <td align="left" ><input type="button"  onclick="javascript:loadHU(0)"  name="hupai" value="胡 牌" class="input2 ">
<br><input type="button"  onclick="javascript:loadHU(1)"  name="qidui" value="七 對" class="input2 ">
      <br>
      <input type="button" name="Submit2"  onclick="javascript:UnloadHU()" value="放 棄" class="input2 "></td>
  </tr>
</table>
<br>
<table borderColorDark='#E1E1E1' class="tb2" border="1" align="center" cellspacing="2" bordercolor="#993300" height="30" cellpadding="4">
  <tr> 
    <td onClick=javascript:s('') onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">重 選</font></td>
  </tr>
  <tr> 
    <td onClick=javascript:s('摸牌') onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">摸 牌</font></td>
  </tr>
  <tr> 
    <td onClick=javascript:s('問牌') onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">問 牌</font></td>
  </tr>
  <tr> 
    <td onclick="javascript:document.all('mjhu').style.display='';mhu='wait';" onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">胡 牌</font></td>
  </tr>
  <tr> 
    <td onClick=javascript:s('認輸了') onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">認 輸</font></td>
  </tr>
  <tr> 
    <td onClick=javascript:window.open("dmjhelp.htm") height="17" onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">幫 
      助</font></td>
  </tr>
</table>
<br>
<div align="center"><font size="2">□</font><font size="2"><%=myMjnumb%>張牌□</font><font size="2"><br>
  □請點擊麻將□</font><br>
  <script language="JavaScript">
var mhu="";
parent.m.location.href="about:blank"
function UnloadHU(){document.all("mjhu").style.display='none';document.all('txt_mjhu').value="";mhu="";}
function loadHU(isqi)
{
if((mhu=="wait"||mhu=="")&&isqi==0) 
{alert("請點擊麻將指定胡牌中的一個對子！\n杠牌不能做對子");
}
else
{
if(isqi==1)mhu="";
window.showModalDialog('dmjhui-ok.asp?to1=<%=mjto%>&onlyDei='+mhu,'','dialogWidth:600px;dialogHeight:250px;scroll:0;status:0;edge:raised')
//UnloadHU();
}
}
function c9(list){document.all('txt_mjhu').value=list;mhu=list}
function s(list){
var sa=parent.f2.document.af.sytemp.value;
var sa2="喫牌";
var sa3=false;
var sa4=false;
var saL=sa.charAt(5);
var listL=list.charAt(0)

if(list=="認輸了")
{parent.sw('[大家]');}
else
{
parent.sw('[<%=mjto%>]');
if(list=="摸牌"||list=="問牌")
{sa4=true;}
else{if(mhu!=""){c9(list);return false;}
		}
}

if(saL==list.charAt(0)&&sa.charAt(6)==list.charAt(1)){
sa2="踫牌";sa3=true
if(sa.charAt(2)=="踫"){sa2="杠牌";}
}

if((saL==listL-1||saL==listL*1+1||saL==listL-2||saL==listL*1+2)&&sa.charAt(6)==list.charAt(1)){sa2="喫牌";sa3=true}
if(sa==""||(sa.charAt(1)!="出"&&sa.charAt(1)!="踫"&&sa.charAt(1)!="杠")||!sa3||sa.length>15){sa="/出牌$ "}
else
{
sa=sa+"\+"
sa=sa.replace("出牌",sa2)
sa=sa.replace("踫牌",sa2)
}
sa=sa+list;
if(sa4){sa="/"+list}
parent.f2.document.af.sytemp.value=sa;
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
function showMj(Mj)
showMj="<td align=center><input type=image border=0 src=""mj/" & intMjp(Mj) & ".gif"" onClick=javascript:s('" & Mj & "') ></td>"
end function

function intMjp(cmj)
dim mj2
dim mjL
mj2=cmj
mjL=left(cmj,1)
if isNumeric(mjL) then 
mj2=right(cmj,1) & mjL
mj2=replace(mj2,"索","1")
mj2=replace(mj2,"筒","2")
mj2=replace(mj2,"萬","3")
else
mj2=replace(mj2,"東風","10")
mj2=replace(mj2,"南風","20")
mj2=replace(mj2,"西風","30")
mj2=replace(mj2,"北風","40")
mj2=replace(mj2,"紅中","41")
mj2=replace(mj2,"白板","42")
mj2=replace(mj2,"發財","43")
end if
intMjp=mj2
end function
%>
</div>
</body></html>


