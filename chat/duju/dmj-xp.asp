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
<title>�~�±N��</title>
<META http-equiv="Content-Type" content="text/html; charset=big5">
<STYLE>
A{ TEXT-DECORATION: none;color:#ffffff}
.td1 {background-color:#FFFFFF; }
TD {FONT-SIZE: 9pt;cursor:hand; }
.tb2 { font-family: "�s�ө���", "serif"; font-size:13; line-height:100%;}
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
  response.write "<script Language='Javascript'>alert('���n�ݧO�H���P!');location.href = 'javascript:history.go(-1)';</script>"
  response.end
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="DMJ.mdb"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'�p�G�A���A�Ⱦ����jet4.0�A�ШϥΤU�����s����k�A�����{�ǩʯ�
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr

sql="select * from dmj where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)
if rs.eof and rs.bof then
rs.close
conn.close
set rs=nothing
set conn=nothing
response.write "<script Language='Javascript'>alert('�z�{�b�S���ѻP�P���A�ЦV�A�����ϥ�\""/���±N 1000\""�R�O�n�D�}���C');location.href = 'javascript:history.go(-1)';</script>"
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
if Rp="�U" then 
pArr_w(Lp)=pArr_w(Lp) & "|" & myMj2(i)
elseif Rp="��" then 
pArr_s(Lp)=pArr_s(Lp) & "|" & myMj2(i)
elseif Rp="��" then 
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
    <td align="right" ><strong>��</strong> <input name="txt_mjhu" type="text" readonly size="5" class="clock"></td>
  </tr>
  <tr> 
    <td align="left" ><input type="button"  onclick="javascript:loadHU(0)"  name="hupai" value="�J �P" class="input2 ">
<br><input type="button"  onclick="javascript:loadHU(1)"  name="qidui" value="�C ��" class="input2 ">
      <br>
      <input type="button" name="Submit2"  onclick="javascript:UnloadHU()" value="�� ��" class="input2 "></td>
  </tr>
</table>
<br>
<table borderColorDark='#E1E1E1' class="tb2" border="1" align="center" cellspacing="2" bordercolor="#993300" height="30" cellpadding="4">
  <tr> 
    <td onClick=javascript:s('') onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">�� ��</font></td>
  </tr>
  <tr> 
    <td onClick=javascript:s('�N�P') onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">�N �P</font></td>
  </tr>
  <tr> 
    <td onClick=javascript:s('�ݵP') onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">�� �P</font></td>
  </tr>
  <tr> 
    <td onclick="javascript:document.all('mjhu').style.display='';mhu='wait';" onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">�J �P</font></td>
  </tr>
  <tr> 
    <td onClick=javascript:s('�{��F') onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">�{ ��</font></td>
  </tr>
  <tr> 
    <td onClick=javascript:window.open("dmjhelp.htm") height="17" onMouseUp="mup(this)" onMouseDown="mdown(this,1)" onMouseOver="mover(this)" onMouseOut="mout(this)"><font color="#FFFFFF">�� 
      �U</font></td>
  </tr>
</table>
<br>
<div align="center"><font size="2">��</font><font size="2"><%=myMjnumb%>�i�P��</font><font size="2"><br>
  �����I���±N��</font><br>
  <script language="JavaScript">
var mhu="";
parent.m.location.href="about:blank"
function UnloadHU(){document.all("mjhu").style.display='none';document.all('txt_mjhu').value="";mhu="";}
function loadHU(isqi)
{
if((mhu=="wait"||mhu=="")&&isqi==0) 
{alert("���I���±N���w�J�P�����@�ӹ�l�I\n���P���వ��l");
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
var sa2="��P";
var sa3=false;
var sa4=false;
var saL=sa.charAt(5);
var listL=list.charAt(0)

if(list=="�{��F")
{parent.sw('[�j�a]');}
else
{
parent.sw('[<%=mjto%>]');
if(list=="�N�P"||list=="�ݵP")
{sa4=true;}
else{if(mhu!=""){c9(list);return false;}
		}
}

if(saL==list.charAt(0)&&sa.charAt(6)==list.charAt(1)){
sa2="��P";sa3=true
if(sa.charAt(2)=="��"){sa2="���P";}
}

if((saL==listL-1||saL==listL*1+1||saL==listL-2||saL==listL*1+2)&&sa.charAt(6)==list.charAt(1)){sa2="��P";sa3=true}
if(sa==""||(sa.charAt(1)!="�X"&&sa.charAt(1)!="��"&&sa.charAt(1)!="��")||!sa3||sa.length>15){sa="/�X�P$ "}
else
{
sa=sa+"\+"
sa=sa.replace("�X�P",sa2)
sa=sa.replace("��P",sa2)
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
mj2=replace(mj2,"��","1")
mj2=replace(mj2,"��","2")
mj2=replace(mj2,"�U","3")
else
mj2=replace(mj2,"�F��","10")
mj2=replace(mj2,"�n��","20")
mj2=replace(mj2,"�護","30")
mj2=replace(mj2,"�_��","40")
mj2=replace(mj2,"����","41")
mj2=replace(mj2,"�ժO","42")
mj2=replace(mj2,"�o�]","43")
end if
intMjp=mj2
end function
%>
</div>
</body></html>


