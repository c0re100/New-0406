<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nowinroom=session("nowinroom")
if Instr(LCase(Application("ljjh_useronlinename"&nowinroom)),LCase(" "&info(0)&" "))=0 then 
	Session("ljjh_inthechat")="0" 
	Response.Redirect "../chaterr.asp?id=001" 
end if 

to1=LCase(trim(request.querystring("to1")))
nickname=info(0)
chatroomsn=session("nowinroom")

onlyDei2=intMjp(Request.querystring("onlyDei"))
if onlyDei2="" then onlyDei2="-1"
onlyDei=CINT(onlyDei2)

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="DMJ.mdb"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'�p�G�A���A�Ⱦ����jet4.0�A�ШϥΤU�����s����k�A�����{�ǩʯ�
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 

sql="select * from dmj where ufrom='"& nickname & "'"
Set Rs=conn.Execute(sql)

if rs.eof or rs.bof then 
rs.close
conn.close
set conn=nothing
set rs=nothing
  response.write "<script Language='Javascript'>alert('�S���A�o�ӵP���Τw����!');</script>"
  response.end

end if
Mymj=rs("Mymj")
dmjto=rs("uto")
xiazhu=rs("duz")
logtime=rs("logtime")
mjID=rs("mjID")
oldto1=strMjp(intMjp(rs("mjmp")))
rs.close

if to1<>dmjto then
conn.close
set conn=nothing
set rs=nothing
  response.write "<script Language='Javascript'>alert('�A����⤣�O[" & to1 & "]��!');</script>"
  response.end

end if
'------------------------�s�榡------------------------
dim xia_class,xia_fir,xia_sql
xia_fir=left(xiazhu,1)
xiazhu=mid(xiazhu,2)

select case xia_fir
	case "1"
		xia_class="����"
		xia_sql="����"
	case "2"
		xia_class="�Ȩ�"
		xia_sql="�Ȩ�"
	case "3"
		xia_class="�k�O"
		xia_sql="�k�O"

	case else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�D�k�ާ@!');</script>"
  response.end
end select
'------------------------�s�榡------------------------

dim oldqys
dim mjqys
mjqys=true
fnarr=split(Mymj,"|")
oldqys=right(fnarr(0),1)
for i=0 to ubound(fnarr)-1
	if oldqys<>right(fnarr(i),1) then mjqys=false
	fnint=fnint & intMjp(fnarr(i)) & "|"
next


dim ARRmj(43)
for j= 1 to 4
for i=10 to 43
if instr(fnint & "|",i & "|")<>0 then 
fnint=replace(fnint & "|",i & "|","",1,1,1)
ARRmj(i)=ARRmj(i)+1
end if
next
next


for i=10 to 43
if ARRmj(i)=4 then
	ARRmj(i)=0
	mjhui4=mjhui4+4
	fpnow =fpnow & " ���l�G" & strMjp(i) & strMjp(i) & strMjp(i) & strMjp(i) 
	zaid=true
end if
next

for i=10 to 43

if ARRmj(i)=3 then

	if i<>onlyDei then
		ARRmj(i)=0
		mjhui3=mjhui3+3
		fpnow =fpnow & " �@���G" & strMjp(i) & strMjp(i) & strMjp(i) 
	else
		ARRmj(i)=1
		fpnow =fpnow & " �@��G" & strMjp(i) & strMjp(i) 
		mjhui2=mjhui2+2
	end if
end if
next


for i=10 to 43
if ARRmj(i)=2 and (i=onlyDei or onlyDei=-1) then
ARRmj(i)=0
	if i<41 then
		if ARRmj(i+1)=2 and ARRmj(i+2)=2 then
			ARRmj(i+1)=ARRmj(i+1)-2
			ARRmj(i+2)=ARRmj(i+2)-2
			fpnow =fpnow & " �@�ﶶ�G" & strMjp(i) & strMjp(i+1) & strMjp(i+2)
			fpnow =fpnow & " �@�ﶶ�G" & strMjp(i) & strMjp(i+1) & strMjp(i+2) 
			mjhui3=mjhui3+6
			mjhui2_3=mjhui2_3+6
			i=i+2
		else
			mjhui2=mjhui2+2
			fpnow =fpnow & "  �@��G" & strMjp(i) & strMjp(i) 
		end if
	else
		mjhui2=mjhui2+2
		fpnow =fpnow & " �@��G" & strMjp(i) & strMjp(i) 
	end if
end if
next


for i=10 to 43
if ARRmj(i)=2 then
ARRmj(i)=1
	if i<40 then
		if i<38 and ARRmj(i+1) >0 and ARRmj(i+2) >0 and i<>10 and i<> 20 and i<>30 and i+1<>10 and i+1<> 20 and i+1<>30 and i+2<>10 and i+2<> 20 and i+2<>30 then
			ARRmj(i+1)=ARRmj(i+1)-1
			ARRmj(i+2)=ARRmj(i+2)-1
			mjhui1=mjhui1+1
			fpnow =fpnow & " �@���G" & strMjp(i) & strMjp(i+1) & strMjp(i+2) 
			i=i+2
		else
			mjhui0=mjhui0+1
			fpnow =fpnow & " �@�i�G" & strMjp(i) 
		end if
	else
		fpnow =fpnow & " �@�i�G" & strMjp(i) 
		mjhui0=mjhui0+1
	end if
end if
next


for i=10 to 43
if ARRmj(i)=1 then
	if i<40 then
		if i<38 and ARRmj(i+1) mod 3=1 and ARRmj(i+2) mod 3=1 and i<>10 and i<> 20 and i<>30 and i+1<>10 and i+1<> 20 and i+1<>30 and i+2<>10 and i+2<> 20 and i+2<>30 then
			mjhui1=mjhui1+1
			fpnow =fpnow & " �@���G" & strMjp(i) & strMjp(i+1) & strMjp(i+2) 
			i=i+2
		else
			mjhui0=mjhui0+1
			fpnow =fpnow & " �@�i�G" & strMjp(i) 
		end if
	else
		fpnow =fpnow & " �@�i�G" & strMjp(i) 
		mjhui0=mjhui0+1
	end if
end if
next


mjhui2_3=mjhui2_3+mjhui2
'mod by hcs
'mjhui0=0
if mjhui0>0 then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�A���S���d���ڡA�A�o�صP�]�s�J�H�٦�" & mjhui0 & "�i�s�P!');</script>"
  response.end
elseif mjhui2>2 and mjhui2_3<>14 then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�A���S���d���ڡA�A�o�صP�]�s�J�H�p�G�A�S���C��N�̦h���঳�@�ӹ�l!');</script>"
  response.end
elseif mjhui2<1 then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�A���S���d���ڡA�A�o�صP�]�s�J�H�A�@�ӹ�l�]�S����!');</script>"
  response.end
else
'mod by hcs
'mjqys=true
if mjqys then
fpnow="�i�M�@��j" & fpnow
xiazhu=xiazhu * 2
elseif mjhui2_3=14 then
fpnow="�i�C��l�j" & fpnow
xiazhu=xiazhu + (xiazhu/2)
xiazhu=xiazhu * 2
elseif mjhui4>0 then
fpnow="�i���j" & fpnow
xiazhu=xiazhu + (xiazhu/2)
end if

sql="delete * from dmj where ufrom='" & nickname & "' or ufrom='" & dmjto &"'"
Set Rs=conn.Execute(sql)
sql="delete * from mjInfo where id=" & mjID
Set Rs=conn.Execute(sql)
conn.close
	Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

'------------------------�s�榡------------------------
sql="select " & xia_sql & " from �Τ� where �m�W='" & dmjto &"'"
set rs=conn.execute(sql)
yin1=rs(""&xia_sql&"")
if clng(yin1)<clng(xiazhu) then 
xiazhu=yin1
sayxia=",�o�a�먭�W�u���o��h" & xia_class  & "�F"
end if

conn.execute "update �Τ� set " & xia_sql & "=" & xia_sql & "-"&xiazhu&" where �m�W='"& dmjto &"'"
conn.execute "update �Τ� set " & xia_sql & "=" & xia_sql & "+"& xiazhu&" where �m�W='"& nickname &"'"
conn.close

		
duidu=nickname&",�����j���u�P,�J�F�A��" & xia_class & "��" & xia_class & "...<br>" & fpnow & "<br>"&nickname&"�q"&to1&",Ĺ��"& xiazhu & xia_class & sayxia
'------------------------�s�榡------------------------
msg=nickname & "�����j���u�P,�J�F�A��������...<br>" & fpnow & " �J�G"&oldto1&"<br>" & nickname & "�q" & to1 & "Ĺ��"& xiazhu & xia_class & sayxia

'msg=replace(msg,"duju/mj/","mj/")

duidu=duidu & "<script Language='JavaScript'>if(parent.myn=='" & nickname & "'||parent.myn=='" & to1 & "')parent.f4.showname();</script>"


Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="����"
sd(195)=info(0)
sd(196)="FFC01F"
sd(197)="FFC01F"
sd(198)="��"
sd(199)="<font color=FFC01F>�i�b�±N�j</font>:"&msg&"!"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
End if

function convJS(Jss)
Jss=Replace(Jss,"\","\\")
Jss=Replace(Jss,"/","\/")
convJS=Replace(Jss,chr(34),"\"&chr(34))
end function
function strMjp(cmj)
strMjp = "<img src=duju/mj/" & cMj & ".gif>"
'strMjp = "<input type=image border=0 src=""duju/mj/" & cMj & ".gif"" >"
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

<body bgcolor="#0099FF" text="#FFFFFF" >
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="4" bordercolorlight="000000" bordercolordark="FFFFff">
  <tr> 
<td height="115"> 
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="4">
        <tr> 
          <td width="100%" height="29" align="center" style="FONT-SIZE: 9pt;cursor:hand;"><br>
            <%=replace(msg,"duju/mj/","mj/") %></td>
        </tr>
        <tr> 
          <td align="center" valign="top" height="58">
<input type=button value='�O�� ���O' onClick="javascript:window.close()" style="height:20;font-size:9pt;background-color:#003300;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button223"> 
          </td>
        </tr>
      </table>
</td>
</tr>
</table>
