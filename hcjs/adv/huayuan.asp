<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script Language=Javascript>alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<!--#include file="data.asp"-->
<!--#include file="func.asp"-->
<%
sh=request.form("sh")
'sy=request.form("sy")
my=info(0)
if request.form("h")="1" then
sql="select ��O,���A,�ާ@�ɶ� from �Τ� where �m�W='" & info(0) & "'"
set rs=conn.execute(sql)
if rs("��O")<-1000 or rs("���A")="��" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../chat/chaterr.asp?id=001"
	response.end
end if
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<30 then
	s=30-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('ĵ�i�G�е�"& s &"��A�_�I,�A�Q�S�R�r�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("��O")<20 then	
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A�h�ҵ{�פw�W�X�d��A���������A�٬O���}�t�q���W�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
sql="update �Τ� set ��O=��O-20 where �m�W='" & info(0) & "'"
	conn.execute sql
        rs.close
        set rs=nothing
	conn.close
        set conn=nothing
        message=huayuan(my,sh,sy)
end if
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
sd(195)="�j�a"
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="��"
sd(199)="<font color=#ff0000>����</font>�G"&message
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<style>
body{font-size:9pt;color:#000000;}
p{font-size:16;color:#000000;}
</style>
</head>
<body  bgproperties="fixed" vlink="#000000" bgcolor="#FFFFFF">
<center>
<table border=1 bgcolor="#948754" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#C6BD9B>
        <table height="260" align="center" width="201">
          <tr>
            <td height="37"> 
              <div align="center"><font color="#000000"><strong>��o�ƥ�</strong></font></div>
            
          <tr>
  <td height="182" valign="top"><%=message%><Br><Br>
  </td>
  </tr>
          <td align=center height="29">&nbsp;
            <div align="right">
                <input type=button value="�� �^ �t �q" onclick="location.href='index.asp'">
                <input type=button value="�� �} �t �q" onclick="location.href='../../welcome.asp'">
 </div>
          </td></tr>
</table>
</td></tr>
</table>
</center>
</body>
</html>
