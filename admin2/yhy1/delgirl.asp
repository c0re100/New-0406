<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%><!--#include file="jiu.asp"--><%
sql="select * from �Өk where �m�W='" &info(0)& "'"
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
set connt=nothing
set rs=nothing
Response.Write "<script Language=Javascript>alert('�A�˿��F�a�A���c�S���o�˪��Өk�r�I');window.close();</script>"
response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select �y�O,�Ȩ� from �Τ� where id=" & info(9)
set rs=conn.execute(sql)
meili=rs("�y�O")
maiyin=meili*200
if rs("�Ȩ�")<maiyin then
conn.close
set conn=nothing
set rs=nothing
Response.Write "<script Language=Javascript>alert('�A���y�]�̨S��"&maiyin&"�U�A�Q��ū���A�����S�I');window.close();</script>"
response.end
else
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
sd(196)="FFCFCF"
sd(197)="FFCFCF"
sd(198)="��"
sd(199)="<font color=FFCFCF>�i�ٿ�ū���j</font><font color=RED>"&info(0)&"</font>��F<font color=FFCFCF>"&maiyin&"��</font>�Ȥl�ש󮳦^�F�ۤv���樭���I���o�M�զۥѨ��I"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
sql="update �Τ� set �Ȩ�=�Ȩ�- "&maiyin&"   where id="&info(9)
set rs=conn.execute(sql)
connt.execute("delete * from �Өk where �m�W='"&info(0)& "'")
connt.close
conn.close
set conn=nothing
set rs=nothing
Response.Write "<script Language=Javascript>alert('��F"&maiyin&"�⮳�F�樭���A�w��A�b�S�����ɭԦb�ӧڭ̳o�u�@�I');window.close();</script>"
response.end
end if
%>
