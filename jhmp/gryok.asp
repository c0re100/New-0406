<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../images/8.jpg">
<%my=info(0)
money=abs(request.form("money"))
if money<>1000 and  money<>10000 and money<>100000 and money<>1000000 and money<>10000000 then 
	Response.Write "<script Language=Javascript>alert('�A�Q�@����H�I');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �Ȩ� from �Τ� where id="& info(9),conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	coon.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
	response.end
end if
if rs("�Ȩ�")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A�S�����r�A�A�Q�@����H�I');window.close();</script>"
	response.end
end if
conn.execute "update �Τ� set �D�w=�D�w+"&int(money/500)&",�Ȩ�=�Ȩ�-"& money &",�ާ@�ɶ�=now() where id="&info(9)
rs.close
rs.open "select ����,�D�w�[,�D�w from �Τ� where id="&info(9),conn
ddj=(rs("����")*1440+2200)+rs("�D�w�[")
if rs("�D�w")>ddj then
conn.execute "update �Τ� set �D�w=" & ddj & "  where id="&info(9)
end if
Response.Write "<script Language=Javascript>alert('�z�V�A���t��|���m�F"& money &"��A�D�w�W��"&int(money/500)&"�I�I');window.close();</script>"
rs.close
conn.close
set rs=nothing
set conn=nothing
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
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>�i�R�ߩ^�m�j["&info(0)&"]</font>�j�L���F�C���򪺩t��|���t��̮��m�F"& money &"��A�n�I�w��^_^" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
