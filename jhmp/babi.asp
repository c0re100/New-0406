<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=trim(request.querystring("id"))
if not isnumeric(id) then 
	Response.Write "<script language=JavaScript>{alert('���ܡG["&id&"]��J���~�AID�ШϥμƦr�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �m�W,�O�d2,mvalue from �Τ� where id=" & id & " and ���ФH='"&info(0)&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�o�ӤH�]���O�A���ШӪ��A�A�p�l�Q�@����I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
toname=rs("�m�W")
if rs("�O�d2")="����"&Month(date()) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��,�A�Q�@����r�A�A�w�g���L�֤F�A�ٷQ����?');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
yudian=rs("mvalue")
conn.Execute "update �Τ� set �O�d2='"&"����"&Month(date())&"' where id="&id
conn.Execute "update �Τ� set allvalue=allvalue+"&int(yudian*0.05)&" where id="&info(9)
rs.close
set rs=nothing
conn.close
set conn=nothing
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
sd(199)="<font color=FFD7F4>�p�D�����G["&info(0)&"]�q("&toname&")���W����w�I:"&int(yudian*0.05)&"�I�A�V�O�a�A�h�ԤH�I</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Response.Write "<script Language=Javascript>alert('���ܡG["&info(0)&"]�q("&toname&")���W����w�I:"&int(yudian*0.05)&"�I�A�V�O�a�A�h�ԤH�I');location.href = 'javascript:history.go(-1)';</script>"
%>
