<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name1=Request.form("name12")
name1=trim(name1)
name1=server.HTMLEncode(name1)
name2=Request.form("name22")
name2=trim(name2)
name2=server.HTMLEncode(name2)
pass=request.form("pass2")
pass=trim(pass)
mess=Request.form("mess2")
mess=trim(mess)
mess=server.HTMLEncode(mess)
if name1="" or name2="" then
	Response.Redirect "../error.asp?id=433"
end if
if pass="" then	Response.Redirect "../error.asp?id=432"
if name1=name2 then
	Response.Redirect "../error.asp?id=434"
end if
temppass=StrReverse(left(pass&"godxtll,./",10))
templen=len(pass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
pass=replace(mmpassword,"'","B")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
'����Τ� �y�O�j��100�A���j��1000
rs.open "SELECT �t��,����,�ʧO,�Ȩ� FROM �Τ� WHERE id="&info(9)&" and �K�X='" & pass & "'",conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=431"
	response.end
end if
if rs("�t��")<>"�L" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=430"
	response.end
end if
if rs("����")="�x��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('���F���@����޲z�B�����v��,�x���T��B�I');history.go(-1);</script>"
	response.end
end if
xb1=trim(rs("�ʧO"))
if rs("�Ȩ�")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�������Q���@���A�S���I�I�I');history.go(-1);</script>"
	response.end
end if
rs.close
rs.open "SELECT ����,�ʧO FROM �Τ� WHERE trim(�m�W)='" & name2 & "'",conn
if not (rs.bof or rs.eof) then
if rs("����")="�x��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('���F���@����޲z�B�����v��,�x���T��B�I');history.go(-1);</script>"
	response.end
end if
	xb2=trim(rs("�ʧO"))
	if xb1<>xb2 then
		conn.execute "update �Τ� set �Ȩ�=�Ȩ�-1000 where �m�W='" & name1& "'"
		sz = "'" & name1 & "','" & name2 & "','" & mess & "', now()"
		conn.Execute "INSERT INTO ��� (�m�W1, �m�W2, ����, �ɶ�) VALUES(" & sz & ")"
		conn.Execute "delete * from ��� where now()-�ɶ�>5"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Redirect "../ok.asp?id=600"
	else
		Response.Redirect "../error.asp?id=435" 
	end if
else
	Response.Write "<script Language=Javascript>alert('���ܡG�S���A�n�����H�r�I�I');history.go(-1);</script>"
	response.end
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%> 