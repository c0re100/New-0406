<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name1=Request.form("name1")
name1=trim(name1)
name1=server.HTMLEncode(name1)
name2=Request.form("name2")
name2=trim(name2)
name2=server.HTMLEncode(name2)
pass=request.form("pass")
pass=trim(pass)
liyou=Request.form("liyou")
liyou=trim(liyou)
liyou=server.HTMLEncode(liyou)
if name1="" or name2="" then
	Response.Write "<script Language=Javascript>alert('���ܡG�m�W���šA����ާ@!');location.href = 'LIFEN.htm';</script>"	
	response.end
end if
if pass="" then
	Response.Write "<script Language=Javascript>alert('���ܡG�m�W���šA����ާ@!');location.href = 'LIFEN.htm';</script>"	
	response.end
end if
if name1=name2 then
	Response.Write "<script Language=Javascript>alert('���ܡG���ۤv!');location.href = 'LIFEN.htm';</script>"	
	response.end
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
rs.open "SELECT �t��,�Ȩ� FROM �Τ� WHERE id=" & info(9) & " and �K�X='" & pass & "'",conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�Τ�W�K�X�����T!');location.href = 'LIFEN.htm';</script>"	
	response.end
end if
xb1=trim(rs("�t��"))
if xb1<>name2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG"& name2 &"�]���O�A���t���I!');location.href = 'LIFEN.htm';</script>"	
	response.end
end if
if rs("�Ȩ�")<101000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�ݭn101000���ץi�H���B���I!');location.href = 'LIFEN.htm';</script>"
	response.end
end if
rs.close
rs.open "SELECT ID FROM �Τ� WHERE trim(�m�W)='" & name2 & "'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG["&name2&"�b����W�䤣��A�Ч�޲z���ѨM�I�I');location.href = 'LIFEN.htm';</script>"	
	response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-1000 where id="&info(9)
rs.close
rs.open "SELECT �m�W1 FROM ���B WHERE �m�W1='" & name1 & "' or �m�W2='" & name2 & "'",conn
If not(Rs.Bof OR Rs.Eof) Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG["&info(0)&"�A�b����W�ӽйL���B�F�I�I');location.href = 'LIFEN.htm';</script>"	
	response.end
end if
sz = "'" & name1 & "','" & name2 & "','" & liyou & "', now()"
conn.Execute "INSERT INTO ���B (�m�W1, �m�W2, �z��, �ɶ�) VALUES(" & sz & ")"
conn.execute "delete * from ��� where now()-�ɶ�>10"
Response.Write "<script Language=Javascript>alert('���ܡG"& info(0) &"�A�����B�n�O���\�I�I');location.href = 'LIFEN.htm';</script>"	
response.end
%>
 