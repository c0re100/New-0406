<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
wg=request("wg")
a=request("a")
wg1=request.form("wg1")
nei=request.form("nl")
if nei>=100000000 then nei=99999999
name=request.form("name")
pass=request.form("pass")
if instr(repass,"'")>0 or instr(pass,",")>0 or instr(nei,",")>0 or instr(wg,",")>0 or instr(wg1,",")>0 or instr(name,",")>0 then
	response.write "�A�n�r�I�«ȥ��͡A�o�^���F�F�a�H�I"
	response.end
end if
if instr(wg,"or")<>0 or instr(pai,"or")<>0 or instr(wg1,"or")<>0 or instr(nl,"or")<>0 or instr(name,"or")<>0 or instr(pass,"or")<>0 then
	response.write "�A�n�r�I�«ȥ��͡A�o�^���F�F�a�H�I"
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
rs.open "SELECT ID FROM �Τ� where id="&info(9) &" and �K�X='" & pass & "'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "�A���Τ�K�X���~��"
	response.end
end if
rs.close
rs.open "SELECT �G�B FROM �Τ� where id="&info(9) &" and �G�B<>'�L'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "�A�S���G�B�A����ЫئX��ޡC"
	response.end
end if
ff=rs("�G�B")
rs.close
rs.open "update �Τ� set �Ȩ�=�Ȩ�-10000 where id="&info(9),conn
if a="m" then
	if info(1)="�k" then
		ff1=info(0)
	else
		ff1=ff
		ff=info(0)
	end if
	conn.Execute "update �X��� set �X���='" & wg1 & "', ���O=" & nei & " where �m�W�k='" & ff1 & "' or �m�W�k='" & ff & "'"
end if
if a="n" then
	if info(1)="�k" then
		ff1=info(0)
	else
		ff1=ff
		ff=info(0)
	end if
	conn.Execute "insert into �X���(�X���,�m�W�k,�m�W�k,���O) values ('" & wg1 & "','" & ff1 & "','" & ff & "','" & nei & "')"
end if
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "stunt.asp"
%>
 