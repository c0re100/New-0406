<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(4)=0 then
	Response.Write "<script language=JavaScript>{alert('�D�|�������}�I');}</script>"
	Response.End
end if
a=request("a")
wg1=request.form("wg1")
wg2=request.form("wg2")
wg3=request.form("wg3")
id=request.form("id")
name=request.form("name")
pass=request.form("pass")
if instr(repass,"'")>0 or instr(pass,",")>0 or instr(wg1,",")>0 or instr(wg2,",")>0 or  instr(wg3,",")>0 or instr(name,",")>0 then
response.write "�A�n�r�I�«ȥ��͡A�o�^���F�F�a�H!!!!"
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
rs.open "SELECT * FROM �Τ� where id="& info(9) &" and �K�X='" & pass & "'",conn
if rs.bof or rs.eof then
rs.close
set rs=nothing
conn.close
set conn=nothing
response.write "�A���K�X����A�O���O�d���F�H�I"
response.end
end if
if a="m" then
wgid=int(abs(request("id")))
conn.Execute "update �ϥΧޯ� set �ޯ�='"&wg1&"',����='"&wg2&"',�I�k����='"&wg3&"' where id="&wgid
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�ק粒���I');}</script>"
Response.Redirect "setaa.asp"
%>
<body background="../ff.jpg">
