<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
elseif pass="" then
Response.Redirect "../error.asp?id=432"
elseif name1=name2 then
Response.Redirect "../error.asp?id=434"
else

temppass=StrReverse(left(pass&"godxtll,./",10))
templen=len(pass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
pass=replace(mmpassword,"'","B")

%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
ljjh_mdb= split(Application("ljjh_mdb"),"|")
conn.open ljjh_mdb(0)
'����Τ� �y�O�j��100�A���j��1000
sql="SELECT �t��,�ʧO,�Ȩ� FROM �Τ� WHERE �m�W='" & name1 & "' and �K�X='" & pass & "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
conn.close
Response.Redirect "../error.asp?id=431"
response.end
else
if rs("�t��")<>"�L" then
conn.close
Response.Redirect "../error.asp?id=430"
response.end
end if
xb1=trim(rs("�ʧO"))
if rs("�Ȩ�")>=1000 then
rs.close
set rs=nothing
set rs=conn.execute("SELECT �ʧO FROM �Τ� WHERE trim(�m�W)='" & name2 & "'")
if not (rs.bof or rs.eof) then
xb2=trim(rs("�ʧO"))
if xb1<>xb2 then
sql="update �Τ� set �Ȩ�=�Ȩ�-1000 where �m�W='" & name1& "'"
conn.execute sql
sz = "'" & name1 & "','" & name2 & "','" & mess & "', now()"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
ljjh_mdb= split(Application("ljjh_mdb"),"|")
conn.open ljjh_mdb(2)
into_db = "INSERT INTO ��� (�m�W1, �m�W2, ����, �ɶ�) VALUES(" & sz & ")"
conn.Execute(into_db)
sql="delete * from ��� where now()-�ɶ�>10"
conn.execute sql
Response.Redirect "../ok.asp?id=600"
else
Response.Redirect "../error.asp?id=435" %>
<script language=vbscript>
MsgBox "�}���򪱯��A�A�̭ǩʧO�@�˰ڡI����̬O�����\�P���ʪ��C"
location.href = "../jh.asp"
</script>

<%                                        end if
else %>
<script language=vbscript>
MsgBox "���򤤨S���A�n�����H�ڡH�d���F�a�I"
location.href = "../jh.asp"
</script>

<%               end if
else %>
<script language=vbscript>
MsgBox "�A���̦�1000��ڡA�S���ٷQ���@���H���ڥh�a�I"
location.href = "../jh.asp"
</script>

<%		end if
end if
conn.close
end if
%>
<script language=vbscript>
MsgBox "<%=message%>"
location.href = "../jh.asp"
</script>

 