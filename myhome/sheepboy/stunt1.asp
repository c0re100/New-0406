<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
wg=request("wg")
a=request("a")
nickname=info(0)
wg1=request.form("wg1")
nei=request.form("nl")
if nei>=100000000 then nei=99999999
name=request.form("name")
pass=request.form("pass")
if instr(repass,"'")>0 or instr(pass,",")>0 or instr(nei,",")>0 or instr(wg,",")>0 or instr(wg1,",")>0 or instr(name,",")>0 then
response.write "�A�n�r�I�o�̬O�T�a�A�ФŶ���!!!!"
response.end
end if
if instr(wg,"or")<>0 or instr(pai,"or")<>0 or instr(wg1,"or")<>0 or instr(nl,"or")<>0 or instr(name,"or")<>0 or instr(pass,"or")<>0 then
%>
<script language="vbscript">
MsgBox "�A�n�r�I�o�̬O�T�a�A�ФŶ���!!!!�I"
window.close()
</script>
<%
response.end
end if
temppass=StrReverse(left(pass&"godxtll,./",10))
templen=len(pass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
pass=replace(mmpassword,"'","B")Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
ff=rs("�t��")
sql="update �Τ� set �Ȩ�=�Ȩ�-10000 where �m�W='" & info(0) & "'"
conn.execute sql
rs.close
if a="m" then
if ljjh_mysex="�k" then
ff1=info(0)
else
ff1=ff
ff=info(0)
end if
%><!--#include file="data2.asp"--><%
rs.open"select user from sheep where sheepname='"&sheepname&"' and user='"&info(0)&"'",conn,1,1

sql="update �Ĥl�� set ���l��='���l�Z����', ���O=1000 where �m�W�k='" & ff1 & "' or �m�W�k='" & ff & "'"
'conn.execute"update �Ĥl�� set ���l��='"&life&"',sheephappy='"&sheephappy&"',feeddate='"&date&"',feedsheepday=feedsheepday+1,workload='1' where sheepname='"&sheepname&"' and user='"&info(0)&"'"
'conn.close
Set Rs=conn.Execute(sql)
end if
if a="n" then
if ljjh_mysex="�k" then
ff1=info(0)
else
ff1=ff
ff=info(0)
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
ljjh_mdb= split(Application("ljjh_mdb"),"|")
conn.open ljjh_mdb(1)
sql="insert into �Ĥl��(���l��,�m�W�k,�m�W�k,���O) values ('" & wg1 & "','" & ff1 & "','" & ff & "','" & nei & "')"
Set Rs=conn.Execute(sql)
end if
conn.close
Response.Redirect "stunt.asp"
%>
