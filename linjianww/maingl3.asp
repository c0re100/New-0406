<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="update �Τ� set �|���O=�|���O+500,����=����+500 where �|������=1 "
Set Rs=conn.Execute(sql)
sql="update �Τ� set �|���O=�|���O+1000,����=����+1000 where �|������=2 "
Set Rs=conn.Execute(sql)
sql="update �Τ� set �|���O=�|���O+1500,����=����+1500 where �|������=3 "
Set Rs=conn.Execute(sql)
sql="update �Τ� set �|���O=�|���O+2000,����=����+2000 where �|������=4 "
Set Rs=conn.Execute(sql)
Set rs=nothing
Set conn=nothing
Response.Write "<script Language='Javascript'>"
Response.Write "alert('�{�b�w�N�Ҧ��|���[�F�|�O�I');"
Response.Write "history.go(-1);"
Response.Write "</script>"
Response.End
%>

 