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
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="update �Τ� set grade=1 where grade<=5 and ����<>'�x��'"
Set Rs=conn.Execute(sql)
sql="update �Τ� set grade=5 where grade<=6 and ����='�x��'"
Set Rs=conn.Execute(sql)
Set rs=nothing
Set conn=nothing
Response.Write "<script Language='Javascript'>"
Response.Write "alert('���q�Τ�޲z��1�šA�x���޲z��5��(�޲z������)�ާ@����!');"
Response.Write "history.go(-1);"
Response.Write "</script>"
Response.End
%>


<html><script language="JavaScript">                                                                  </script></html>

<html><script language="JavaScript">                                                                  </script></html>


<html></html>
<html></html>
<html></html>
<html></html>
<html></html>
<html></html>


 