<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

//sql="update �Ǫ��� set �ާ@�ɶ�=now()"
//Set Rs=conn.Execute(sql)
sql="update �Ǫ��� set ���H��=0"
Set Rs=conn.Execute(sql)
Set rs=nothing
Set conn=nothing
Response.Write "<script Language='Javascript'>"
Response.Write "alert('�{�b�w�i��F�ץ��I');"
Response.Write "history.go(-1);"
Response.Write "</script>"
Response.End
%>
