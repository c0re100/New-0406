<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
gf=trim(request.form("gf"))
zm=trim(Request.Form("zm"))
zl=trim(Request.Form("zl"))
tz=trim(Request.Form("tz"))
dz=trim(Request.Form("dz"))
lz=trim(Request.Form("lz"))
zs=trim(Request.Form("zs"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT * FROM ���X where id=1",conn
conn.Execute "update ���X set gfmoney=" & gf & ",zmmoney=" & zm & ",zlmoney=" & zl & ",tzmoney=" & tz & ",dzmoney=" & dz & ",lznum=" & lz & ",zsnum=" & zs & " where id=1"

sql="insert into �޲z�O�� (�m�W,�ɶ�,ip,�O��) values ('"&info(0)&"',now(),'"&info(15)&"','�u��ק�')"
conn.Execute(sql)
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�ާ@���\�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
