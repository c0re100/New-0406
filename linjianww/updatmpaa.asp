<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
aa=request("aa")
cc=request("cc")
dd=request("dd")
ee=request("ee")
ff=request("ff")
gg=request("gg")
hh=request("hh")
ii=request("ii")
okcc=request("okcc")
if okcc="�s�W�ۦ�" then
sql="insert into ¾�~�ޯ�(�ޯ�,�׽m����,�׽m�k�O,���Ӫk�O,�򥻧���,�]�k,����,�I�k����) values ('"&aa&"','"&cc&"','"&dd&"','"&ee&"','"&ff&"','"&gg&"','"&hh&"','"&ii&"')"
rs=conn.execute(sql)
conn.close
else
sql="Update ¾�~�ޯ� Set �ޯ�='"&aa&"',�׽m����='"&cc&"',�׽m�k�O='"&dd&"',���Ӫk�O='"&ee&"',�򥻧���="&ff&",�]�k='"&gg&"',����='"&hh&"',�I�k����='"&ii&"' Where �ޯ�='" & aa & "'"
conn.execute(sql)
conn.close
end if
set conn=nothing
Response.Redirect "adminaa.asp"
response.end
%>
