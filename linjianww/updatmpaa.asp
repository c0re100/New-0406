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
if okcc="新增招式" then
sql="insert into 職業技能(技能,修練條件,修練法力,消耗法力,基本攻擊,魔法,圖檔,施法說明) values ('"&aa&"','"&cc&"','"&dd&"','"&ee&"','"&ff&"','"&gg&"','"&hh&"','"&ii&"')"
rs=conn.execute(sql)
conn.close
else
sql="Update 職業技能 Set 技能='"&aa&"',修練條件='"&cc&"',修練法力='"&dd&"',消耗法力='"&ee&"',基本攻擊="&ff&",魔法='"&gg&"',圖檔='"&hh&"',施法說明='"&ii&"' Where 技能='" & aa & "'"
conn.execute(sql)
conn.close
end if
set conn=nothing
Response.Redirect "adminaa.asp"
response.end
%>
