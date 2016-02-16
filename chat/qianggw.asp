<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
id=int(trim(request("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
if info(9)<>id then
	Response.Write "<script Language=Javascript>alert('矗ボ硂ぃ琌"&info(9)&""&id&"ゴぃ具');</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from 珇 where 珇='瑈琍瞈' and 局Τ=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into 珇(珇,局Τ,摸,ю阑,绊㏕,ず,砰,ň眘,计秖,蝗ㄢ,弧,单,穦) values ('瑈琍瞈',"&info(9)&",'珇',0,1000,0,0,0,1,10000000,lxl,0,"&aaao&")"			
else 
id=rs("id")
conn.execute "update 珇 set 计秖=计秖+1,穦="&aaao&" where id="& id &" and 局Τ="&info(9)
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('眔瑈琍瞈');</script>"
%> 