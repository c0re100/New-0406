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
	Response.Write "<script Language=Javascript>alert('���ܡG�o���O"&info(9)&"�A"&id&"��������ߡI');</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from ���~ where ���~�W='�y�P�\' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,�|��) values ('�y�P�\',"&info(9)&",'���~',0,1000,0,0,0,1,10000000,lxl,0,"&aaao&")"			
else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�o��@�Ӭy�P�\�I');</script>"
%> 