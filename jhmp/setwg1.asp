<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
pai=lcase(request("pai"))
a=request("a")
wg1=trim(request.form("wg1"))
id=request.form("id")
lx=trim(request.form("lx"))
nl=trim(request.form("nl"))
if lx<>"攻擊" and  lx<>"防御" and lx<>"恢復" then
Response.Write "<script language=JavaScript>{alert('武功類型錯誤，請輸入攻擊、防御、恢復！');location.href = 'setwg.asp';}</script>"
	response.end
end if

if instr(repass,"'")>0 or instr(pass,",")>0 or instr(nl,",")>0 or instr(wg,",")>0 or instr(wg1,",")>0 or instr(name,",")>0 then
	response.write "你好呀！這是禁地，請勿亂闖!!!!"
	response.end
end if
if instr(wg,"or")<>0 or instr(pai,"or")<>0 or instr(wg1,"or")<>0 or instr(nl,"or")<>0 or instr(name,"or")<>0 or instr(pass,"or")<>0 then
	Response.Write "<script language=JavaScript>{alert('你好呀！這是禁地，請勿亂闖!!!!');location.href = 'setwg.asp';}</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 掌門 FROM 門派 where trim(掌門)='" & info(0) & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你好呀！這是禁地，請勿亂闖!!!!');location.href = 'setwg.asp';}</script>"
	response.end
end if

if a="m" then
if Request.Form("submit")="刪除" then
	wgid=int(abs(request("id")))
	conn.Execute "delete * from 武功  where id="&wgid
end if

	wgid=int(abs(request("id")))
	conn.Execute "update 武功 set 武功='" & wg1 & "',類型='" & lx & "', 內力=" & nl & " where id="&wgid
end if
if a="n" then
rs.close
rs.open "SELECT 武功 FROM 武功 where trim(武功)='" & wg1 & "'",conn
if not rs.eof or not rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('該武功招式江湖裡已經有了，請重建其他招式!!!!');location.href = 'setwg.asp';}</script>"
	response.end
end if
	conn.Execute "insert into 武功(武功,類型,級別,門派,內力) values ('" & wg1 & "','" & lx & "',1,'" & pai & "'," & nl & ")"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('操作成功！');location.href = 'setwg.asp';}</script>"
	Response.End
%>
<body background="../linjianww/0064.jpg">
