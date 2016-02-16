<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"jhmp")=0 then 
Response.Write "<script language=javascript>{alert('對不起，程序拒絕您的操作！！！\n     按確定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
you=trim(request("you"))
if InStr(you,"oR")<>0 or InStr(you,"Or")<>0 or inStr(you,"&")<>0 or inStr(you,"&&")<>0 or InStr(you,"OR")<>0 or InStr(you,"or")<>0 or InStr(you,"=")<>0 or InStr(you,"`")<>0 or InStr(you,"'")<>0 or InStr(you," ")<>0 or InStr(you," ")<>0 or InStr(you,"'")<>0 or InStr(you,chr(34))<>0 or InStr(you,"\")<>0 or InStr(you,",")<>0 or InStr(you,"<")<>0 or InStr(you,">")<>0 then Response.Redirect "../error.asp?id=54"

if info(6)<>"掌門" then Response.Redirect "../error.asp?id=451"
if info(0)=you then
		Response.Write "<script language=JavaScript>{alert('嚴重警告，不要搞亂');location.href = 'javascript:history.back()';}</script>"
		Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

rs.open "Select 身份 from 用戶 where 門派='"&info(5)&"' and id="&info(9),conn
if rs("身份")<>"掌門" then
rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=javascript>alert('你不是本派掌門不要亂闖！！');window.close();</script>"
			response.end
end if
rs.close
rs.open "Select 身份 from 用戶 where 門派='"&info(5)&"' and 姓名='"&you&"'",conn
conn.execute "update 用戶 set 身份='弟子',grade=1 where 姓名='"&you&"'"
if rs.eof or rs.bof then
rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=javascript>alert('此人不是本派的！！');location.href = 'javascript:history.back()';</script>"
			response.end
			else
rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=javascript>alert('該弟子身份正式取消！！');location.href = 'javascript:history.back()';</script>"
			response.end
			end if
%>