<%'結盟
function jiemen(to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('結盟對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if

if info(2)<5 then
	Response.Write "<script language=JavaScript>{alert('幫派結盟只有掌門才可以操作！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 門派,同盟 FROM 用戶 WHERE 姓名='"&toname&"'" ,conn
if rs("門派")="官府" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('官府不可能與你結盟，請重新操作！');}</script>"
	Response.End
end if
menpai=rs("門派")
tongmen=rs("同盟")
if  Instr(1,tongmen,"|"& info(5) & "|")<>0  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('弄錯了吧，對方已經是你的同盟幫派了！');}</script>"
	Response.End
end if

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Application("ljjh_tongmen")=toname
jiemen="["&info(5)&"]掌門["&info(0)&"]向["&menpai&"]掌門{"&toname&"}提出要求結盟：<input type=button value='同意' onClick=javascript:;window.open('jiemenok.asp?id="&info(9)&"&yn=1','d');this.disabled=1 style=background-color:#8AAFF1;color:FFFFFF;border: 1 double onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF'>"
end function
%>
