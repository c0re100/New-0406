<%'幫派大戰
function bangpai(to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('結盟對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if

if info(2)<5 then
	Response.Write "<script language=JavaScript>{alert('幫派大戰只有掌門才可以操作！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select * FROM 用戶 WHERE 姓名='"&toname&"'" ,conn
if rs("門派")="官府" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('官府不可操作！');}</script>"
	Response.End
end if
menpai=rs("門派")
'set rs=conn.execute ("Select count(*) from 用戶 where 門派='"&info(5)&"'")
'renshu=rs(0)-1

rs.close
rs.open "select count(*) FROM 用戶 WHERE 門派='"&menpai&"' or 門派='"&info(5)&"'",conn
renshu=rs(0)-1
if renshu<100 then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('雙方門下弟子太少了，至少得有100個幫眾才有開戰資格！');}</script>"
	Response.End
end if
rs.close
rs.open "select * FROM 幫派大戰 WHERE 主戰幫派='"&menpai&"'" ,conn
if not rs.bof or not rs.eof then

	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你已經申請幫派戰了！');}</script>"
	Response.End
end if

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Application("ljjh_bpz")=toname
bangpai="["&info(5)&"]掌門["&info(0)&"]向["&menpai&"]掌門{"&toname&"}提出申請幫派大戰：<input type=button value='同意' onClick=javascript:;window.open('bpzok.asp?id="&info(9)&"&yn=1','d');this.disabled=1 style=background-color:#8AAFF1;color:FFFFFF;border: 1 double onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF'>"
end function
%>
