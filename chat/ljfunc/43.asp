<%
function seejj()
if info(5)="遊俠" or info(5)="未知" or info(5)="無" or info(5)="" then
Response.Write "<script language=JavaScript>{alert('你還沒有門派，不能查看門派基金！');}</script>"
Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 門派基金 FROM 用戶 WHERE id="&info(9),conn
myjj=rs("門派基金")
rs.close
rs.open "select 幫派基金 FROM 門派 WHERE 門派='" & info(5) & "'",conn
bpjj=rs("幫派基金")
seejj=info(0) & "你現的對本門的貢獻：<font color=red>"&myjj&"</font>兩,["&info(5)&"]的基金數為：<font color=red>"&bpjj&"兩！</font>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>