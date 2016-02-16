<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
name=info(0)
sql="SELECT 等級 FROM 用戶 WHERE 姓名='" & name & "'"
Set Rs=conn.Execute(sql)
jishu=rs("等級")
sql="select 美味度 from 巧面 where 擁有者='" & name & "' and 時限=false"
set rs=conn.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>MsgBox "你沒有面啊，是不是想騙我，快去面館作面吧。"
location.href = "qmg.htm"
</script><%
conn.close
response.end
else
mw=rs("美味度")
tl=mw*8
conn.execute("update 巧面 set 時限=true where 擁有者='"&name&"' and 時限=false")
conn.execute("update 用戶 set 體力=體力+"&tl&" where 姓名='"&name&"'")
Response.write"你把你親手做的面喫了，體力增加" & tl & "點。"
set rs=nothing
conn.close
set conn=nothing
end if	
%> 