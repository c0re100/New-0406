<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
name=info(0)
sql="select 美味度 from 原料 where 類型='輔料' and 擁有者='" & name & "' and 時限=false"
set rs=conn.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>MsgBox "你沒有9腸，怎麼做面啊，快到菜市場去買吧。"
location.href = "caichang.asp"
</script><%conn.close
response.end
else
randomize timer
mw=2*rnd*rs("美味度")
sql="select * from 巧面 where  擁有者='" & name & "' and 時限=false"
set rs=conn.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>MsgBox "你還沒有做第一步啊。"
location.href = "m1.htm"
</script><%else
sql="update 巧面 set 美味度=美味度+" &mw& " where 擁有者='" & name & "' and 時限=false"
rs=conn.execute(sql)
sql="update 原料 set 時限=true where 類型='輔料' and 擁有者='" & name & "' and 時限=false"
set rs=conn.execute(sql)
%>
<script language=vbscript>MsgBox "你把9腸放進鍋裡，一碗好面就快在眼前了。"
location.href = "m3.htm"
</script><%
set rs=nothing
conn.close
set conn=nothing
end if	
end if	
%> 