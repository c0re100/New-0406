<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(9)
username=info(0)
myid=info(0)
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
%>
<meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5">
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();}</script>
<style>
td{font-size:9pt;}
</style>
<link href="../css/css4.css" rel="stylesheet" type="text/css">
<body oncontextmenu=self.event.returnValue=false background="../bk.jpg" bgproperties="fixed" topmargin="3">
<div align=center><img src="../232.gif" width="108" height="96"> <br>
  <table width="95%" border="1" cellpadding="2" cellspacing="1" bordercolor="#FFFFFF">
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select * from 用戶 where id="&info(9),conn
if rs.eof or rs.bof then
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script language=vbscript>
MsgBox "你不是江湖中人。"
</script>
<%
conn.close
else
'and "&proviso
sql="select 技能,等級 from 使用技能 where 姓名='"&username&"' order by 等級 desc"
Set Rs=conn.Execute(sql)
msgarr=";"
do while not (rs.EOF or rs.BOF)
attackname=rs("技能")
doasz=rs("等級")
msgarr=msgarr&attackname&";"
%>
<tr bgcolor=ffffff onMouseOver=this.style.backgroundColor='#FFcc00' onMouseOut=this.style.backgroundColor=''><td><a href="xiulian2.asp?mg=<%=rs("技能")%>"><%=rs("技能")%></a></td><td><%=rs("等級")%>級</td><td><a href="javascript:s('/打架$ <%=rs("技能")%>')">打架</a></td></tr>
<%
rs.movenext
loop
rs.Close
sql="select 技能 from 職業技能"
Set Rs=conn.Execute(sql)
do while not (rs.EOF or rs.BOF)
attackname=rs("技能")
if instr(msgarr,";"&attackname&";")=0 then msg=msg+"<tr bgcolor=ffffff onMouseOver=this.style.backgroundColor='#FFcc00' onMouseOut=this.style.backgroundColor=''><td><a href='xiulian2.asp?mg="&attackname&"' onmouseover=""window.status='修習新的技能';return true;"" onmouseout=""window.status='';return true;"" title='修習新的技能'><font color=FF0000>"&attackname&"</font></a></td><td>0級</td><td>禁用</td></tr>"
rs.MoveNext
loop
rs.Close
set rs=nothing
conn.Close
set conn=nothing
end if
%>
<font color="#FFFFFF" size="2"><%=msg%> </font></table>
<br>
<table width="125" border="0" cellspacing="0" cellpadding="0">
<tr>
<td bgcolor="#FFFFCC">1.打架殺傷力計算方式<br><br>
打架消耗法力x招式級數x戰鬥等級x會員等級-(對方的防禦力/10)=打架殺傷力</td>
</tr>
<tr>
<td bgcolor="#FFFFCC"><br>2.殺死對才可以得對方的所有銀兩,和法力的10分之1</td>
</tr>
</table>
</div>
</body>
