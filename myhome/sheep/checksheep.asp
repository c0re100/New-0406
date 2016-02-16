<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0

%>
<!--#include file="data2.asp"-->
<%
info=Session("info")
if info(0)="" then
response.redirect"warning.htm"
else
IF Request.ServerVariables("ReQuest_Method") = "POST" THEN
sheepname=request("nick")
if sheepname="" or len(sheepname)="" then

%>
<script language="Vbscript">
msgbox"寵物名字填寫不正確。",0,"提示"
history.back
</script>
<%
else
rs.open"select 時間,說明,體力,攻擊,類型id,防御,經驗,技能,忠誠,行動點,識別 from 寵物狀態 where 名字='"&sheepname&"' and user='"&info(0)&"' and 體力>0",conn,1,1
if rs.bof then
conn.execute "Delete * From 寵物狀態 Where user='" &info(0)& "'"
conn.execute "Delete * From 技能列表 Where 擁有者='" &info(0)& "'"
conn.execute "Delete * From 道具列表 Where 擁有者='" &info(0)& "'"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('您的寵物已經死了或您不是這隻寵物的主人！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
else
if date()-rs("時間")>=1 then
sql="update 寵物狀態 set 行動點=行動點+4,時間=date() where user='"&info(0)&"'"
conn.Execute(sql)
rs.close
rs.open"select 體力,說明,攻擊,類型id,技能,防御,經驗,忠誠,行動點,識別 from 寵物狀態 where 名字='"&sheepname&"' and user='"&info(0)&"' and 體力>0",conn,1,1
end if
power=rs("體力")
at=rs("攻擊")
id=rs("說明")
fy=rs("防御")
jy=rs("經驗")
zc=rs("忠誠")
ap=rs("行動點")
checker=rs("識別")
jn=rs("技能")


%>

<html>

<head>
<title>寵物育成模式</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" background="../../ff.jpg" text="#000000" leftmargin="0" topmargin="0">
<table border="1" width="648" cellspacing="1" cellpadding="0" align="center" bordercolor="#000000">
<tr>
<td width="642" align="center"><font color="#0000FF">育成寵物模式 </font></td>
</tr>
<tr>
<td width="642">
<table border="0" width="100%" cellspacing="1" cellpadding="0"
height="20">
<tr>
<td class="p3" width="8%"><img border="0" src="image/<%=id%>.gif" ></td>
          <td class="p3" width="29%"><font color="#000000">名字：<%=sheepname%> 
            --<%=jn%></font></td>
<td class="p3" colspan="2" width="63%"> <font color="#000000">【<a href="change.asp">更換名字</a>】【<a href="showstunt.asp" target="_blank">查看寵物技能</a>】【<a href="indexsheep.asp" target="_blank">到寵物商店區</a>】【<a href="javascript:this.location.reload()">刷新本頁</a>】</font></td>
</tr>
</table>
<table border="0" width="103%" cellspacing="1" cellpadding="0"
height="192">
<tr>
          <td class="p2" width="25%" height="18"><font color="#000000">體 力：</font></td>
<td class="p2" width="25%" height="18"><font color="#000000"><%=power%>
</font></td>
          <td class="p2" width="25%" height="18"><font color="#000000">行動點：</font></td>
<td class="p2" width="25%" height="18"> <font color="#000000"><%=ap-checker%></font></td>
</tr>
<tr>
          <td class="p3" width="25%" height="18"><font color="#000000">攻 擊：</font></td>
<td class="p3" width="25%" height="18"><font color="#000000"><%=at%></font></td>
          <td class="p3" width="25%" height="18"><font color="#000000">經 驗：</font></td>
<td class="p3" width="25%" height="18"><font color="#000000"><%=jy%></font></td>
</tr>
<tr>
          <td class="p2" width="25%" height="18"><font color="#000000">防 御：</font></td>
<td class="p2" width="25%" height="18"><font color="#000000"><%=fy%></font></td>
          <td class="p2" width="25%" height="18"><font color="#000000">忠 誠：</font></td>
<td class="p2" width="25%" height="18"><font color="#000000"><%=zc%></font></td>
</tr>
<tr>
          <td class="p1" width="100%" colspan="4" height="17"><font color="#000000">寵物育成注意事項：</font></td>
</tr>
<tr>
<td class="p3" width="100%" colspan="4" height="71">
<p>-使用道具是指你在寵物商店購買的增加各種屬性或起<font color="#FF3333">特別功效</font>道具的使用。<br>
              -鍛練可增加寵物的<font color="#FF0000">經驗值、攻擊力或防御力</font>。<br>
              當行動點變為0後，將不能進行育成，每天有<font color="#FF3333">4點行動點</font>累加！<br>
              冒險將會有機會使用寵物的各種屬性大幅提高，但也有機會減各種屬種甚至使用寵物死亡。<br>
              休息能恢復寵物的體力值，每天有<font color="#FF3333">4次行動點數</font>，用完後就不可對寵物進行育成，隻能等下一天。<br>
              寵物可以在聊天室進行<font color="#FF3333">攻擊</font>呵或起特效，這樣培養一個強大的寵物是必不可少的呵。</p>
</td>
</tr>
<tr>
<td class="p2" width="25%" height="38" align="center">
<input onClick="javascript:window.document.location.href='item.asp'" type="button" value="使用道具"
name="B1" style="font-family: 新細明體; font-size: 12px">
<br>
</td>
<td class="p2" width="25%" height="38" align="center">
<input onClick="javascript:window.document.location.href='test.asp?name=<%=request("nick")%>'" type="button" value="鍛練"
name="B13" style="font-family: 新細明體; font-size: 12px">
</td>
<td class="p2" width="25%" height="38" align="center">
<input onClick="javascript:window.document.location.href='adv.asp?name=<%=request("nick")%>'" type="button" value="冒險"
name="B132" style="font-family: 新細明體; font-size: 12px">
</td>
<td class="p2" width="25%" height="38" align="center">
<input onClick="javascript:window.document.location.href='relax.asp?name=<%=request("nick")%>'" type="button" value="休息"
name="B133" style="font-family: 新細明體; font-size: 12px">
</td>
</tr>
</table>
</td>
</tr>
</table>
<%
end if
rs.close
conn.close
%>
</body>
</html>
<%
end if
end if
end if
set rs=nothing
set conn=nothing
%>