<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 性別,銀兩 from 用戶 where id="&info(9),conn
sex=rs("性別")
yin=rs("銀兩")
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<!--#include file="data1.asp"--><%
sql="select * from 房屋 where 戶主='" & my & "' or 伴侶='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('您還沒有買房子呢！');location.href = 'fangwu.asp';}</script>"
elseif rs("狀態")<>"正常" then
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('您的房屋是不是有點狀況呀！');location.href = '../jh.asp';}</script>"
else
lx=rs("類型")
%>
<html>
<head>
<title>我 的 小 屋</title>
<style>
p{font-size:9pt; color:#000000}
td,select,input{font-size:9pt; color:#000000; height:14pt}
textarea{font-size:9pt; color:#000000}
A:link {COLOR: #ffffff; FONT-SIZE: 9pt;FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:visited {COLOR: #ffffff;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:active {FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:hover {COLOR: #ffff00; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
</style>
</head>
<body bgcolor=#990099>
<center>
<p align="center"><b><font style="font-size: 9pt" color="#ffee00">歡迎<%=my%>回到自己的小窩</font></b><br><br>
<TABLE width='95%' ALIGN=center CELLSPACING=2 BORDER=2 CELLPADDING=5 BGCOLOR='#90c088'><tr><td>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=100%>
<tr bgcolor="#C4DEFF">
<td align="center" width=10%><a href="xiaowu.asp"><font color="#000000"><%if lx="一般房屋" or lx="高級公寓" or lx="花園洋房" or lx="豪華別墅" then%>客廳</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu1.asp"><font color="#000000"><%if lx="一般房屋" or lx="高級公寓" or lx="花園洋房" or lx="豪華別墅" then%>臥室</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu2.asp"><font color="#000000"><%if lx="一般房屋" or lx="高級公寓" or lx="花園洋房" or lx="豪華別墅" then%>儲藏室</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu3.asp"><font color="#000000"><%if lx="高級公寓" or lx="花園洋房" or lx="豪華別墅" then%>餐廳</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu4.asp"><font color="#000000"><%if lx="高級公寓" or lx="花園洋房" or lx="豪華別墅" then%>衛生間</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu5.asp"><font color="#000000"><%if lx="高級公寓" or lx="花園洋房" or lx="豪華別墅" then%>小金庫</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu6.asp"><font color="#000000"><%if lx="花園洋房" or lx="豪華別墅" then%>花園</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu7.asp"><font color="#000000"><%if lx="豪華別墅" then%>遊泳池</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu8.asp"><font color="#000000"><%if lx="豪華別墅" then%>健身房</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu9.asp"><font color="#000000"><%if lx="花園洋房" or lx="豪華別墅" then%>書房</font></a><%end if%></td>
</tr></table></TABLE>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=678 height="281">
<tr bgcolor=#f7f7f7>
<td height="1" width="101" bgcolor="#FFFFFF" rowspan="10"><%if sex="女" then%><img src="image/logonan.jpg"><%end if%><%if sex="男" then%><img src="image/logonv.jpg"><%end if%></td>                 
  <td height="1" width="380" bgcolor="#C4DEFF" rowspan="10"> 
    <%if lx="豪華別墅" then%>
    <img src="image/35028_2.jpg" width="436" height="279"> 
    <%end if%>
  </td>                 
<td height="13" width="159" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">金庫裡有</font><font color="#ff0000"><%if sex="男" then%><%=rs("銀兩")%><%end if%><%if sex="女" then%><%=rs("伴侶銀兩")%><%end if%></font>兩</p>
</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"><%if sex="男" then%><form method="post" action="putmoney.asp"><%end if%><%if sex="女" then%><form method="post" action="putmoney.asp"><%end if%>
我要存<input type="TEXT" maxlength="10" value="0" name="money" style="width:80;border:1 solid #9a9999; font-size:9pt; background-color:#ffffff" size="10"><b>兩</b></td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"><INPUT TYPE="submit" VALUE="確定"><INPUT TYPE="reset" VALUE="重寫"></td>                 
</tr></form>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"><%if sex="男" then%><form method="post" action="putmoney.asp"><%end if%><%if sex="女" then%><form method="post" action="putmoney.asp"><%end if%>
我要取<input type="TEXT" maxlength="10" value="0" name="money" style="width:80;border:1 solid #9a9999; font-size:9pt; background-color:#ffffff" size="10"><b>兩</b></td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"><INPUT TYPE="submit" VALUE="確定"><INPUT TYPE="reset" VALUE="重寫"></td>                 
</tr></form>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF">現在的世道不好啊錢</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF">還是放在家裡比較安</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF">全些。</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"> </td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"> </td>                 
</tr>
<tr>
</table>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=678 height="1"> 
<tr>
<td height="164" width="102" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">：你回來了</font></p>
</td>                 
<td height="164" width="403" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">當前位置：金庫</font></p>
</td>                 
<td height="164" width="157" bgcolor="#C4DEFF">
  <p align="center"><a href="../jh.asp"><img src="image/Bd_wd.GIF" width="40" height="20" border="0"></a></p>
</td>
<td height="164" width="157" bgcolor="#C4DEFF">
<font color="#ff0000">您身上的銀兩：<%=yin%></font>
</td>                                                   
</tr>
</table>
<br><br><br><br><br><!--#include file="copy.inc"-->
</body></html>
<%end if
set rs=nothing	
	conn.close
	set conn=nothing
%>

