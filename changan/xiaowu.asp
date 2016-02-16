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
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');window.close();</script>"
	Response.End
end if
my=info(0)
sex=info(1)
%>
<!--#include file="data1.asp"--><%
sql="select * from 房屋 where 戶主='" & my & "' or 伴侶='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('您還沒有買房子呢！');location.href = 'fangwu.asp';}</script>"
elseif rs("狀態")<>"正常" then
rs.close
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
p{font-size:9pt; color:#ffee00}
td,select,input{font-size:9pt; color:#000000; height:9pt}
textarea{font-size:9pt; color:#000000}
A:link {COLOR: #ffffff; FONT-SIZE: 9pt;FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:visited {COLOR: #ffffff;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:active {FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:hover {COLOR: #ffff00; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
</style>
</head>
<body bgcolor=#990099>
<center>
<p align="center"><b><font style="font-size: 9pt">歡迎<%=my%>回到自己的小窩</font></b><br><br>
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
<td height="1" width="401" bgcolor="#C4DEFF" rowspan="10"><%if lx="一般房屋" then%><img src="image/kt-01.jpg"><%end if%><%if lx="高級公寓" then%><img src="image/kt-02.jpg"><%end if%><%if lx="花園洋房" then%><img src="image/kt-03.jpg"><%end if%><%if lx="豪華別墅" then%><img src="image/kt-04.jpg"><%end if%> </td>                 

</tr>
<tr>
  <td height="27" width="159" bgcolor="#C4DEFF" > 
    <p align="center"><a href="moshu.asp"><font color="#000000">學習魔法</font></a></p> 
 </td>                 
</tr>
<tr>
<td height="20" width="159" bgcolor="#C4DEFF">
<p align="center"><a href="shangwang.asp"><font color="#000000">上會網吧</font></a></p> 
 </td>                 
</tr>
<tr>
<td height="23" width="159" bgcolor="#C4DEFF">
<p align="center"><a href="kadianshi.asp"><font color="#000000">看電視吧</font></a></p>  </td>                 
</tr>
<tr>
<td height="22" width="159" bgcolor="#C4DEFF"> 
<p align="center"><a href="hekafei.asp"><font color="#000000">來杯咖啡</font></a></p>                 
</tr>
<tr>
              
<td height="164" width="403" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">當前位置：客廳</font></p>
</td>                 
                
</tr>
</table>
<br><br><br><br><br><!--#include file="copy.inc"-->
</body></html>
<%end if
set rs=nothing
set conn=nothing
%>