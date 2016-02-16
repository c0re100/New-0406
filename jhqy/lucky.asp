<%

' --- 各項運程內容設定   軒緣世紀江湖 收集 http://www.2vh.net
luck1 = "<img src=img/01.gif border=0></td></tr><tr><td align=center height=15 bgcolor=#c00000><font color=#FFFFFF>大吉</font></td></tr><tr><td bgcolor=#FFdbdb align=center valign=top height=130>似乎你今天的運氣很旺盛!<br>有什麼想做的事便去嘗試吧,<br>或許會有成果!<br>又或者趕快去買個彩卷吧!"
luck2 = "<img src=img/02.gif border=0></td></tr><tr><td align=center height=15 bgcolor=#c00000><font color=#FFFFFF>中吉</font></td></tr><tr><td bgcolor=#FFdbdb align=center valign=top height=130>恭喜你呀!<br>今天應該事事如意,<br>或許壞天氣也會轉晴呢!<br>不過凡事都是腳踏實地好了,<br>這隻是種小玩意!"
luck3 = "<img src=img/03.gif border=0></td></tr><tr><td align=center height=15 bgcolor=#c00000><font color=#FFFFF>小吉</font></td></tr><tr><td bgcolor=#FFdbdb align=center valign=top height=130>心情也算不錯了吧!<br>雖然有時會遇到失敗,<br>但切勿放棄啊!"
luck4 = "<img src=img/04.gif border=0></td></tr><tr><td align=center height=15 bgcolor#c00000><font color=#FFFFF>小兇</font></td></tr><tr><td bgcolor=#FFdbdb align=center valign=top height=130>也許今天不適宜干大事,<br>睡個覺就好了!"
luck5 = "<img src=img/05.gif border=0></td></tr><tr><td align=center height=15 bgcolor=#c0000><font color=#FFFFFF>中兇</font></td></tr><tr><td bgcolor#FFdbdb align=center valign=top height=130>噢不要在乎這些無聊玩意!<br>事情好與壞是由你決定的,<br>別灰心呀!"


randomize
x = int(rnd * 100)
' 確率 20%
if x < 20 then 
 luck = luck1
' 確率 20%
elseif x < 40 then
 luck = luck2
' 確率 30%
elseif x < 70 then
 luck = luck3
' 確率 20%
elseif x < 90 then
 luck = luck4
' 確率 10%
else luck = luck5
end if
%>
<html>

<head>
<meta http-equiv="Content-Type"
content="text/html; charset=big5">
<title>今日運程</title>
<STYLE TYPE="text/css">
<!--
tr, td,body,th    {font-size: 9pt}
a:link    {font-size: 9pt; text-decoration:none; color:<%=link%> }
a:visited {font-size: 9pt; text-decoration:none; color:<%=vlink%> }
a:active  {font-size: 9pt; text-decoration:none; color:<%=alink%> }
a:hover   {font-size: 9pt; text-decoration:underline; color:<%=alink%> }
input,select,textarea     {font-size: 9pt; border: 1 solid black}
-->
</STYLE>
</head>

<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
<div align="left">

<table border="1" cellspacing="1" width="157"
bordercolor="#FFADAD">
    <tr>
        <td height=142 align=center valign=middle><%=luck%>
        </td>
    </tr>
</table>
</div>
</body>
</html>
