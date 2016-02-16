<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(0)<>"江湖總站" then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>物品發送管理</title>
<style>
A:visited{TEXT-DECORATION: none ;color:005FA2}
A:active{TEXT-DECORATION: none;color:005FA2}
A:link{text-decoration: none;color:005FA2}
A:hover { color: #C81C00; POSITION: relative; BORDER-BOTTOM: #808080 1px dotted A:hover ;}
.t{LINE-HEIGHT: 1.4}
BODY{FONT-FAMILY: 新細明體; FONT-SIZE: 9pt;color:009FE0
scrollbar-face-color:#6080B0;scrollbar-shadow-color:#FFFFFF;scrollbar-highlight-color:#FFFFFF;
scrollbar-3dlight-color:#FFFFFF;scrollbar-darkshadow-color:#FFFFFF;scrollbar-track-color:#FFFFFF;
scrollbar-arrow-color:#FFFFFF;
SCROLLBAR-HIGHLIGHT-COLOR: buttonface;SCROLLBAR-SHADOW-COLOR: buttonface;
SCROLLBAR-3DLIGHT-COLOR: buttonhighlight;SCROLLBAR-TRACK-COLOR: #eeeeee;
SCROLLBAR-DARKSHADOW-COLOR: buttonshadow}
TD,DIV,form ,OPTION,P,TD,BR{FONT-FAMILY: 新細明體; FONT-SIZE: 9pt}
INPUT{BORDER-RIGHT: #2F754F 1px dashed; BORDER-TOP: #2F754F 1px dashed; FONT-SIZE: 9pt; BORDER-LEFT: #2F754F 1px dashed; COLOR: #000000; BORDER-BOTTOM: #2F754F 1px dashed; FONT-FAMILY: 新細明體; BACKGROUND-COLOR: #ffffff}
textarea, select{BORDER-RIGHT: #2F754F 1px dashed; BORDER-TOP: #2F754F 1px dashed; FONT-SIZE: 9pt; BORDER-LEFT: #2F754F 1px dashed; COLOR: #000000; BORDER-BOTTOM: #2F754F 1px dashed; FONT-FAMILY: 新細明體; BACKGROUND-COLOR: #ffffff}
</style>
</head>
<table width="340" border="0" bgcolor="#DFFFFF"  style="BORDER-COLLAPSE: collapse ;border: 1px dashed  #000000">
  <form method=POST action=sendwok.asp>
    <tr> 
      <td height="17"><font color="#990000">用戶姓名：</font></td>
      <td><font color="#990000"> 
        <input name=a type="text" size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">物品名：</font></td>
      <td><font color="#990000"> 
        <input name=b type="text" size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">數量：</font></td>
      <td><font color="#990000"> 
        <input name=c type="text" size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">堅固度：</font></td>
      <td><font color="#990000"> 
        <input name=k type="text"  size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">類型：</font></td>
      <td><font color="#990000"> 
        <select name=d>
          <option value="物品" selected>物品 </option>
          <option value="彈藥">彈藥</option>
          <option value="法寶">法寶 </option>
          <option value="法器">法器 </option>
          <option value="左手">左手 </option>
          <option value="右手">右手 </option>
          <option value="盔甲">盔甲 </option>
          <option value="頭盔">頭盔 </option>
          <option value="盔甲">盔甲 </option>
          <option value="雙腳">雙腳 </option>
          <option value="裝飾">裝飾 </option>
          <option value="禮品">禮品 </option>
          <option value="卡片">卡片 </option>
          <option value="藥品">藥品 </option>
          <option value="鮮花">鮮花 </option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">說明：</font></td>
      <td><font color="#990000"> 
        <input name=e type="text" size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">內力：</font></td>
      <td><font color="#990000"> 
        <input name=f type="text"  size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">體力：</font></td>
      <td><font color="#990000"> 
        <input name=g type="text" size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">攻擊：</font></td>
      <td><font color="#990000"> 
        <input name=h type="text" size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">防禦：</font></td>
      <td><font color="#990000"> 
        <input name=i type="text"  size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">銀兩：</font></td>
      <td><font color="#990000"> 
        <input name=j type="text"  size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">圖檔：</font></td>
      <td><font color="#990000"> 
        <input name=l type="text"  size=20>
        sm</font></td>
    <tr> 
      <td><font color="#990000">等級：</font></td>
      <td><font color="#990000"> 
        <input name=m type="text"  size=20>
        </font></td>
    </tr>
        <tr> 
      <td><font color="#990000">會員：</font></td>
      <td><font color="#990000"> 
        <select name=n>
          <option value="1" selected>會員 </option>
          <option value="0">非會員</option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td  colspan="2"><font color="#990000"><div align="center"> </font> <div align="center"> 
          <font color="#990000"> 
          <input type=submit  name="submit" value="確定送出">
          </font></div></td>
    </tr>
  </form>
</table>
<p>下面是範例 </p>
<table width="622" border="0" cellpadding="2" cellspacing="2">
  <tr bgcolor="#000000"> 
    <td> <div align="center"><font color="#FFFFFF">物品名</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">類型</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">說明</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">內力</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">堅固度 </font></div></td>
    <td> <div align="center"><font color="#FFFFFF">體力</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">攻擊</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">防禦</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">等級</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">銀兩</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">圖檔</font></div></td>
  </tr>
  <tr bgcolor="#FFFFCC"> 
    <td height="20"><font color="#990000">流星</font></td>
    <td><font color="#990000">物品</font></td>
    <td>111111</td>
    <td><font color="#990000">0</font></td>
    <td>1000</td>
    <td><font color="#990000">0</font></td>
    <td><font color="#990000">0</font></td>
    <td><font color="#990000">0</font></td>
    <td>0</td>
    <td>10000000</td>
    <td>111111</td>
  </tr>
  <tr bgcolor="#FFFFCC"> 
    <td height="20">龍珠</td>
    <td><font color="#990000">物品</font></td>
    <td>50000</td>
    <td>0</td>
    <td>100</td>
    <td>0</td>
    <td>0</td>
    <td>0</td>
    <td>0</td>
    <td>10000000</td>
    <td>111111</td>
  </tr>
  <tr bgcolor="#FFFFCC"> 
    <td height="20">流星淚</td>
    <td><font color="#990000">物品</font></td>
    <td>lxl</td>
    <td>0</td>
    <td>1000</td>
    <td>0</td>
    <td>0</td>
    <td>0</td>
    <td>0</td>
    <td>10000000</td>
    <td>111111</td>
  </tr>
  <tr bgcolor="#FFFFCC"> 
    <td height="20">子彈</td>
    <td>彈藥</td>
    <td>2012</td>
    <td>0</td>
    <td>100</td>
    <td>0</td>
    <td>0</td>
    <td>0</td>
    <td>0</td>
    <td>1000000</td>
    <td>2012</td>
  </tr>
</table>
</body>
</html>