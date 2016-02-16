<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
session("ljjh_jm")=session("ljjh_jm")+1
if session("ljjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
randomize timer
regjm=int(rnd*9998)+1
id=LCase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');window.close();}</script>"
Response.End 
end if
%>
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=big5'>
<SCRIPT LANGUAGE="JavaScript">
function check(got) {
var name=got.name.value;
var e_mail=got.e_mail.value;
if (name=="") {alert("請輸入用戶名 !!");return false;}
if (e_mail=="") {alert("請輸入電郵 !!");return false;}
if (got.psw.value=="") {alert("請輸入密碼 !!");return false;}
if (got.pswc.value=="") {alert("請輸入密碼確認 !!");return false;}
if (got.e_mail.value=="") {alert("請輸入電郵 !!");return false;}
if(name.indexOf("\\")!=-1 || name.indexOf("1")!=-1 || name.indexOf("2")!=-1  || name.indexOf("3")!=-1  || name.indexOf("4")!=-1  || name.indexOf("5")!=-1  || name.indexOf("6")!=-1  || name.indexOf("7")!=-1  || name.indexOf("8")!=-1  || name.indexOf("9")!=-1  || name.indexOf("0")!=-1  || name.indexOf("１")!=-1  || name.indexOf("２")!=-1  || name.indexOf("３")!=-1  || name.indexOf("４")!=-1  || name.indexOf("５")!=-1  || name.indexOf("６")!=-1  || name.indexOf("７")!=-1  || name.indexOf("８")!=-1  || name.indexOf("９")!=-1  || name.indexOf("０")!=-1 || name.indexOf("��")!=-1　||　name.indexOf("#")!=-1 || name.indexOf("&")!=-1 || name.indexOf("\"")!=-1 || name.indexOf("'")!=-1) {alert("姓名中請勿含有特殊符號");return false;}
if(CheckIfEnglish(name)) {window.alert("姓名中不能有非法字串、英文、數字；只能使用繁體中文字！^_^!");return false;}
function CheckIfEnglish( String ) {
var Letters = "撚ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890 １２３４５６７８９０浚撚 \"";
var i;
var c;
for( i = 0; i < String.length; i ++ ) {
c = String.charAt( i );
if (Letters.indexOf( c ) < 0)
return false;}
return true;}
return true;}
</SCRIPT>
<title>用戶注冊</title>
<link rel="stylesheet" href="css.css">
</head>
<body leftmargin="0" topmargin="0" bgcolor="#EEF7FF">
<div align=center>
  <table border="0" cellspacing="0" cellpadding="0" width="568" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
<td background="1.gif" width="568" height="83">　</td>
	</tr>
    <tr> 
      <td background="2.gif">
      <FORM action=joinjhnow.asp method=post name=regform>
      <div align="center">
        <center>
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="480" id="AutoNumber1" height="79" bordercolor="#D9ECFF">
          <tr>
                <TD width="448" height="1" bgcolor="#EEF7FF" align=center>
                歡迎離到<font color="#0000FF">魔神江湖</font></TD>
              </tr>
          <tr>
                <TD width="448" height="1" bgcolor="#EEF7FF" align=center>
                本站新增左好多程式同新卡</TD>
              </tr>
          <tr>
                <TD width="448" height="1" bgcolor="#EEF7FF" align=center>
                <font color="#FF0000">本站一注冊就是4級會員,拉人仲送野~仲唔叫多朋友來玩</font></TD>
              </tr>
          <tr>
            <td height="301"> 
                <p>姓名：<FONT COLOR="#FF0000">*</FONT> 
                  <input type=text name=name size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'"> 
                用戶名5個中文字以內<font color="#FF0000">(建意:不要用符號) 
                  <input type=hidden name=regjm value="3621">
                  </font>
                  <br>
                  密碼：<FONT COLOR="#FF0000">*</FONT> 
                  <input type=password name=psw size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'"> 
                6-10個漢字<br>
                  確認：<FONT COLOR="#FF0000">*</FONT> 
                  <input type=password name=pswc size=12 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'"> 
                同上 <BR>
                  找回：<FONT COLOR="#FF0000">*</FONT> 
                  <INPUT TYPE=password NAME=pswc2 SIZE=12 MAXLENGTH="10" STYLE="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" ONMOUSEOVER="this.style.backgroundColor = '#FFCC00'" ONMOUSEOUT="this.style.backgroundColor = '#88B0D8'">
                  <FONT COLOR="#FF0000">用於密碼遺忘時取回密碼</FONT><br>
                  性別：<FONT COLOR="#FF0000">*</FONT> 
                  <select name=sex size=1>
                    <option value=男 selected>男 
                    <option value=女>女 
                  </select>
                  <BR>
                  生日：<FONT COLOR="#FF0000">*</FONT> 
                  <input type=text name=regyear size=4 maxlength="4" value="1976" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'"> 
                年 
                  <select name=regmonth class=p1 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                    <option value="01" selected>01</option>
                    <option value="02">02</option>
                    <option value="03">03</option>
                    <option value="04">04</option>
                    <option value="05">05</option>
                    <option value="06">06</option>
                    <option value="07">07</option>
                    <option value="08">08</option>
                    <option value="09">09</option>
                    <option value="10">10</option>
                    <option value="11">11</option>
                    <option value="12">12</option>
                  </select> 月 
                  <input name=regday type=text style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="11" size=2 maxlength="2"> 
                日 <br>
                  OICQ： 
                  <input name="oicq" type="text" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="123456" size="9" maxlength="9">&nbsp;&nbsp; 
                填寫QQ號與朋友聯繫5位以上！<br>
                  職業： 
                  <select type=text name=zhiye size=1 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                    <option value="戰士" selected>戰士</option>
                    <option value="勇士">勇士</option>
                    <option value="道士">道士</option>
                 </select>
                  <br>
                  信箱：<FONT COLOR="#FF0000">*</FONT> 
                  <input type=text name=e_mail size=25 style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="dai@asz.cn" maxlength="30">
                  例如: a@l.com <BR>
                  <FONT COLOR="#FF0000">最好填寫真實信箱，沒有可以隨便填</FONT><br>
                  認證：<FONT COLOR="#FF0000">*</FONT> 
                  <input type=text name=regjm1 size=5 maxlength="5" value=3621  style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  防止使用注冊機請輸入認證碼：<font color="#FF0000">3621</font><br>
                  介紹人： 
                  <input size=10 name=jsr value="" maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'">
                  介紹人姓名，沒有可以不填寫! <br>
                  名單頭像：<img id=face src="0-2.gif" width='84' height='82' alt=個人形像代表> 
                  <select name=face size=1 onChange="document.images['face'].src='../ico/'+options[selectedIndex].value+'-2.gif';" style="BACKGROUND-COLOR: #99CCFF; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000">
                    
                    <option value='0'>0</option>
                    
                    <option value='1'>1</option>
                    
                    <option value='2'>2</option>
                    
                    <option value='3'>3</option>
                    
                    <option value='4'>4</option>
                    
                    <option value='5'>5</option>
                    
                    <option value='6'>6</option>
                    
                    <option value='7'>7</option>
                    
                    <option value='8'>8</option>
                    
                    <option value='9'>9</option>
                    
                    <option value='10'>10</option>
                    
                    <option value='11'>11</option>
                    
                    <option value='12'>12</option>
                    
                    <option value='13'>13</option>
                    
                    <option value='14'>14</option>
                    
                    <option value='15'>15</option>
                    
                    <option value='16'>16</option>
                    
                    <option value='17'>17</option>
                    
                    <option value='18'>18</option>
                    
                    <option value='19'>19</option>
                    
                    <option value='20'>20</option>
                    
                    <option value='21'>21</option>
                    
                    <option value='22'>22</option>
                    
                    <option value='23'>23</option>
                    
                    <option value='24'>24</option>
                    
                    <option value='25'>25</option>
                    
                    <option value='26'>26</option>
                    
                    <option value='27'>27</option>
                    
                    <option value='28'>28</option>
                    
                    <option value='29'>29</option>
                    
                    <option value='30'>30</option>
                    
                    <option value='31'>31</option>
                    
                    <option value='32'>32</option>
                    
                    <option value='33'>33</option>
                    
                    <option value='34'>34</option>
                    
                    <option value='35'>35</option>
                    
                    <option value='36'>36</option>
                    
                    <option value='37'>37</option>
                    
                    <option value='38'>38</option>
                    
                    <option value='39'>39</option>
                    
                    <option value='40'>40</option>
                    
                    <option value='41'>41</option>
                    
                    <option value='42'>42</option>
                    
                    <option value='43'>43</option>
                    
                    <option value='44'>44</option>
                    
                    <option value='45'>45</option>
                    
                    <option value='46'>46</option>
                    
                    <option value='47'>47</option>
                    
                    <option value='48'>48</option>
                    
                  </select>
                  <input type=submit value=申請 name="submit" style="border-left:1px solid solid; border-right:1px solid solid; border-top:1px solid #000000; border-bottom:1px solid #000000; font-size: 9pt; ">
                  <input type=button value=關閉 onClick="window.close()" name="button" style="border-left:1px solid solid; border-right:1px solid solid; border-top:1px solid #000000; border-bottom:1px solid #000000; font-size: 9pt; ">
                </p>
              </td>
          </tr>
          </table>
        </center>
      </div>
      </form><!=></td>
    </tr>
    <tr> 
      <td width="500" background="3.gif" height="56">　</td>
    </tr>
  </table>   
</div>   
<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
<NOSCRIPT><IFRAME SRC=*.HTML><IFRAME></noscript><font size="2"><NOSCRIPT>
</body>
</html>