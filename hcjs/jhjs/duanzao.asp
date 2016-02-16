<html>
<head>
<title>煆造兵器</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--a:hover {  color: #0000FF; cursor: hand}-->
</style>
<link rel="stylesheet" href="../../css.css">
</head>
<script>      
window.scrollBy(0, 100)      
window.resizeTo(0,0)      
window.moveTo(0,0)     
//setInterval("move()",30);      
setTimeout("move()", 1);      
var mxm=50      
var mym=25      
var mx=0      
var my=0      
var sv=50      
var status=1      
var szx=0      
var szy=0      
var c=255      
var n=0      
var sm=30      
var cycle=2      
var done=2      
function move()      
	{      
	if (status == 1)      
		{      
		mxm=mxm/1.05      
		mym=mym/1.05      
		mx=mx+mxm      
		my=my-mym      
		mxm=mxm+(400-mx)/100      
		mym=mym-(300-my)/100      
		window.moveTo(mx,my)      
		rmxm=Math.round(mxm/10)      
		rmym=Math.round(mym/10)      
		if (rmxm == 0)      
			{      
			if (rmym == 0)      
				{      
				status=2      
				}      
			}      
		}      
	if (status == 2)      
		{      
		sv=sv/1.1      
		scrratio=1+1/3      
		mx=mx-sv*scrratio/2      
		my=my-sv/2      
		szx=szx+sv*scrratio      
		szy=szy+sv      
		window.moveTo(mx,my)      
		window.resizeTo(szx,szy)      
		if (sv < 0.1)      
			{      
			status=3      
			}      
		}      
	if (status == 3)      
		{          
		c=c-16           
		if (c<0)      
			{status=8}      
		}      
	if (status == 4)      
		{      
		c=c+16          
		if (c > 239)      
			{status=5}      
		}      
	if (status == 5)      
		{      
		c=c-16           
		if (c < 0)      
			{      
			status=6      
			cycle=cycle-1      
			if (cycle > 0)      
				{      
				if (done == 1)      
					{status=7}      
				else      
					{status=4}      
				}      
			}      
		}      
	if (status == 6)      
		{      
		document.title = "Net3721.com"      
		alert("JavaScript Net3721")      
		cycle=2      
		status=4      
		done=1      
		}      
	if (status == 7)      
		{      
		c=c+4            
		if (c > 128)      
			{status=8}      
		}      
	if (status == 8)      
		{               
		status=9      
}      
	var timer=setTimeout("move()",0.3)      
	}      
</script>

<body bgcolor="#330033" text="#FFFFFF" topmargin="0" leftmargin="0">
<div align="center"><font color="#FFCC66"><b><font size="+3" color="#00FF00"><br>
  <br>
  煆造兵器<br>
  <br>
  </font><font color="#00FF00">（天下絕世無雙兵器,威力無比）</font></b></font> <br>
  <br>
</div>
<table width="100%" border="1" cellpadding="4" cellspacing="0" align="center">
  <tr bgcolor="#000099"> 
    <td width="16%"> 
      <div align="center"><font color="#FFCC66"><b>名稱</b></font></div>
    </td>
    <td width="67%"> 
      <div align="center"><font color="#FFCC66"><b>所需物品</b></font></div>
    </td>
    <td width="17%"> 
      <div align="center"><font color="#FFCC66"><b>說明</b></font></div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl1.asp"><font color="#FFFFFF"><img src="001/wq1.gif" width="36" height="79" border="0">赤劍</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">木頭：10 礦石：10 鐵片：10 朱石：10 水晶：10 合金塊：10 冰水:10個 （堅固度 10萬 
        攻擊性)</div>
    </td>
    <td width="17%"> 
      <div align="center">5級以上玩家可以煆造</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl2.asp"><font color="#FFFFFF"><img src="001/wq2.gif" width="35" height="64" border="0"></font></a><a href="xl2.asp"><font color="#FFFFFF">蝴蝶劍</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">紅寶石：10 礦石：20 鐵片：20 朱石：20 水晶：20 合金塊：20 冰水:20個 （堅固度 30萬 
        防御性)</div>
    </td>
    <td width="17%"> 
      <div align="center">15級以上玩家可以煆造</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl3.asp"><font color="#FFFFFF"><img src="001/wq3.gif" width="41" height="81" border="0"></font></a><a href="xl3.asp"><font color="#FFFFFF">松根劍</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">綠寶石：20 礦石：40 鐵片：40 朱石：40 水晶：40 合金塊：40 冰水:30個 （堅固度 50萬 
        攻擊性)</div>
    </td>
    <td width="17%"> 
      <div align="center">25級以上玩家可以煆造</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl4.asp"><font color="#FFFFFF"><img src="001/wq4.gif" width="39" height="85" border="0"></font></a><a href="xl4.asp"><font color="#FFFFFF">刺蠻劍</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">藍寶石：20 礦石：60 鐵片：60 朱石：60 水晶：60 合金塊：60 冰水:40個 （堅固度 80萬 
        防御性)</div>
    </td>
    <td width="17%"> 
      <div align="center">35級以上玩家可以煆造</div>
    </td>
  </tr>
  <tr> 
    <td width="16%" height="16"> 
      <div align="center"><a href="xl5.asp"><font color="#FFFFFF"><img src="001/wq5.gif" width="33" height="87" border="0">玄冥劍</font></a></div>
    </td>
    <td width="67%" height="16"> 
      <div align="center">木頭：10 礦石：100 鐵片：100 朱石：100 水晶：100 合金塊：100 （堅固度 100萬 
        攻擊性)</div>
    </td>
    <td width="17%" height="16"> 
      <div align="center">45級以上玩家可以煆造</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl6.asp"><font color="#FFFFFF"><img src="001/wq6.gif" width="40" height="77" border="0"></font></a><a href="xl6.asp"><font color="#FFFFFF">寒星魔劍</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">木頭：20 礦石：120 特種金屬：120 黑石：120 水晶：120 合金塊：120 （堅固度 120萬 
        防御性)</div>
    </td>
    <td width="17%"> 
      <div align="center">55級以上玩家可以煆造</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl7.asp"><font color="#FFFFFF"><img src="001/wq7.gif" width="29" height="90" border="0"></font></a><a href="xl7.asp"><font color="#FFFFFF">行雲鬼刃</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">藍寶石：30 礦石：140 特種金屬：140 黑石：140 水晶：140 朱石：140 （堅固度 140萬 
        攻擊性)</div>
    </td>
    <td width="17%"> 
      <div align="center">65級以上玩家可以煆造</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl8.asp"><font color="#FFFFFF"><img src="001/wq8.gif" width="36" height="92" border="0"></font></a><a href="xl13.asp"><font color="#FFFFFF">嗜血劍</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">朱石：160 礦石：160 特種金屬：160 黑石：160 水晶：160 合金塊：160 （堅固度 160萬 
        防御性)</div>
    </td>
    <td width="17%"> 
      <div align="center">75級以上玩家可以煆造</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl9.asp"><font color="#FFFFFF"><img src="001/wq9.gif" width="37" height="90" border="0"></font></a><a href="xl8.asp"><font color="#FFFFFF">青頭魔劍</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">藍寶石：30 礦石：180 特種金屬：180 黑石：180 水晶：180 鐵片：180 （堅固度 180萬 
        攻擊性)</div>
    </td>
    <td width="17%"> 
      <div align="center">85級以上玩家可以煆造</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl10.asp"><font color="#FFFFFF"><img src="001/wq10.gif" width="40" height="77" border="0"></font></a><a href="xl9.asp"><font color="#FFFFFF">融雪劍</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">木頭：20 礦石：200 特種金屬：200 黑石：200 水晶：200 朱石：200 （堅固度 200萬 
        防御性)</div>
    </td>
    <td width="17%"> 
      <div align="center">95級以上玩家可以煆造</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl11.asp"><font color="#FFFFFF"><img src="001/wq11.gif" width="38" height="70" border="0"></font></a><a href="xl10.asp"><font color="#FFFFFF">蛟龍劍</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">冰水:260個 礦石：220 藍寶石：60 黑石：220 水晶：220 朱石：220 （堅固度 220萬 
        攻擊性)</div>
    </td>
    <td width="17%"> 
      <div align="center">105級以上玩家可以煆造</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl12.asp"><font color="#FFFFFF"><img src="001/wq12.gif" width="29" height="93" border="0"></font></a><a href="xl11.asp"><font color="#FFFFFF">七星寶劍</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">冰水：260 特種金屬：250 綠寶石：60 黑石：250 水晶：250 朱石：250 （堅固度 240萬 
        防御性)</div>
    </td>
    <td width="17%"> 
      <div align="center">125級以上玩家可以煆造</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl13.asp"><font color="#FFFFFF"><img src="001/wq13.gif" width="32" height="97" border="0"></font></a><a href="xl12.asp"><font color="#FFFFFF">結拜巨刀</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">水晶：200 特種金屬：280 綠寶石：120 黑石：280 虎肉：230 朱石：280 （堅固度 260萬 
        攻擊性)</div>
    </td>
    <td width="17%"> 
      <div align="center">155級以上玩家可以煆造</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl14.asp"><font color="#FFFFFF"><img src="001/wq14.gif" width="38" height="91" border="0"></font></a>鳳眼劍</div>
    </td>
    <td width="67%"> 
      <div align="center">水晶：300 特種金屬：300 藍寶石：40 黑石：300 朱石：300 冰水:230個 （堅固度 280萬 
        防御性)</div>
    </td>
    <td width="17%"> 
      <div align="center">205級以上玩家可以煆造</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl15.asp"><font color="#FFFFFF"><img src="001/wq15.gif" width="35" height="85" border="0"></font></a>赤碧劍</div>
    </td>
    <td width="67%"> 
      <div align="center">水晶：340 特種金屬：340 紅寶石：80 黑石：340 朱石：340 冰水:260個 （堅固度 300萬 
        攻擊性)</div>
    </td>
    <td width="17%"> 
      <div align="center">245級以上玩家可以煆造</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl16.asp"><font color="#FFFFFF"><img src="001/wq16.gif" width="37" height="86" border="0"></font></a>至尊劍</div>
    </td>
    <td width="67%"> 
      <div align="center">水晶：380 特種金屬：380 金屬板：40 黑石：380 虎肉：230 冰水:300個 （堅固度 320萬 
        防御性)</div>
    </td>
    <td width="17%"> 
      <div align="center">265級以上玩家可以煆造</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl17.asp"><font color="#FFFFFF"><img src="001/wq17.gif" width="39" height="98" border="0"></font></a>屠龍刀</div>
    </td>
    <td width="67%"> 
      <div align="center">藍寶石：100 黑石：400 金屬板：440 綠寶石：200 紅寶石：300 冰水:400個 （堅固度 
        340萬 攻擊性)</div>
    </td>
    <td width="17%"> 
      <div align="center">300級以上玩家可以煆造</div>
    </td>
  </tr>
</table>
</body>
</html>