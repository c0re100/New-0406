<html>
<head>
<title>�ڳy�L��</title>
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
  �ڳy�L��<br>
  <br>
  </font><font color="#00FF00">�]�ѤU���@�L���L��,�¤O�L��^</font></b></font> <br>
  <br>
</div>
<table width="100%" border="1" cellpadding="4" cellspacing="0" align="center">
  <tr bgcolor="#000099"> 
    <td width="16%"> 
      <div align="center"><font color="#FFCC66"><b>�W��</b></font></div>
    </td>
    <td width="67%"> 
      <div align="center"><font color="#FFCC66"><b>�һݪ��~</b></font></div>
    </td>
    <td width="17%"> 
      <div align="center"><font color="#FFCC66"><b>����</b></font></div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl1.asp"><font color="#FFFFFF"><img src="001/wq1.gif" width="36" height="79" border="0">���C</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">���Y�G10 �q�ۡG10 �K���G10 ���ۡG10 �����G10 �X�����G10 �B��:10�� �]��T�� 10�U 
        ������)</div>
    </td>
    <td width="17%"> 
      <div align="center">5�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl2.asp"><font color="#FFFFFF"><img src="001/wq2.gif" width="35" height="64" border="0"></font></a><a href="xl2.asp"><font color="#FFFFFF">�����C</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">���_�ۡG10 �q�ۡG20 �K���G20 ���ۡG20 �����G20 �X�����G20 �B��:20�� �]��T�� 30�U 
        ���s��)</div>
    </td>
    <td width="17%"> 
      <div align="center">15�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl3.asp"><font color="#FFFFFF"><img src="001/wq3.gif" width="41" height="81" border="0"></font></a><a href="xl3.asp"><font color="#FFFFFF">�Q�ڼC</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">���_�ۡG20 �q�ۡG40 �K���G40 ���ۡG40 �����G40 �X�����G40 �B��:30�� �]��T�� 50�U 
        ������)</div>
    </td>
    <td width="17%"> 
      <div align="center">25�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl4.asp"><font color="#FFFFFF"><img src="001/wq4.gif" width="39" height="85" border="0"></font></a><a href="xl4.asp"><font color="#FFFFFF">���Z�C</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">���_�ۡG20 �q�ۡG60 �K���G60 ���ۡG60 �����G60 �X�����G60 �B��:40�� �]��T�� 80�U 
        ���s��)</div>
    </td>
    <td width="17%"> 
      <div align="center">35�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
  <tr> 
    <td width="16%" height="16"> 
      <div align="center"><a href="xl5.asp"><font color="#FFFFFF"><img src="001/wq5.gif" width="33" height="87" border="0">�ȭ߼C</font></a></div>
    </td>
    <td width="67%" height="16"> 
      <div align="center">���Y�G10 �q�ۡG100 �K���G100 ���ۡG100 �����G100 �X�����G100 �]��T�� 100�U 
        ������)</div>
    </td>
    <td width="17%" height="16"> 
      <div align="center">45�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl6.asp"><font color="#FFFFFF"><img src="001/wq6.gif" width="40" height="77" border="0"></font></a><a href="xl6.asp"><font color="#FFFFFF">�H�P�]�C</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">���Y�G20 �q�ۡG120 �S�ت��ݡG120 �¥ۡG120 �����G120 �X�����G120 �]��T�� 120�U 
        ���s��)</div>
    </td>
    <td width="17%"> 
      <div align="center">55�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl7.asp"><font color="#FFFFFF"><img src="001/wq7.gif" width="29" height="90" border="0"></font></a><a href="xl7.asp"><font color="#FFFFFF">�涳���b</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">���_�ۡG30 �q�ۡG140 �S�ت��ݡG140 �¥ۡG140 �����G140 ���ۡG140 �]��T�� 140�U 
        ������)</div>
    </td>
    <td width="17%"> 
      <div align="center">65�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl8.asp"><font color="#FFFFFF"><img src="001/wq8.gif" width="36" height="92" border="0"></font></a><a href="xl13.asp"><font color="#FFFFFF">�ݦ�C</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">���ۡG160 �q�ۡG160 �S�ت��ݡG160 �¥ۡG160 �����G160 �X�����G160 �]��T�� 160�U 
        ���s��)</div>
    </td>
    <td width="17%"> 
      <div align="center">75�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl9.asp"><font color="#FFFFFF"><img src="001/wq9.gif" width="37" height="90" border="0"></font></a><a href="xl8.asp"><font color="#FFFFFF">�C�Y�]�C</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">���_�ۡG30 �q�ۡG180 �S�ت��ݡG180 �¥ۡG180 �����G180 �K���G180 �]��T�� 180�U 
        ������)</div>
    </td>
    <td width="17%"> 
      <div align="center">85�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl10.asp"><font color="#FFFFFF"><img src="001/wq10.gif" width="40" height="77" border="0"></font></a><a href="xl9.asp"><font color="#FFFFFF">�ĳ��C</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">���Y�G20 �q�ۡG200 �S�ت��ݡG200 �¥ۡG200 �����G200 ���ۡG200 �]��T�� 200�U 
        ���s��)</div>
    </td>
    <td width="17%"> 
      <div align="center">95�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl11.asp"><font color="#FFFFFF"><img src="001/wq11.gif" width="38" height="70" border="0"></font></a><a href="xl10.asp"><font color="#FFFFFF">���s�C</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">�B��:260�� �q�ۡG220 ���_�ۡG60 �¥ۡG220 �����G220 ���ۡG220 �]��T�� 220�U 
        ������)</div>
    </td>
    <td width="17%"> 
      <div align="center">105�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl12.asp"><font color="#FFFFFF"><img src="001/wq12.gif" width="29" height="93" border="0"></font></a><a href="xl11.asp"><font color="#FFFFFF">�C�P�_�C</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">�B���G260 �S�ت��ݡG250 ���_�ۡG60 �¥ۡG250 �����G250 ���ۡG250 �]��T�� 240�U 
        ���s��)</div>
    </td>
    <td width="17%"> 
      <div align="center">125�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl13.asp"><font color="#FFFFFF"><img src="001/wq13.gif" width="32" height="97" border="0"></font></a><a href="xl12.asp"><font color="#FFFFFF">�������M</font></a></div>
    </td>
    <td width="67%"> 
      <div align="center">�����G200 �S�ت��ݡG280 ���_�ۡG120 �¥ۡG280 ��סG230 ���ۡG280 �]��T�� 260�U 
        ������)</div>
    </td>
    <td width="17%"> 
      <div align="center">155�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl14.asp"><font color="#FFFFFF"><img src="001/wq14.gif" width="38" height="91" border="0"></font></a>�񲴼C</div>
    </td>
    <td width="67%"> 
      <div align="center">�����G300 �S�ت��ݡG300 ���_�ۡG40 �¥ۡG300 ���ۡG300 �B��:230�� �]��T�� 280�U 
        ���s��)</div>
    </td>
    <td width="17%"> 
      <div align="center">205�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl15.asp"><font color="#FFFFFF"><img src="001/wq15.gif" width="35" height="85" border="0"></font></a>���ѼC</div>
    </td>
    <td width="67%"> 
      <div align="center">�����G340 �S�ت��ݡG340 ���_�ۡG80 �¥ۡG340 ���ۡG340 �B��:260�� �]��T�� 300�U 
        ������)</div>
    </td>
    <td width="17%"> 
      <div align="center">245�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl16.asp"><font color="#FFFFFF"><img src="001/wq16.gif" width="37" height="86" border="0"></font></a>�ܴL�C</div>
    </td>
    <td width="67%"> 
      <div align="center">�����G380 �S�ت��ݡG380 ���ݪO�G40 �¥ۡG380 ��סG230 �B��:300�� �]��T�� 320�U 
        ���s��)</div>
    </td>
    <td width="17%"> 
      <div align="center">265�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
  <tr> 
    <td width="16%"> 
      <div align="center"><a href="xl17.asp"><font color="#FFFFFF"><img src="001/wq17.gif" width="39" height="98" border="0"></font></a>�O�s�M</div>
    </td>
    <td width="67%"> 
      <div align="center">���_�ۡG100 �¥ۡG400 ���ݪO�G440 ���_�ۡG200 ���_�ۡG300 �B��:400�� �]��T�� 
        340�U ������)</div>
    </td>
    <td width="17%"> 
      <div align="center">300�ťH�W���a�i�H�ڳy</div>
    </td>
  </tr>
</table>
</body>
</html>