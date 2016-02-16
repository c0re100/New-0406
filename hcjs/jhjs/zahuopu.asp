<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<html>
<head>
<title>shop</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body topmargin="0"
leftmargin="0" bgcolor="#003399">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> <td align="center" valign="top"> 
      <table border="0" cellpadding="2" cellspacing="2" width="100%" bgcolor="#FF9900">
        <tr bgcolor="#FFFF00"> 
          <td height="12" colspan="5"> <div align="center"></div>
            <div align="center">[ <font color="#FF3300">雜貨鋪 </font>]</div></td>
        </tr>
        <tr bgcolor="#FF0000"> 
          <td width="174" height="13"> 
            <div align="center"><font color="#FFFFFF">物 
              品 名</font></div></td>
          <td width="156" height="13"> 
            <div align="center"><font color="#FFFFFF"> 
              藥 品 圖</font></div></td>
          <td width="554" height="13"> 
            <div align="center"><font color="#FFFFFF">售 
              價</font></div></td>
          <td width="60" height="13"> 
            <div align="center"><font color="#FFFFFF">數 
              量</font></div></td>
          <td width="63" height="13"> 
            <div align="center"><font color="#FFFFFF">操 
              作</font></div></td>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=1'>
            <td width="174"> <div align="center">子彈</div></td>
            <td width="156"> <div align="center"><img src="001/2012.gif" border="0" width="100" height="20" ></div></td>
            <td width="554"> <div align="center">1000萬,金幣1000</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=2'>
            <td width="174"> <div align="center">老虎肉</div></td>
            <td width="156"> <div align="center"><img src="001/400.gif" border="0" width="31" height="42" ></div></td>
            <td width="554"> <div align="center">2000萬,金幣5000</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=3'>
            <td width="174"> <div align="center">狼牙棒</div></td>
            <td width="156"> <div align="center"><img src="001/2001.gif" border="0" ></div></td>
            <td width="554"> <div align="center">3000萬,金幣1萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=4'>
            <td width="174"> <div align="center">破天錐</div></td>
            <td width="156"> <div align="center"><img src="001/2002.gif" border="0" ></div></td>
            <td width="554"> <div align="center">4000萬,金幣2萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=5'>
            <td width="174"> <div align="center">血滴子</div></td>
            <td width="156"> <div align="center"><img src="001/2004.gif" border="0" width="22" height="19" ></div></td>
            <td width="554"> <div align="center">5000萬,金幣3萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=6'>
            <td width="174"> <div align="center">搶劫令</div></td>
            <td width="156"> <div align="center"><img src="001/2005.gif" border="0" width="20" height="28" ></div></td>
            <td width="554"> <div align="center">6000萬,金幣4萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=7'>
            <td width="174"> <div align="center">紅寶石</div></td>
            <td width="156"> <div align="center"><img src="001/2006.gif" border="0" width="42" height="37" ></div></td>
            <td width="554"> <div align="center">7000萬,金幣5萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=8'>
            <td width="174"> <div align="center">綠寶石</div></td>
            <td width="156"> <div align="center"><img src="001/2007.gif" border="0" width="43" height="36" ></div></td>
            <td width="554"> <div align="center">8000萬,金幣6萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=9'>
            <td width="174"> <div align="center">藍寶石</div></td>
            <td width="156"> <div align="center"><img src="001/2008.gif" border="0" width="43" height="38" ></div></td>
            <td width="554"> <div align="center">9000萬,金幣7萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=10'>
            <td width="174"> <div align="center">魔力鑽石</div></td>
            <td width="156"> <div align="center"><img src="001/1100.gif" border="0" width="27" height="25" ></div></td>
            <td width="554"> <div align="center">1億,金幣8萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=11'>
            <td width="174"> <div align="center">生日蛋糕</div></td>
            <td width="156"> <div align="center"><img src="001/2009.gif" border="0" width="100" height="78" ></div></td>
            <td width="554"> <div align="center">2億,金幣9萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=12'>
            <td width="174"> <div align="center">絕情刀</div></td>
            <td width="156"> <div align="center"><img src="001/2010.gif" border="0" width="60" height="64" ></div></td>
            <td width="554"> <div align="center">3億,金幣10萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=13'>
            <td width="174"> <div align="center">水晶球</div></td>
            <td width="156"> <div align="center"><img src="001/1600.gif" border="0" width="21" height="21" ></div></td>
            <td width="554"> <div align="center">4億,金幣11萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=14'>
            <td width="174"> <div align="center">大鯊魚</div></td>
            <td width="156"> <div align="center"><img src="001/300.gif" border="0" width="55" height="27" ></div></td>
            <td width="554"> <div align="center">5億,金幣12萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=15'>
            <td width="174"> <div align="center">大草魚</div></td>
            <td width="156"> <div align="center"><img src="001/301.gif" border="0" width="78" height="51" ></div></td>
            <td width="554"> <div align="center">6億,金幣13萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=16'>
            <td width="174"> <div align="center">小鯉魚</div></td>
            <td width="156"> <div align="center"><img src="001/302.gif" border="0" width="58" height="30" ></div></td>
            <td width="554"> <div align="center">7億,金幣14萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=17'>
            <td width="174"> <div align="center">排雲秘籍</div></td>
            <td width="156"> <div align="center"><img src="001/99004.gif" border="0" width="60" height="47" ></div></td>
            <td width="554"> <div align="center">8億,金幣15萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=18'>
            <td width="174"> <div align="center">特種金屬</div></td>
            <td width="156"> <div align="center"><img src="001/2003303.gif" border="0" width="45" height="38" ></div></td>
            <td width="554"> <div align="center">9億,金幣16萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=19'>
            <td width="174"> <div align="center">明月秘訣</div></td>
            <td width="156"> <div align="center"><img src="001/99003.gif" border="0" width="60" height="41" ></div></td>
            <td width="554"> <div align="center">10億,金幣17萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=20'>
            <td width="174"> <div align="center">浮影劍譜</div></td>
            <td width="156"> <div align="center"><img src="001/99005.gif" border="0" width="60" height="61" ></div></td>
            <td width="554"> <div align="center">11億,金幣18萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=21'>
            <td width="174"> <div align="center">雷霆秘籍</div></td>
            <td width="156"> <div align="center"><img src="001/99006.gif" border="0" width="60" height="30" ></div></td>
            <td width="554"> <div align="center">12億,金幣19萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=22'>
            <td width="174"> <div align="center">黑石</div></td>
            <td width="156"> <div align="center"><img src="001/2003305.gif" border="0" width="34" height="34" ></div></td>
            <td width="554"> <div align="center">13億,金幣20萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=23'>
            <td width="174"> <div align="center">金屬板</div></td>
            <td width="156"> <div align="center"><img src="001/2003304.gif" border="0" width="42" height="36" ></div></td>
            <td width="554"> <div align="center">14億,金幣21萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=24'>
            <td width="174"> <div align="center">朱石</div></td>
            <td width="156"> <div align="center"><img src="001/2003306.gif" border="0" width="31" height="39" ></div></td>
            <td width="554"> <div align="center">15億,金幣22萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=25'>
            <td width="174"> <div align="center">冰水</div></td>
            <td width="156"> <div align="center"><img src="001/151.gif" border="0" width="16" height="32" ></div></td>
            <td width="554"> <div align="center">16億,金幣23萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=26'>
            <td width="174"> <div align="center">鐵片</div></td>
            <td width="156"> <div align="center"><img src="001/2003301.gif" border="0" width="43" height="30" ></div></td>
            <td width="554"> <div align="center">17億,金幣24萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=27'>
            <td width="174"> <div align="center">礦石</div></td>
            <td width="156"> <div align="center"><img src="001/2014.gif" border="0" width="35" height="30" ></div></td>
            <td width="554"> <div align="center">18億,金幣25萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=28'>
            <td width="174"> <div align="center">木頭</div></td>
            <td width="156"> <div align="center"><img src="001/2015.gif" border="0" width="31" height="37" ></div></td>
            <td width="554"> <div align="center">19億,金幣26萬</div></td>
            <td width="60"> <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
              </div></td>
            <td width="63"> <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
              </div></td>
          </form>
        </tr>
      </table>
    </td></tr> </table><br> 
</body>

</html>
