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
            <div align="center">[ <font color="#FF3300">���f�Q </font>]</div></td>
        </tr>
        <tr bgcolor="#FF0000"> 
          <td width="174" height="13"> 
            <div align="center"><font color="#FFFFFF">�� 
              �~ �W</font></div></td>
          <td width="156" height="13"> 
            <div align="center"><font color="#FFFFFF"> 
              �� �~ ��</font></div></td>
          <td width="554" height="13"> 
            <div align="center"><font color="#FFFFFF">�� 
              ��</font></div></td>
          <td width="60" height="13"> 
            <div align="center"><font color="#FFFFFF">�� 
              �q</font></div></td>
          <td width="63" height="13"> 
            <div align="center"><font color="#FFFFFF">�� 
              �@</font></div></td>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=1'>
            <td width="174"> <div align="center">�l�u</div></td>
            <td width="156"> <div align="center"><img src="001/2012.gif" border="0" width="100" height="20" ></div></td>
            <td width="554"> <div align="center">1000�U,����1000</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=2'>
            <td width="174"> <div align="center">�Ѫ��</div></td>
            <td width="156"> <div align="center"><img src="001/400.gif" border="0" width="31" height="42" ></div></td>
            <td width="554"> <div align="center">2000�U,����5000</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=3'>
            <td width="174"> <div align="center">�T����</div></td>
            <td width="156"> <div align="center"><img src="001/2001.gif" border="0" ></div></td>
            <td width="554"> <div align="center">3000�U,����1�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=4'>
            <td width="174"> <div align="center">�}���@</div></td>
            <td width="156"> <div align="center"><img src="001/2002.gif" border="0" ></div></td>
            <td width="554"> <div align="center">4000�U,����2�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=5'>
            <td width="174"> <div align="center">��w�l</div></td>
            <td width="156"> <div align="center"><img src="001/2004.gif" border="0" width="22" height="19" ></div></td>
            <td width="554"> <div align="center">5000�U,����3�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=6'>
            <td width="174"> <div align="center">�m�T�O</div></td>
            <td width="156"> <div align="center"><img src="001/2005.gif" border="0" width="20" height="28" ></div></td>
            <td width="554"> <div align="center">6000�U,����4�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=7'>
            <td width="174"> <div align="center">���_��</div></td>
            <td width="156"> <div align="center"><img src="001/2006.gif" border="0" width="42" height="37" ></div></td>
            <td width="554"> <div align="center">7000�U,����5�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=8'>
            <td width="174"> <div align="center">���_��</div></td>
            <td width="156"> <div align="center"><img src="001/2007.gif" border="0" width="43" height="36" ></div></td>
            <td width="554"> <div align="center">8000�U,����6�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=9'>
            <td width="174"> <div align="center">���_��</div></td>
            <td width="156"> <div align="center"><img src="001/2008.gif" border="0" width="43" height="38" ></div></td>
            <td width="554"> <div align="center">9000�U,����7�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=10'>
            <td width="174"> <div align="center">�]�O�p��</div></td>
            <td width="156"> <div align="center"><img src="001/1100.gif" border="0" width="27" height="25" ></div></td>
            <td width="554"> <div align="center">1��,����8�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=11'>
            <td width="174"> <div align="center">�ͤ�J�|</div></td>
            <td width="156"> <div align="center"><img src="001/2009.gif" border="0" width="100" height="78" ></div></td>
            <td width="554"> <div align="center">2��,����9�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=12'>
            <td width="174"> <div align="center">�����M</div></td>
            <td width="156"> <div align="center"><img src="001/2010.gif" border="0" width="60" height="64" ></div></td>
            <td width="554"> <div align="center">3��,����10�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=13'>
            <td width="174"> <div align="center">�����y</div></td>
            <td width="156"> <div align="center"><img src="001/1600.gif" border="0" width="21" height="21" ></div></td>
            <td width="554"> <div align="center">4��,����11�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=14'>
            <td width="174"> <div align="center">�j�T��</div></td>
            <td width="156"> <div align="center"><img src="001/300.gif" border="0" width="55" height="27" ></div></td>
            <td width="554"> <div align="center">5��,����12�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=15'>
            <td width="174"> <div align="center">�j��</div></td>
            <td width="156"> <div align="center"><img src="001/301.gif" border="0" width="78" height="51" ></div></td>
            <td width="554"> <div align="center">6��,����13�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=16'>
            <td width="174"> <div align="center">�p�U��</div></td>
            <td width="156"> <div align="center"><img src="001/302.gif" border="0" width="58" height="30" ></div></td>
            <td width="554"> <div align="center">7��,����14�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=17'>
            <td width="174"> <div align="center">�ƶ����y</div></td>
            <td width="156"> <div align="center"><img src="001/99004.gif" border="0" width="60" height="47" ></div></td>
            <td width="554"> <div align="center">8��,����15�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=18'>
            <td width="174"> <div align="center">�S�ت���</div></td>
            <td width="156"> <div align="center"><img src="001/2003303.gif" border="0" width="45" height="38" ></div></td>
            <td width="554"> <div align="center">9��,����16�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=19'>
            <td width="174"> <div align="center">���믵�Z</div></td>
            <td width="156"> <div align="center"><img src="001/99003.gif" border="0" width="60" height="41" ></div></td>
            <td width="554"> <div align="center">10��,����17�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=20'>
            <td width="174"> <div align="center">�B�v�C��</div></td>
            <td width="156"> <div align="center"><img src="001/99005.gif" border="0" width="60" height="61" ></div></td>
            <td width="554"> <div align="center">11��,����18�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=21'>
            <td width="174"> <div align="center">�p�^���y</div></td>
            <td width="156"> <div align="center"><img src="001/99006.gif" border="0" width="60" height="30" ></div></td>
            <td width="554"> <div align="center">12��,����19�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=22'>
            <td width="174"> <div align="center">�¥�</div></td>
            <td width="156"> <div align="center"><img src="001/2003305.gif" border="0" width="34" height="34" ></div></td>
            <td width="554"> <div align="center">13��,����20�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=23'>
            <td width="174"> <div align="center">���ݪO</div></td>
            <td width="156"> <div align="center"><img src="001/2003304.gif" border="0" width="42" height="36" ></div></td>
            <td width="554"> <div align="center">14��,����21�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=24'>
            <td width="174"> <div align="center">����</div></td>
            <td width="156"> <div align="center"><img src="001/2003306.gif" border="0" width="31" height="39" ></div></td>
            <td width="554"> <div align="center">15��,����22�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=25'>
            <td width="174"> <div align="center">�B��</div></td>
            <td width="156"> <div align="center"><img src="001/151.gif" border="0" width="16" height="32" ></div></td>
            <td width="554"> <div align="center">16��,����23�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=26'>
            <td width="174"> <div align="center">�K��</div></td>
            <td width="156"> <div align="center"><img src="001/2003301.gif" border="0" width="43" height="30" ></div></td>
            <td width="554"> <div align="center">17��,����24�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=27'>
            <td width="174"> <div align="center">�q��</div></td>
            <td width="156"> <div align="center"><img src="001/2014.gif" border="0" width="35" height="30" ></div></td>
            <td width="554"> <div align="center">18��,����25�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
        <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyzh.asp?id=28'>
            <td width="174"> <div align="center">���Y</div></td>
            <td width="156"> <div align="center"><img src="001/2015.gif" border="0" width="31" height="37" ></div></td>
            <td width="554"> <div align="center">19��,����26�U</div></td>
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
                <input type="SUBMIT" name="Submit" value="�i��">
              </div></td>
          </form>
        </tr>
      </table>
    </td></tr> </table><br> 
</body>

</html>
