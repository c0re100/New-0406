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
if info(0)<>"�����`��" then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>���~�o�e�޲z</title>
<style>
A:visited{TEXT-DECORATION: none ;color:005FA2}
A:active{TEXT-DECORATION: none;color:005FA2}
A:link{text-decoration: none;color:005FA2}
A:hover { color: #C81C00; POSITION: relative; BORDER-BOTTOM: #808080 1px dotted A:hover ;}
.t{LINE-HEIGHT: 1.4}
BODY{FONT-FAMILY: �s�ө���; FONT-SIZE: 9pt;color:009FE0
scrollbar-face-color:#6080B0;scrollbar-shadow-color:#FFFFFF;scrollbar-highlight-color:#FFFFFF;
scrollbar-3dlight-color:#FFFFFF;scrollbar-darkshadow-color:#FFFFFF;scrollbar-track-color:#FFFFFF;
scrollbar-arrow-color:#FFFFFF;
SCROLLBAR-HIGHLIGHT-COLOR: buttonface;SCROLLBAR-SHADOW-COLOR: buttonface;
SCROLLBAR-3DLIGHT-COLOR: buttonhighlight;SCROLLBAR-TRACK-COLOR: #eeeeee;
SCROLLBAR-DARKSHADOW-COLOR: buttonshadow}
TD,DIV,form ,OPTION,P,TD,BR{FONT-FAMILY: �s�ө���; FONT-SIZE: 9pt}
INPUT{BORDER-RIGHT: #2F754F 1px dashed; BORDER-TOP: #2F754F 1px dashed; FONT-SIZE: 9pt; BORDER-LEFT: #2F754F 1px dashed; COLOR: #000000; BORDER-BOTTOM: #2F754F 1px dashed; FONT-FAMILY: �s�ө���; BACKGROUND-COLOR: #ffffff}
textarea, select{BORDER-RIGHT: #2F754F 1px dashed; BORDER-TOP: #2F754F 1px dashed; FONT-SIZE: 9pt; BORDER-LEFT: #2F754F 1px dashed; COLOR: #000000; BORDER-BOTTOM: #2F754F 1px dashed; FONT-FAMILY: �s�ө���; BACKGROUND-COLOR: #ffffff}
</style>
</head>
<table width="340" border="0" bgcolor="#DFFFFF"  style="BORDER-COLLAPSE: collapse ;border: 1px dashed  #000000">
  <form method=POST action=sendwok.asp>
    <tr> 
      <td height="17"><font color="#990000">�Τ�m�W�G</font></td>
      <td><font color="#990000"> 
        <input name=a type="text" size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">���~�W�G</font></td>
      <td><font color="#990000"> 
        <input name=b type="text" size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">�ƶq�G</font></td>
      <td><font color="#990000"> 
        <input name=c type="text" size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">��T�סG</font></td>
      <td><font color="#990000"> 
        <input name=k type="text"  size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">�����G</font></td>
      <td><font color="#990000"> 
        <select name=d>
          <option value="���~" selected>���~ </option>
          <option value="�u��">�u��</option>
          <option value="�k�_">�k�_ </option>
          <option value="�k��">�k�� </option>
          <option value="����">���� </option>
          <option value="�k��">�k�� </option>
          <option value="����">���� </option>
          <option value="�Y��">�Y�� </option>
          <option value="����">���� </option>
          <option value="���}">���} </option>
          <option value="�˹�">�˹� </option>
          <option value="§�~">§�~ </option>
          <option value="�d��">�d�� </option>
          <option value="�ī~">�ī~ </option>
          <option value="�A��">�A�� </option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">�����G</font></td>
      <td><font color="#990000"> 
        <input name=e type="text" size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">���O�G</font></td>
      <td><font color="#990000"> 
        <input name=f type="text"  size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">��O�G</font></td>
      <td><font color="#990000"> 
        <input name=g type="text" size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">�����G</font></td>
      <td><font color="#990000"> 
        <input name=h type="text" size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">���m�G</font></td>
      <td><font color="#990000"> 
        <input name=i type="text"  size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">�Ȩ�G</font></td>
      <td><font color="#990000"> 
        <input name=j type="text"  size=20>
        </font></td>
    </tr>
    <tr> 
      <td><font color="#990000">���ɡG</font></td>
      <td><font color="#990000"> 
        <input name=l type="text"  size=20>
        sm</font></td>
    <tr> 
      <td><font color="#990000">���šG</font></td>
      <td><font color="#990000"> 
        <input name=m type="text"  size=20>
        </font></td>
    </tr>
        <tr> 
      <td><font color="#990000">�|���G</font></td>
      <td><font color="#990000"> 
        <select name=n>
          <option value="1" selected>�|�� </option>
          <option value="0">�D�|��</option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td  colspan="2"><font color="#990000"><div align="center"> </font> <div align="center"> 
          <font color="#990000"> 
          <input type=submit  name="submit" value="�T�w�e�X">
          </font></div></td>
    </tr>
  </form>
</table>
<p>�U���O�d�� </p>
<table width="622" border="0" cellpadding="2" cellspacing="2">
  <tr bgcolor="#000000"> 
    <td> <div align="center"><font color="#FFFFFF">���~�W</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">����</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">����</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">���O</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">��T�� </font></div></td>
    <td> <div align="center"><font color="#FFFFFF">��O</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">����</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">���m</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">����</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">�Ȩ�</font></div></td>
    <td> <div align="center"><font color="#FFFFFF">����</font></div></td>
  </tr>
  <tr bgcolor="#FFFFCC"> 
    <td height="20"><font color="#990000">�y�P</font></td>
    <td><font color="#990000">���~</font></td>
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
    <td height="20">�s�]</td>
    <td><font color="#990000">���~</font></td>
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
    <td height="20">�y�P�\</td>
    <td><font color="#990000">���~</font></td>
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
    <td height="20">�l�u</td>
    <td>�u��</td>
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