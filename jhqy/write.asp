<!--#include file="config.asp"-->
<!--#include file="conn.asp"-->


<html>
<meta http-equiv="Content-Type" content="text/html; charset=big5"> 
<head>
<style type="text/css">
<!--
a {  text-decoration: none}  
a:hover {  text-decoration: underline} 
table {  font-size: 9pt}
body,table,p,td,input,select,textarea {  font-size: 9pt} 
-->
</style>
<title>k666軟件園 收集 http://www.vv66.net</title>
</head>

<body bgcolor="<%=bgcolor%>" text="<%=textcolor%>" link="<%=linkcolor%>">

<form method="post" action="save.asp" align="center">
  <center><table
    border="1" cellspacing="1" width="400" bordercolor="<%=titlelightcolor%>"
    height="310">
        <tr>
            
        <td align="center" colspan="2" bgcolor="<%=titledarkcolor%>"><font color="#FFFF00">【<%=title%>】 
          祈求願望</font></td>
        </tr>
        <tr>
            <td colspan="2" height="23"><p align="center">&nbsp;<font
            color="<%=titlelightcolor%>">帶 （*） 必須填寫 </font></p>
            </td>
        </tr>
        <tr>
            <td align="center" width="100" height="23"><p
            align="left">您的姓名： </p>
            </td>
            <td height="23"><p align="left"><input type="text"
            size="20" maxlength="20" name="name"> （*） </p>
            </td>
        </tr>

        <tr>
            <td align="center" width="100" height="24"><p
            align="left">您的性別： </p>
            </td>
            <td height="24"><p align="left"><select name="sex"
            size="1">
                <option value="男"> 男 </option>
                <option value="女"> 女 </option>
            </select> </p>
            </td>
        </tr>
        <tr>
            <td align="center" width="100" height="23"><p
            align="left">您的信箱： </p>
            </td>
            <td height="23"><p align="left"><input type="text"
            size="30" maxlength="50" name="email"> （*） </p>
            </td>
        </tr>
        <tr>
            <td align="center" width="100" height="23"><p
            align="left">您的主頁： </p>
            </td>
            <td height="23"><p align="left">
            <input type="text"
            size="30" maxlength="50" name="homepage"
            value="http://">
             </p>
            </td>
        </tr>
        <tr>
            <td width="100" height="23"><p align="left">願望類別：
            </p>
            </td>
            <td height="23"><p align="left">
            <select name="wishtype"
            size="1">
              <option value="love">戀  愛</option>
                <option value="study">學  業</option>
                <option value="health">健  康</option>
                <option value="family">家  庭</option>
                <option value="work">事  業</option>
                <option value="future">將  來</option>
                <option value="wealth">財  富</option>
                <option value="life">生  活</option>
            </select> （*） </p>
            </td>
        </tr>
        <tr>
            <td width="100" class="pt9">居住地方：</td>
            <td class="pt9"><table border="0" width="300">
                <tr>
                    <td>
                <input type="radio" name="address" value="中國" checked>中國</td>
              <td>
                <input type="radio" name="address" value="9港">9港</td>
                    <td><input type="radio" name="address" value="日本">日本 </td>
                    <td><input type="radio" name="address" value="臺灣">臺灣</td>
                    <td><input type="radio" name="address" value="新加玻">新加玻</td>
                </tr>
                <tr>
                    <td><input type="radio" name="address" value="馬來西亞">馬來西亞</td>
                    <td><input type="radio" name="address" value="加拿大">加拿大</td>
                    <td><input type="radio" name="address" value="澳洲">澳洲</td>
                    <td><input type="radio" name="address" value="美國">美國</td>
                    <td><input type="radio" name="address" value="英國">英國</td>
                </tr>
                <tr>
                    <td><input type="radio" name="address" value="法國">法國</td>
                    <td><input type="radio" name="address" value="非洲">非洲</td>
                    <td><input type="radio" name="address" value="澳門">澳門</td>
                    <td><input type="radio" name="address" value="其它">其它</td>
                    <td> （*） </td>
                </tr>
            </table>
            </td>
        </tr>
        <tr>
            <td width="100">祈求願望： </td>
            <td>
          <div align="center">
            <textarea name="info" rows="4" cols="45"></textarea>
            <br>
            （* 100字以內） </div>
        </td>
        </tr>
        <tr>
            
        <td colspan="2" bgcolor="<%=titledarkcolor%>" height="15" align=center><font color="#FFFF00">送出後合起雙眼誠心祈求...</font></td>
        </tr>
        <tr>
            <td align="center" colspan="2" height="25"><input
            type="submit" style="border:1 solid <%=titlelightcolor%>;color:<%=titlelightcolor%>;background-color:white" value="送  出">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input style="border:1 solid <%=titlelightcolor%>;color:<%=titlelightcolor%>;background-color:white" type="reset" value="清  除"> </td>
        </tr>
    </table>
    </center></div>
</form>
<!--#include file="copyright.asp"-->
</body>
</html>
