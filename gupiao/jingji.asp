<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>

<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> 
<title>股票經紀人辦公室</title>
<style>
input,body,select,td{font-size:14;line-height:160%}
</style>
</head>

<body bgcolor="#990000">
<p align=center style='font-size:16;color:yellow'>歡迎<%=info(0)%>光臨股票經紀人的辦公室！
<p align=center style='font-size:14;color:blue'><a href=index.asp>返回股票市場</a></p>
<!--#include file="jhb.asp"-->
<%
sql="select 經紀人 from 客戶 where 帳號='"&info(0)&"'"
set rs=conn.execute(sql)
if not rs.eof then
jjr=rs(0)
%>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table><tr><td>
<p align=center style="font-size:14;color:#000000"></p> 
</td></tr>  

<tr><td style="color:red;font-size:9pt">  
<br><p align=center>您的股票經濟人<%=jjr%>在此為您服務</p><br>
<a href=cha.asp>查詢今日交易</a>
<a href=cun.asp>存錢進股票帳戶</a><br>
<a href=qu.asp>從股票帳戶提錢</a>
<br> 
<p align=center><a href=index.asp>離開</a></p>
</table></table>  
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing%>
<p align=center style='font-size:16;color:yellow'>申請股票帳戶/更換經濟人 
<form method=POST action='jingji2.asp' id=form2 name=form2>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
        <table width=100%>
          <tr> 
            <td width="47%">帳號： 
              <input type=text name=name size=10 value="<%=info(0)%>" maxlength="0">
            </td>
            <td width="53%">經濟人： 
              <select name=jjr size=1>
                <option value="阿Pip">阿Pip</option>
                <option value="千豬格格">千豬格格</option>
                <option value="印度阿三">印度阿三</option>
                <option value="錢夫人">錢夫人</option>
                <option value="阿土伯">阿土伯</option>
                <option value="羊百萬">羊百萬</option>
                <option value="約翰喬">約翰喬</option>
                <option value="沙隆巴斯">沙隆巴斯</option>
                <option value="忍太郎">忍太郎</option>
                <option value="無錢蓮">無錢蓮</option>
                <option value="索鏍絲">索鏍絲</option>
              </select>
            </td>
          </tr>
          <tr> 
            <td colspan=2 align=center> 
              <input type=submit value=確定 id=submit2 name=submit2>
              <input type=reset value="取消" id=reset2 name=reset2>
            </td>
          </tr>
          <tr> 
            <td colspan=2 style='font-size:9pt'> 
              <hr>
              1.立戶時，隻須填第一個密碼；我們將同時為您安排一位股票經紀人，幫助您完成股票交易，每筆買賣將收取交易金額的1%作為酬勞。<br>
              2.修改密碼時，第一個密碼為舊密碼，第二個密碼為新密碼！由於本系統采用用戶加密的辦法，不會出現被人偷錢的事情，所以密碼可以不必修改！ 
              <br>
              3.立戶時帳號可以任意填寫，但是每個江湖人士隻可以申請一個股票帳號！</td>
          </tr>
            <td width="47%"> 
          </table>
</td></tr>
</table>
</form>
</body>
</html>
