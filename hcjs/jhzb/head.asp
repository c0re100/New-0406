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
<title>我的包袱</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="style.css">

</head>
<body background="back.gif">
<div align="left">
  <table width="700" border="1" align="center" borderColorDark=#ffffff borderColorLight=#ff9999 
 cellPadding=5>
    <tr> 
      <td background="f001.gif" width="169" height="247"> 
        <table width="100%" border="0" height="57" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="49" width="38" rowspan="2">&nbsp;</td>
            <td height="49" width="48" rowspan="2"> 
              <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

 rs.open "SELECT sm FROM 物品 WHERE 擁有者=" & info(9) &" and 數量>0 and 裝備=true and (類型='頭盔')  order by 銀兩 ",conn
if not rs.bof and not rs.eof then%>
              <div align="center"><img src="../jhjs/001/<%=rs("sm")%>.gif" border="0" width="48" height="49"%> 
              </div>
              <%end if%>
            </td>
            <td height="19" width="44">&nbsp;</td>
          </tr>
          <tr> 
            <td height="34" width="44"> 
              <% rs.close
			  rs.open "SELECT sm FROM 物品 WHERE 擁有者=" & info(9) &" and 數量>0 and 裝備=true and (類型='裝飾')  order by 銀兩 ",conn
if not rs.bof and not rs.eof then%>
              <div align="center"><img src="../jhjs/001/<%=rs("sm")%>.gif" border="0" width="44" height="24"%> 
              </div>
              <%end if%>
            </td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="81" width="43"> 
              <% rs.close
			  rs.open "SELECT sm FROM 物品 WHERE 擁有者=" & info(9) &" and 數量>0 and 裝備=true and (類型='左手')  order by 銀兩 ",conn
if not rs.bof and not rs.eof then%>
              <img src="../jhjs/001/<%=rs("sm")%>.gif" border="0"  width="32" height="32"> 
              <%end if%>
            </td>
            <td height="96" width="63" rowspan="2" valign="top"> 
              <% rs.close
			  rs.open "SELECT sm FROM 物品 WHERE 擁有者=" & info(9) &" and 數量>0 and 裝備=true and (類型='盔甲')  order by 銀兩 ",conn
if not rs.bof and not rs.eof then%>
              <img src="../jhjs/001/<%=rs("sm")%>.gif" border="0"  width="60" height="86"> 
              <%end if%>
            </td>
            <td height="81" width="55" valign="bottom"> 
              <div align="right"> 
                <% rs.close
			  rs.open "SELECT sm FROM 物品 WHERE 擁有者=" & info(9) &" and 數量>0 and 裝備=true and (類型='右手')  order by 銀兩 ",conn
if not rs.bof and not rs.eof then%>
                <div align="center"><img src="../jhjs/001/<%=rs("sm")%>.gif" border="0"  width="46" height="67"> 
                </div>
                <%end if%>
              </div>
            </td>
          </tr>
          <tr> 
            <td height="37" width="43">&nbsp;</td>
            <td height="37" width="55">&nbsp;</td>
          </tr>
        </table>
        <table width="53%" border="0" height="63" cellspacing="0" cellpadding="0" align="center">
          <tr> 
            <td> 
              <% rs.close
			   rs.open "SELECT sm FROM 物品 WHERE 擁有者=" & info(9) &" and 數量>0 and 裝備=true and (類型='雙腳')  order by 銀兩 ",conn

if not rs.bof and not rs.eof then%>
              <div align="center"><img src="../jhjs/001/<%=rs("sm")%>.gif" border="0"  width="32" height="32"> 
                <%end if%>
              </div>
            </td>
          </tr>
        </table>
      </td>
      <td valign="top" rowspan="2" width="517"> 
        <div align="center"><img src="dao.gif" width="40" height="15">武器裝備一覽<img src="dao1.gif" width="40" height="15"> 
          <font color="#CC0000" face="幼圓"><a href="javascript:this.location.reload()">刷新</a></font></div>
        <div align="center"> 
          <table width="501" border="1" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
            <tr> 
              <td width="12%"> 
                <div align="center">物品名</div>
              </td>
              <td width="6%"> 
                <div align="center">類型</div>
              </td>
              <td width="5%"> 
                <div align="center">類型</div>
              </td>
              <td width="7%"> 
                <div align="center">數量</div>
              </td>
              <td width="7%"> 
                <div align="center">加攻擊</div>
              </td>
              <td width="9%"> 
                <div align="center">加防御</div>
              </td>
              <td width="8%"> 
                <div align="center">堅固度</div>
              </td>
              <td width="8%"> 
                <div align="center">等級</div>
              </td>
              <td width="10%"> 
                <div align="center">價值</div>
              </td>
              <td width="28%"> 
                <div align="center">操作</div>
              </td>
            </tr>
            <%
			rs.close
			  dim page
			  page=request.querystring("page")
PageSize = 6
			 
			rs.open "SELECT ID,物品名,類型,數量,攻擊,防御,堅固度,等級,銀兩,裝備,sm FROM 物品 WHERE 擁有者=" & info(9) &" and 數量>0 and (類型='頭盔' or 類型='左手' or 類型='右手' or 類型='雙腳' or 類型='裝飾' or 類型='盔甲')  order by 銀兩 ",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page

count=0
do while not rs.eof and count<rs.PageSize
%>
            <tr align="center"> 
              <td width="12%" height="15"><%=rs("物品名")%></td>
              <td width="6%" height="15"><img src="../jhjs/001/<%=rs("sm")%>.gif" border="0" width="32" height="32"%> 
              </td>
              <td width="5%" height="15"> <%=rs("類型")%> </td>
              <td width="7%" height="15"> <%=rs("數量")%> </td>
              <td width="7%" height="15"><%=rs("攻擊")%></td>
              <td width="9%" height="15"><%=rs("防御")%></td>
              <td width="8%" height="15"><%=rs("堅固度")%></td>
              <td width="8%" height="15"><%=rs("等級")%></td>
              <td width="10%" height="15"><%=rs("銀兩")%></td>
              <td width="28%" height="15"> 
                <%if rs("裝備")=false then%>
                <a href="wupin2.asp?id=<%=rs("ID")%>">裝備</a> 
                <%else%>
                <a href="wupin2.asp?id=<%=rs("ID")%>">解除</a> 
                <%end if%>
              </td>
            </tr>
            <%
rs.movenext
count=count+1
loop
%>
            <tr align="center"> 
              <td colspan="10" height="15"><font size="2">[共<font color="#7e22e4"><b><%=rs.pagecount%></b></font>頁] 
                [<a
href="head.asp?page=<%=page-1%>">上一頁</a>][第<%=page%>頁][<a
href="head.asp?page=<%=page+1%>">下一頁</a>]</font></td>
            </tr>
          </table>
          注意：物品隻能裝備一個，多買隻能送人！</div>
      </td>
    </tr>
    <tr> 
      <td valign="top" height="91" width="151"> 
        <%
	  rs.close
	  rs.open "SELECT id,姓名,等級,攻擊,防御,a1,a2,a3,a4,a5,a6,b1,b2,b3,b4,b5,b6 FROM 用戶 WHERE id="&info(9),conn
%>
        <img src="../../ico/<%=info(10)%>-2.gif" width="84" height="84" align="left"> 
        <p>ID：<%=rs("id")%><br>
          姓名：<%=rs("姓名")%><br>
          等級：<%=rs("等級")%><br>
          攻擊：<%=rs("等級")*150%><br>
          防御：<%=rs("等級")*200%><br>
          <br>
          額外攻擊力：<%=rs("a1")+rs("a2")+rs("a3")+rs("a4")+rs("a5")+rs("a6")%><br>
          額外防御力：<%=rs("b1")+rs("b2")+rs("b3")+rs("b4")+rs("b5")+rs("b6")%></p>
      </td>
    </tr>
  </table>
  <table width="505" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
<td>
<div align="center">
<input onClick="JavaScript:window.close()" title=關閉窗口 type=button value="關閉窗口" name="button2">
</div>
</td>
</tr>
</table>
</div>
</body>
</html>