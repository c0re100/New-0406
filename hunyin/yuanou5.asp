<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 配偶 from 用戶 where 姓名='" & info(0) & "'",conn
peier=rs("配偶")
if rs("配偶")="無" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你還沒結婚怎麼離？');window.close();</script>"
	response.end
end if
rs.close
rs.open "select id from 用戶 where 姓名='" & peier & "'",conn
peierid=rs("id")
rs.close
rs.open "select id from 物品 where 物品名='流星淚' and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
			rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你忘了帶一顆流星淚了！');window.close();</script>"
	response.end			
                        else 
id=rs("id")
conn.execute "update 物品 set 數量=數量-1 where id="& id &" and 擁有者="&info(9)
				
		        end if
rs.close
rs.open "select id from 物品 where 物品名='流星淚' and 擁有者=" & peierid,conn
if rs.eof and rs.bof then
			conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,會員) values ('流星淚',"&peierid&",'物品',0,1000,0,0,0,1,10000000,111111,0,"&aaao&")"			
                        else 
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&peierid
				
		        end if
conn.execute "update 用戶 set 配偶='無',結婚時間=date(),小孩='無' where 姓名='" & info(0) & "'"
conn.execute "update 用戶 set 配偶='無',結婚時間=date(),小孩='無' where 姓名='" & peier & "'"
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
message="[" & info(0) & "]和[" & peier & "]感情破裂，宣告離婚！"
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="消息"
sd(195)="大家"
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="對"
sd(199)="<font color=FFFDAF>【離婚消息】"& message &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<title>離婚</title>

<STYLE type=text/css>A:hover {
	CURSOR: resize
}
A {
	TEXT-DECORATION: none
}
select       { background-color: #FFFFFF; font-size: 9pt; border-left: medium solid #999900; 
              border-right: medium solid #FFCC66; 
               border-top: medium solid #999900; 
               border-bottom: medium solid #FFCC66  }
</STYLE>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body text="#00FFFF" bgcolor="#FFFFFF" link="#99FF33" vlink="#FFFFFF" alink="#FF0000" leftmargin="0" topmargin="0">
<table width="691" border="0" cellspacing="0" cellpadding="0" height="119">
  <tr>
    <td background="lj1.gif"> 
      <p><img src="111.gif" width="77" height="111" align="left"><font color="#FF0000">閣下的姻緣已經解除了，希望閣下以後好好挑選自己的伴侶，別再<br>
        來找我了。</font></p>
      <p><a href="#" onClick="window.close()">好的。</a> </p>
    </td>
  </tr>
</table>
<div align="center"></div>

</body>

</html>

 