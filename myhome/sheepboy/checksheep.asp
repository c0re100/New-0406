<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")

%>

<!--#include file="data1.asp"-->
<!--#INCLUDE file="check.asp"-->
<%
cname=check(request("nick"),"小羊名字") 
%>
<%
IF Request.ServerVariables("ReQuest_Method") = "POST" THEN
sheepname=request.form("nick")
if sheepname="" or len(sheepname)="" then
%>
<script language="Vbscript">
msgbox"小孩的名字填寫不正確！",0,"提示"
history.back
</script>
<%
else
rs.open"select dellifeday from rules",conn,1,1
if rs.bof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('沒有規則存在！');location.href = 'javascript:history.go(-1)';}</script>"
else
dellifeday=rs("dellifeday")
rs.close
rs.open"select logintoday,feeddate,life,hungry,clean,sheephappy,sheephealth from sheep where sheepname='"&sheepname&"' and (user='"&info(0)&"' or sheep002='"&info(0)&"')",conn,1,1
if rs.bof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('您不是這個小孩的主人！');location.href = 'javascript:history.go(-1)';}</script>"
else
logintoday=rs("logintoday")
feeddate=rs("feeddate")
life=rs("life")
hungry=rs("hungry")
clean=rs("clean")
sheephappy=rs("sheephappy")
sheephealth=rs("sheephealth")
rs.close
if datediff("d",logintoday,date)<>0 then
tempdatetoday=date
conn.execute"update sheep set logintoday='"&tempdatetoday&"' where sheepname='"&sheepname&"' and (user='"&info(0)&"' or sheep002='"&info(0)&"')"
temptime=datediff("d",feeddate,date)
templife=life-dellifeday*temptime
temphungry=hungry-dellifeday*temptime/2
if temphungry<0 then
temphungry=0
end if
tempclean=clean-dellifeday*temptime/2
if tempclean<0 then
tempclean=0
end if
tempsheephappy=sheephappy-dellifeday*temptime/2
if tempsheephappy<0 then
tempsheephappy=0
end if
tempsheephealth=sheephealth-dellifeday*temptime/2
if tempsheephealth<0 then
tempsheephealth=0
end if
'conn.execute"update sheep set clean='"&tempclean&"',sheephappy='"&tempsheephappy&"',sheephealth='"&tempsheephealth&"',hungry='"&temphungry&"' where sheepname='"&sheepname&"' and id="&tempid&" and user='"&info(0)&"'"
conn.execute"update sheep set life='"&templife&"',hungry='"&temphungry&"',clean='"&tempclean&"',sheephappy='"&tempsheephappy&"',sheephealth='"&tempsheephealth&"' where sheepname='"&sheepname&"' and user='"&info(0)&"'"
end if
rs.open"select life,clean,sheephappy,sheephealth,hungry from sheep where sheepname='"&sheepname&"' and (user='"&info(0)&"' or sheep002='"&info(0)&"')"
if rs("life")<=0 or rs("clean")<=0 or rs("sheephappy")<=0 or rs("sheephealth")<=0 or rs("hungry")<=0 then
rs.close
conn.execute"delete * from sheep where sheepname='"&sheepname&"' and user='"&info(0)&"'"
conn.close
 Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT 配偶 FROM 用戶 WHERE 姓名='" & info(0) & "'"
rs.open sql,conn,1,3
peier=rs("配偶")
conn.execute("update 用戶 set 銀兩=銀兩-1000000,小孩='無',孩名='無'  where 姓名='"&info(0)&"'")
conn.execute("update 用戶 set 銀兩=銀兩-1000000,小孩='無',孩名='無'  where 姓名='"&peier&"'")
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language="vbscript">
msgbox"由於您沒有精心照顧小孩！您的小孩已經死了，重新再養一個吧。",0,"FLUSH"
history.back
</script>
<%
else
rs.close
rs.open"select life,sheep001,sheep002,sheepname,sheepdate,hungry,clean,sheephealth,sheepxb,sheephappy,wugong,leinl,gongji,feedsheepday,milk from sheep where sheepname='"&sheepname&"' and (user='"&info(0)&"' or sheep002='"&info(0)&"')",conn,1,1
old=int(now()-rs("sheepdate"))
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">


<HTML><HEAD><TITLE>靈劍江湖-江湖育兒院</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type><LINK 
href="jh/clubcom.css" rel=stylesheet type=text/css>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#000000 text=#ffffff topMargin=0>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=80%>
  <TBODY> 
  <TR>
    <TD background=jh/history_top_bg.gif vAlign=top width=100%> 
      <TABLE align=center border=0 cellPadding=2 cellSpacing=0 class=p9 
      width="100%">
        <TBODY> 
        <TR>
          <TD height=25 width="31%">&nbsp;</TD>
          <TD height=25 vAlign=top width="33%">
            <DIV align=center><font color="#000000">江湖育兒院</font></DIV>
          </TD>
          <TD height=25 vAlign=top 
width="36%">&nbsp;</TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD background=jh/history_table_bg.gif height=158 vAlign=top width="100%"> 
      <TABLE align=center border=0 cellPadding=0 cellSpacing=0 class=mountain 
      width=85%>
        <TBODY> 
        <TR>
          <TD vAlign=top width="100%"><BR>
            <TABLE align=center background=jh/table_bg.gif border=0 
            cellPadding=0 cellSpacing=0 class=table1 width="99%">
              <TBODY> 
              <TR>
                <TD align=middle vAlign=center> 
                  <P class=p9><font color="#FF0000"><b>【江湖提示】</b></font><br>
                    這兒是喂養小孩的地方，你小孩就住在這兒，看著她無憂無慮的生活著，你覺得在江湖裡不在孤獨。不過你每天得來這照顧她哦，不然她會死掉的哦<br>
                  </P>
                  <table border="0" width="97%" cellspacing="1" cellpadding="0"
                height="330">
                    <tr> 
                      <td class="p9" colspan="4" height="18">
                        <div align="center">
                          <p>孩子他爹：<%=rs("sheep001")%> <br>
                            孩子他媽：<%=rs("sheep002")%></p>
                          <p> 
                            <%if rs("sheepxb")="男" then %>
                            寶貝兒子:<%=rs("sheepname")%> <font size="2"><img src="baby3.gif" width="144" height="133" align="top"> 
                            <%else%>
                            寶貝女兒： <%=rs("sheepname")%> <img src="girl15.gif" width="45" height="57" align="top"> 
                            <%end if%>
                            </font></p>
                        </div>
                      </td>
                    </tr>
                    <tr> 
                      <td class="p9" width="25%" height="18">□-小孩生命值：</td>
                      <td class="p2" width="25%" height="18"> 
                        <%
if rs("life")>80 then
typelife="健康"
end if
if rs("life")<=80 and rs("life")>60 then
typelife="良好"
end if
if rs("life")<=60 and rs("life")>40 then
typelife="虛弱"
end if
if rs("life")<=40 and rs("life")>20 then
typelife="生病"
end if
if rs("life")<=20 and rs("life")>0 then
typelife="病危"
end if
if rs("life")<=0 then
typelife="死亡"
end if
response.write typelife
                      %>
                      </td>
                      <td class="p2" width="25%" height="18"><span class="p9">□-小孩的年齡</span>：</td>
                      <td class="p2" width="25%" height="18"><%=old%></td>
                    </tr>
                    <tr> 
                      <td class="p3" width="25%" height="18"><span class="p9">□-饑餓程度</span>：</td>
                      <td class="p3" width="25%" height="18"><%=rs("hungry")%></td>
                      <td class="p3" width="25%" height="18"><span class="p9">□-清潔程度</span>：</td>
                      <td class="p3" width="25%" height="18"><%=rs("clean")%></td>
                    </tr>
                    <tr> 
                      <td class="p2" width="25%" height="18"><span class="p9">□-健康程度</span>：</td>
                      <td class="p2" width="25%" height="18"><%=rs("sheephealth")%></td>
                      <td class="p9" width="25%" height="18">□-快樂程度：</td>
                      <td class="p2" width="25%" height="18"><%=rs("sheephappy")%></td>
                    </tr>
                    <tr> 
                      <td class="p2" width="25%" height="18"><span class="p9">□-武功程度</span>：</td>
                      <td class="p2" width="25%" height="18"><%=rs("wugong")%></td>
                      <td class="p9" width="25%" height="18">□-內力程度：</td>
                      <td class="p2" width="25%" height="18"><%=rs("leinl")%></td>
                    </tr>
                    <tr> 
                      <td class="p2" width="25%" height="18"><span class="p9">□-攻擊程度</span>：</td>
                      <td class="p2" width="25%" height="18"><%=rs("gongji")%></td>
                      <td class="p9" width="25%" height="18">&nbsp;</td>
                      <td class="p2" width="25%" height="18">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td class="p9" width="100%" colspan="4" height="17"><font color="#FF0000">□-光榮媽媽敬告：</font></td>
                    </tr>
                    <tr> 
                      <td class="p9" width="100%" colspan="4" height="71">□-你通過不同的方法照顧你的小孩會得到不同的工作量，你只有長時間的積累經驗，每次耐心細致的觀察你的小孩每天的情況，纔能把你們孩子培養成為當代高手。你的小孩每天可管理8次(一周期),照顧小孩五個周期以上就可以到孤兒院打工並且可以帶著你們的孩子闖蕩江湖。<br>
                        □-注：每次照料小孩花費50元，當小孩子的四項值中任何一項值為0時，小孩會死去。</td>
                    </tr>
                    <tr> 
                      <td class="p2" width="25%" height="38"> 
                        <form method="POST" action="sheepeat.asp">
                          <p align="center"> 
                            <input type="submit" value="喂奶" 
                        name="B1" style="font-family: 新細明體; font-size: 12px">
                            <font              
                        color="#E18C59">饑餓+10</font> 
                            <input type="hidden"              
                        name="num" value="<%=tempid%>">
                            <input type="hidden"              
                        name="sheepname" value="<%=sheepname%>">
                          </p>
                        </form>
                      </td>
                      <td class="p2" width="25%" height="38"> 
                        <form method="POST" action="sheepclean.asp">
                          <p align="center"> 
                            <input type="submit" value="洗澡" 
                        name="B12" style="font-family: 新細明體; font-size: 12px">
                            <font              
                        color="#E18C59">清潔+10</font> 
                            <input type="hidden"              
                        name="num" value="<%=tempid%>">
                            <input type="hidden"              
                        name="sheepname" value="<%=sheepname%>">
                          </p>
                        </form>
                      </td>
                      <td class="p2" width="25%" height="38"> 
                        <form method="POST" action="sheepsun.asp">
                          <p align="center"> 
                            <input type="submit" value="陽光" 
                        name="B12" style="font-family: 新細明體; font-size: 12px">
                            <font              
                        color="#E18C59">健康+10</font> 
                            <input type="hidden"              
                        name="num" value="<%=tempid%>">
                            <input type="hidden"              
                        name="sheepname" value="<%=sheepname%>">
                          </p>
                        </form>
                      </td>
                      <td class="p2" width="25%" height="38"> 
                        <form method="POST" action="sheeppei.asp">
                          <p align="center"> 
                            <input type="submit" value="陪伴" 
                        name="B12" style="font-family: 新細明體; font-size: 12px">
                            <font              
                        color="#E18C59">快樂+10</font> 
                            <input type="hidden"              
                        name="num" value="<%=tempid%>">
                            <input type="hidden"              
                        name="sheepname" value="<%=sheepname%>">
                          </p>
                        </form>
                      </td>
                    </tr>
                    <tr> 
                      <td class="p2" width="25%" height="38"> 
                        <form method="POST" action="sheepwu.asp">
                          <p align="center"> 
                            <input type="submit" value="練武" 
                        name="B1" style="font-family: 新細明體; font-size: 12px">
                            <font color="#E18C59">武功+50</font> 
                            <input type="hidden"              
                        name="num2" value="<%=tempid%>">
                            <input type="hidden"              
                        name="sheepname" value="<%=sheepname%>">
                          </p>
                        </form>
                      </td>
                      <td class="p2" width="25%" height="38"> 
                        <form method="POST" action="sheepnl.asp">
                          <p align="center"> 
                            <input type="submit" value="靜坐" 
                        name="B1" style="font-family: 新細明體; font-size: 12px">
                            <font color="#E18C59">內力+50</font> 
                            <input type="hidden"              
                        name="num22" value="<%=tempid%>">
                            <input type="hidden"              
                        name="sheepname" value="<%=sheepname%>">
                          </p>
                        </form>
                      </td>
                      <td class="p2" width="25%" height="38"> 
                        <form method="POST" action="sheepgj.asp">
                          <p align="center"> 
                            <input type="submit" value="戰場" 
                        name="B1" style="font-family: 新細明體; font-size: 12px">
                            <font color="#E18C59">攻擊+50</font> 
                            <input type="hidden"              
                        name="num222" value="<%=tempid%>">
                            <input type="hidden"              
                        name="sheepname" value="<%=sheepname%>">
                          </p>
                        </form>
                      </td>
                      <td class="p2" width="25%" height="38">&nbsp;</td>
                    </tr>
                  </table>
                  <div align="center">
                    <p>&nbsp;</p>
                    <p><a href="javascript:window.close()"><font color="#FFFFFF">關閉窗口</font></a> 
                    </p>
                  </div>
                  <P class=p9> 
                    <%if rs("feedsheepday")>=3 then%>
                    <br>
                  </P>
                  <table border="0" width="470" cellspacing="1" cellpadding="0" 
            height="60">
                    <tr> 
                      <td class="p9" width="100%"><b><font color="#FF0000">恭喜！今天您的小孩在孤兒院打工了，打工時間為</font></b><%=rs("milk")%><b><font color="#FF0000">個工作單位</font></b><br>
                        <font color="#000000"><b><font color="#FF0000">快到</font></b></font><a href="indexsheep.asp"><font color="#FFFFFF">孤兒院</font></a><b><font color="#FF0000">去工資吧！</font></b></td>
                    </tr>
                  </table>
                  <P class=p9 align="center"><a href="javascript:window.close()"><font color="#FFFFFF">關閉窗口</font></a> 
                  </P>
                  <P align=center class="p9">&nbsp;</P>
                  </TD></TR></TBODY></TABLE>
            <P align=center>&nbsp; </P>
          </TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD bgColor=#847939 height=1 width="100%"><IMG height=1 src="jh/page_point.gif" 
      width=1></TD>
  </TR></TBODY></TABLE>
<%rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end if                  
                
%> </BODY></HTML>
<%                 
            
end if                  
end if                  
end if                  
end if                  
end if               
%> 

