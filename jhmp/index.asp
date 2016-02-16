<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<head>
<title>江湖門派</title>
<meta http-equiv=Content-Type content="text/html; charset=big5">
<LINK href="../css.css" rel=stylesheet>
</head>
<body bgcolor="#FFFFFF" background="../ff.jpg"
leftmargin="0" topmargin="7">
<div align="center">
  <script>
if (!document.layers&&!document.all)
event="test"
function showtip2(current,e,text){

if (document.all&&document.readyState=="complete"){
document.all.tooltip2.innerHTML='<marquee style="border:1px solid black">'+text+'</marquee>'
document.all.tooltip2.style.pixelLeft=event.clientX+document.body.scrollLeft+10
document.all.tooltip2.style.pixelTop=event.clientY+document.body.scrollTop+10
document.all.tooltip2.style.visibility="visible"
}

else if (document.layers){
document.tooltip2.document.nstip.document.write('<b>'+text+'</b>')
document.tooltip2.document.nstip.document.close()
document.tooltip2.document.nstip.left=0
currentscroll=setInterval("scrolltip()",100)
document.tooltip2.left=e.pageX+10
document.tooltip2.top=e.pageY+10
document.tooltip2.visibility="show"
}
}
function hidetip2(){
if (document.all)
document.all.tooltip2.style.visibility="hidden"
else if (document.layers){
clearInterval(currentscroll)
document.tooltip2.visibility="hidden"
}
}

function scrolltip(){
if (document.tooltip2.document.nstip.left>=-document.tooltip2.document.nstip.document.width)
document.tooltip2.document.nstip.left-=5
else
document.tooltip2.document.nstip.left=150
}

</script>
</div>
<div id="tooltip2" style="position:absolute;clip:rect(0 150 50 0);width:111px;background-color:#99FF99; top: 31px; left: 64px; visibility: hidden; height: 13px"> 
</div>
<div align="center">
  <table width="590" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="90" valign="top"> <P align="center"><a href="#" onClick="window.open('creat.htm','zhuangti','scrollbars=no,resizable=no,width=620,height=450')"><br>
          自創門派</a></P>
        <P align="center"><a href="#" onClick="window.open('setwg.asp','zhuangti','scrollbars=no,resizable=no,width=620,height=450')">武功設置</a></P>
        <P ALIGN="CENTER"><a href="#" onClick="window.open('managemp.asp?id=<%=info(5)%>','zhuangti','scrollbars=no,resizable=no,width=620,height=450')">掌門管理</a></P></td>
      <td width="554" align="center" valign="top"> <table width="500" border="1" cellpadding="1" cellspacing="0" align="left" bordercolorlight="#000000" bordercolordark="#FFFFFF">
          <tr align="center" bgcolor="#FF6600"> 
            <td width="120" height="11">門 派</td>
            <td width="120" height="11">掌 門</td>
            <td width="80" height="11">弟 子 數</td>
            <td width="60" height="11">適 合</td>
            <td width="100" height="11">操 作</td>
          </tr>
          <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim jhmenpai(100,2)
rs.open "SELECT 門派 FROM 門派",conn
do while not rs.bof and not rs.eof
   	menpai=menpai+1
	jhmenpai(menpai,1)=trim(rs("門派"))
	rs.movenext
loop
for xx=1 to menpai
tmprs=conn.execute("Select count(*) As 數量 from 用戶 where  門派='"&jhmenpai(xx,1)&"'")
regren=tmprs("數量")
sql="update 門派 set 弟子數="& regren &" where 門派='"& jhmenpai(xx,1) &"'"
Set Rs=conn.Execute(sql)
next
rs.open "SELECT 門派,掌門,弟子數,適合,簡介,創立時間,幫派基金 FROM 門派",conn
do while not rs.bof and not rs.eof
%>
          <%if trim(rs("門派"))<>"官府" and trim(rs("門派"))<>"遊俠" then%>
          <tr bgcolor="#99CCFF"> 
            <td><a href="mp.asp?my=<%=info(0)%>&amp;id=<%=rs("門派")%>"><%=rs("門派")%></a></td>
            <td><%=rs("掌門")%></td>
            <td><%=rs("弟子數")%></td>
            <td><%=rs("適合")%></td>
            <td> <a href="#" onClick="window.open('mp1.asp?my=<%=info(0)%>&amp;id=<%=rs("門派")%>','zhuangti','scrollbars=no,resizable=no,width=430,height=290')" onMouseOver="showtip2(this,event,'<%=rs("簡介")%> 該門派建立於：<%=rs("創立時間")%> 現有基金：<%=rs("幫派基金")%>')" onMouseOut="hidetip2()">加入</a> 
              | <a href="#" onClick="window.open('mp2.asp?my=<%=info(0)%>&amp;id=<%=rs("門派")%>','zhuangti','scrollbars=no,resizable=no,width=430,height=290')">離開</a> 
            </td>
          </tr>
          <%end if%>
          <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
        </table></td>
    </tr>
  </table>
</div>
</body>
</html>
