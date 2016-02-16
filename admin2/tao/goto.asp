<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
  s=now()-session("tao_time")
  if int(s*1000000)<30 then
  response.write "請不要刷呀，悠著點，江湖上不少人呢！"
  response.end
  end if
%>
<!--#include file="data.asp"-->
<%
  randomize timer
  t=int(rnd*100)
    if t<30 then
    session("tao_time")=now()
  randomize timer
  jin=int(rnd*100)
  sql="select * from 淘金者 where 擁有者='" & info(0) & "'"
  set rs=conn.execute(sql)
  if rs.eof or rs.bof then
       sql="insert into 淘金者(擁有者,金,銀,銅) values('" & info(0) & "'," & jin & ",0,0)"
       conn.execute(sql)
       else
       sql="update 淘金者 set 金=金+" & jin & " where 擁有者='" & info(0) & "'"
       conn.execute(sql)
       end if
       conn.execute("update 礦石 set 總點=總點+1,總流=總流+" & jin& " where 礦種='金'")
       rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "<%=info(0)%>" & "恭喜您!獲得" & <%=jin%> & "點金礦!"
  location.href="index.asp"
</script>
<%
    elseif t<65 and t>=30 then
  session("tao_time")=now()
  randomize timer
  yin=int(rnd*200)
  sql="select * from 淘金者 where 擁有者='" & info(0) & "'"
  set rs=conn.execute(sql)
  if rs.eof or rs.bof then
       sql="insert into 淘金者(擁有者,金,銀,銅) values('" & info(0) & "',0," & yin & ",0)"
       conn.execute(sql)
       else
       sql="update 淘金者 set 銀=銀+" & yin & " where 擁有者='" & info(0) & "'"
       conn.execute(sql)
       end if
       conn.execute("update 礦石 set 總點=總點+1,總流=總流+" & yin& " where 礦種='銀'")
       rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "<%=info(0)%>" & "恭喜您!獲得" & <%=yin%> & "點銀礦!"
  location.href="index.asp"
</script>
<%
    else
  session("tao_time")=now()
  randomize timer
  tong=int(rnd*300)
  sql="select * from 淘金者 where 擁有者='" & info(0) & "'"
  set rs=conn.execute(sql)
  if rs.eof or rs.bof then
       sql="insert into 淘金者(擁有者,金,銀,銅) values('" & info(0) & "',0,0," & tong & ")"
       conn.execute(sql)
       else
       sql="update 淘金者 set 銅=銅+" & tong & " where 擁有者='" & info(0) & "'"
       conn.execute(sql)
       end if
       conn.execute("update 礦石 set 總點=總點+1,總流=總流+" & tong & " where 礦種='銅'")
       rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "<%=info(0)%>" & "恭喜您!獲得" & <%=tong%> & "點銅礦!"
  location.href="index.asp"
</script>
<%
    end if
%>