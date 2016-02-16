<%money = clng(request.form("money"))
my=session("myname")
if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../error.asp?id=120"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr
sql="select * from 用戶 where 姓名='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%><script language=vbscript>
            MsgBox "你不是江湖中人!"
          location.href = "../jh.asp"
       </script><%
	conn.close
	response.end
else%>
<!--#include file="data1.asp"--><%
randomize()
         waittime=(rnd()*10000000)
         for i=0 to waittime
         next
sql="select * from 房屋 where 戶主='" & my & "' or  伴侶='" & my & "'"
set rs=conn.execute(sql)
yin=rs("伴侶銀兩")
if money>yin then%>
<script language=vbscript>
                         MsgBox "錯誤！您的金庫裡沒有那麼多銀兩。"
                         location.href = "xiaowu5.asp"
                    </script>
	<%elseif money>50000 or money<1 then%>
       <script language=vbscript>
                         MsgBox "錯誤！您填寫的數量不能小於1。"
                         location.href = "xiaowu5.asp"
                    </script><%
		else
       sql="update 房屋 set 伴侶銀兩=伴侶銀兩-'" & money & "' where 戶主='" & my & "' or  伴侶='" & my & "'"
       rs=conn.execute(sql)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr
       sql="update 用戶 set 銀兩=銀兩+'" & money & "' where 姓名='" & my & "'"
       rs=conn.execute(sql)
	conn.close
       if yin<0 then
       sql="update 用戶 set 銀兩=銀兩+'" & yin & "' where 姓名='" & my & "'"
       rs=conn.execute(sql)
	conn.close
       end if
	Response.Redirect "xiaowu5.asp"
			conn.close
			response.end
		end if
   rs.close
   set rs=nothing
end if
%>

<html><script language="JavaScript">                                                                  </script></html>