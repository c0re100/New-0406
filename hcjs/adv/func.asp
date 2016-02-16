<%
function huayuan(name1,sh1,sy1)
%> <!--#include file="data.asp"--> <%
sql="SELECT 寶物 FROM 用戶 WHERE 姓名='" & name1 & "'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
      huayuan=""
else
      if rs("寶物")="靈劍水晶石" then
            huayuan="" & name1 & "你已經有靈劍水晶石了必須修煉以後纔能到孤島尋寶！" 
      else
            randomize timer
            rand=Int(79 * Rnd + 10)            
                 if rand>=10 and rand<=50 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，你發現了別人留下的100兩銀子！"
                        sql="update 用戶 set 體力=體力-11,銀兩=銀兩+100,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=51 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，你發現了別人留下的5000兩銀子！"
                        sql="update 用戶 set 體力=體力-11,銀兩=銀兩+5000,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=52 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，你發現了別人留下的500兩銀子！"
                        sql="update 用戶 set 體力=體力-11,銀兩=銀兩+500,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=53 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，你踫到了野獸，力戰之下你打死了野獸，但卻傷了3000內力！"
                        sql="update 用戶 set 體力=體力-11,內力=內力-3000,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=54 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，被腳下石頭一拌摔了一跤，丟了100兩銀子！"
                        sql="update 用戶 set 體力=體力-11,銀兩=銀兩-100,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=55 or rand=85 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，你發現了一奇特的山草藥，一喫之下，身中劇毒而死！"                         
                        sql="update 用戶 set 狀態='死',操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql 
                       
                        sql="insert into 人命(死者,時間,兇手,死因) values ('" & name1 & "',now(),'無','在孤島中喫錯有毒食物而死')"
                        conn.execute sql
                       call boot(name1)
                 elseif rand=56 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，你發現了古代寶藏，正想去拿之際，寶藏的機關打開，你被亂箭射傷，損失體力10000！"
                        sql="update 用戶 set 體力=體力-10000,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=57 or rand=86 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，你失足跌落懸崖，從此江湖少了一個無名大俠！"
                        sql="update 用戶 set 狀態='死' where 姓名='" & name1 & "'"
                        conn.execute sql 
                       
                        sql="insert into 人命(死者,時間,兇手,死因) values ('" & name1 & "',now(),'無','在孤島失足跌落懸崖')"
                        conn.execute sql
                      call boot(name1)
                 elseif rand=58 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，巧遇一絕色女子被老虎追趕，你英雄救美，魅力上升1000！"
                        sql="update 用戶 set 體力=體力-100,魅力=魅力+1000,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=59 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，巧遇一世外高人，經指點，你道德上升200！"
                        sql="update 用戶 set 體力=體力-11,道德=道德+200,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=60 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，洗了洗臉！魅力上升50"
                        sql="update 用戶 set 體力=體力-10,魅力=魅力+50 where,操作時間=now() 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=61 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，巧遇一絕色女子，你把她救出孤島，道德上升500！"
                        sql="update 用戶 set 體力=體力-1,道德=道德+500,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=62 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，發現了一個大寶藏，經過一番冒險，你成功拿到五十萬寶藏及武功絕學,內力上升1000！"
                        sql="update 用戶 set 內力=內力+1000,銀兩=銀兩+500000,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand>=63 or rand<=66 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，你發現了別人留下的1000兩銀子！"
                        sql="update 用戶 set 體力=體力-11,銀兩=銀兩+1000,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=67 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，洗了洗臉！魅力上升100"
                        sql="update 用戶 set 體力=體力-10,魅力=魅力+100,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=68 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，你發現了大寶藏，但在拿的途中被亂箭射死！"
                        sql="update 用戶 set 狀態='死' where 姓名='" & name1 & "'"
                        conn.execute sql 
                       
                        sql="insert into 人命(死者,時間,兇手,死因) values ('" & name1 & "',now(),'無','孤島中拿寶藏時被亂箭射死')"
                        conn.execute sql
                     call boot(name1)
                 elseif rand=69 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，踫到毒蛇，幸虧武功高強，虛驚一場！"
                        sql="update 用戶 set 體力=體力-11 where,操作時間=now() 姓名='" & name1 & "'"
conn.execute sql
                 elseif rand=70 or rand=87 then
huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，遇到仇敵剛巧也在此地，在大鬥30回合後，你不敵而死於仇敵劍下！"
                        sql="update 用戶 set 狀態='死' where 姓名='" & name1 & "'"
                        conn.execute sql 
                       
                        sql="insert into 人命(死者,時間,兇手,死因) values ('" & name1 & "',now(),'無','在孤島冒險中被殺')"
                        conn.execute sql
                       call boot(name1)
                 elseif rand=71 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，你發現了別人留下的1萬兩銀子！"
                        sql="update 用戶 set 體力=體力-11,銀兩=銀兩+10000,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=72 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，突然在牆角發現1兩銀子！"
                        sql="update 用戶 set 體力=體力-11,銀兩=銀兩+1,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=73 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，你發現了一種奇特的山草藥，一喫之下，內力增進2000！"
                        sql="update 用戶 set 體力=體力-11,內力=內力+2000,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=74 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，你發現了別人留下的300兩銀子！"
                        sql="update 用戶 set 體力=體力-11,銀兩=銀兩+300,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=75 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，你發現了別人留下的200兩銀子！"
                        sql="update 用戶 set 體力=體力-11,銀兩=銀兩+200,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=76 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，洗了洗臉！魅力上升100"
                        sql="update 用戶 set 體力=體力-10,魅力=魅力+100,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=77 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，失足跌下山坡！大難不死，體力損失5000，但大難不死，必有後福，發現5000兩銀子"
                        sql="update 用戶 set 體力=體力-5000,銀兩=銀兩+5000,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=78 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，巧遇一世外高人，經指點，你道德上升2000！"
                        sql="update 用戶 set 體力=體力-11,道德=道德+2000,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=79 or rand=88 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，你在冒險中踫到獅子，力戰下還是慘死於獅子口下！"
                        sql="update 用戶 set 狀態='死' where 姓名='" & name1 & "'"
                        conn.execute sql 
                                              sql="insert into 人命(死者,時間,兇手,死因) values ('" & name1 & "',now(),'無','在孤島冒險中身亡')"
                        conn.execute sql
                        call boot(name1)
                 elseif rand=80 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，你發現了上乘的內力心法，內力大進1000點！"
                        sql="update 用戶 set 體力=體力-11,內力=內力+1000,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=81 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，巧遇一絕色女子，你色心大起，對其施暴，道德下降1000！"
                        sql="update 用戶 set 體力=體力-1000,道德=道德-1000,操作時間=now() where 姓名='" & name1 & "'"
                        conn.execute sql
                 elseif rand=82 or rand=89 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，你因找不到食物，活活餓死了，從此江湖少了一個無名大俠！"
                        sql="update 用戶 set 狀態='死' where 姓名='" & name1 & "'"
                        conn.execute sql 
                        sql="insert into 人命(死者,時間,兇手,死因) values ('" & name1 & "',now(),'無','在孤島冒險中身亡')"
                        conn.execute sql
                     call boot(name1)
                 elseif rand=83 then
                        huayuan="" & name1 & "在孤島上" & sy1 & sh1 & "，意外地發現了先人在孤島留下的寶物藏身地，但是裡面什麼都沒有！"
             end if    
       end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
sub boot(to1)
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(to1) then
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=onlinelist(i)
newonlinelist(js+1)=onlinelist(i+1)
newonlinelist(js+2)=onlinelist(i+2)
newonlinelist(js+3)=onlinelist(i+3)
newonlinelist(js+4)=onlinelist(i+4)
newonlinelist(js+5)=onlinelist(i+5)
js=js+6
else
kickip=onlinelist(i+2)
end if
next
useronlinename=useronlinename&" "
if kickip<>"" then
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
else
Application.UnLock
Response.Redirect "manerr.asp?id=204&kickname=" & server.URLEncode(kickname)
end if
Application.UnLock
end sub
%>