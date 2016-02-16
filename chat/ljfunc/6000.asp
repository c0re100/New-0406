<%
'求神問卜
function kk(fn1,to1,toname)
if toname="大家" or toname=info(0) or toname=Application("ljjh_automanname")  then
	Response.Write "<script language=JavaScript>{alert('對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要10級以上方可用！');}</script>"
	Response.End
end if
conn.execute "update 用戶 set 內力=內力+100000 where id="&info(9)
randomize()
rnd1=int(rnd()*4)
if rnd1<3 then
kk=【求神問卜】info(0) & "誠心祈求抽到一支【健康簽】,※ 算命先生 ※解籤說道:這是一支健康簽.算是下下籤.籤文上說你今年身體欠佳需要好好休養.不可過份操勞.要不然內力會慢慢消失※"&info(0) & 傳授你2000點內力讓你身體安康.....
"" 
else
conn.execute "update 用戶 set 體力=體力-100000 where id="&info(9)
randomize()
rnd1=int(rnd()*4)
if rnd1<3 then
kk=【求神問卜】info(0) & "誠心祈求抽到一支【噩運簽】：※ 算命先生 ※解籤說道:這是一支噩運簽.非常非常的不好.看來你是噩運難逃了本算命師只能消耗你的體力來為你解除噩運※"&info(0) & ※扣你體力2000.....

"" 
else
conn.execute "update 用戶 set 金錢=金錢+200000 where id="&info(9)
randomize()
rnd1=int(rnd()*4)
if rnd1<3 then
kk=【求神問卜】info(0) & "誠心祈求抽到一支【發財簽】：※ 算命先生 ※解籤說道:這是一支發財簽相當的好.照籤文上所說你不管做什麼財運都會跟這來.真是恭喜你了這是上等簽※"&info(0) & ※順便贈送你20萬兩.....
"" 
else
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>