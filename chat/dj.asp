<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
Response.Write "<script Language=Javascript>alert('���ܡG�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');window.close();</script>"
response.end
end if
if Application("ljjh_dj")=0 then
Response.Write "<script Language=Javascript>alert('���ܡG�A�ӱߤF!');</script>"
response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
randomize()
r=int(rnd*28)
select case r
case 1
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@��C���M<img src='../hcjs/jhjs/001/11001.gif' border=0 width=32 height=32  >"
rs.open "select id from ���~ where ���~�W='�C���M' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�C���M',"&info(9)&",'�k��',150,100,0,0,0,1,20000,11001,5,11001,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

end if
rs.close
case 2
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�ⶭ�ޤM<img src='../hcjs/jhjs/001/11002.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='���ޤM' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('���ޤM',"&info(9)&",'�k��',450,220,0,0,0,1,20000,11002,18,11002,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

end if
case 3
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�ⶭ�ޤM<img src='../hcjs/jhjs/001/11003.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�s��M' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�s��M',"&info(9)&",'�k��',650,350,0,0,0,1,20000,11003,28,11003,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

end if
rs.close
case 4
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�ⶭ�ޤM<img src='../hcjs/jhjs/001/11004.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�|�v�M' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�|�v�M',"&info(9)&",'�k��',860,420,0,0,0,1,20000,11004,38,11004,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
rs.close
case 5
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@��U�H��<img src='../hcjs/jhjs/001/11005.gif' border=0 width=32 height=32  >"
rs.open "select id from ���~ where ���~�W='�U�H��' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�U�H��',"&info(9)&",'�k��',4000,800,0,0,0,1,20000,11005,60,11005,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
case 6
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�Ӭy�P<img src='../hcjs/jhjs/001/111111.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�y�P' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�y�P',"&info(9)&",'���~',0,1000,0,0,0,1,10000000,111111,0,111111,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
rs.close
case 7
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@���s�W�C<img src='../hcjs/jhjs/001/11007.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�s�W�C' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�s�W�C',"&info(9)&",'�k��',35000,3190,0,0,0,1,20000,11007,145,11007,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
rs.close
case 8
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@��C���C<img src='../hcjs/jhjs/001/11008.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�C���C' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�C���C',"&info(9)&",'�k��',38000,4070,0,0,0,1,20000,11008,185,11008,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
case 9
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@��L���C<img src='../hcjs/jhjs/001/11006.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�L���C' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�L���C',"&info(9)&",'�k��',25000,1870,0,0,0,1,20000,11006,85,11006,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
rs.close
case 10
mess=info(0)&"�����Ǫ���o20�I�g��!"
conn.execute "update �Τ� set allvalue=allvalue+20,��O=��O-100 where id="&info(9)
case 11
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�Ӭï]�٫�<img src='../hcjs/jhjs/001/6100.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�ï]�٫�' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�ï]�٫�',"&info(9)&",'�˹�',0,80,0,0,120,1,20000,6100,12,6100,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
rs.close
case 12
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@���ŬP�٫�<img src='../hcjs/jhjs/001/6101.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�ŬP�٫�' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�ŬP�٫�',"&info(9)&",'�˹�',0,100,0,0,150,1,20000,6101,20,6101,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
rs.close
case 13
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�Ӭ��P�٫�<img src='../hcjs/jhjs/001/6102.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='���P�٫�' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('���P�٫�',"&info(9)&",'�˹�',0,180,0,0,400,1,20000,6102,40,6102,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
rs.close
case 14
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�Ӧ�G�٫�<img src='../hcjs/jhjs/001/6103.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='��G�٫�' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('��G�٫�',"&info(9)&",'�˹�',0,280,0,0,600,1,20000,6103,50,6103,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
rs.close
case 15
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�Ӥ���٫�<img src='../hcjs/jhjs/001/6104.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='����٫�' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('����٫�',"&info(9)&",'�˹�',0,380,0,0,800,1,20000,6104,60,6104,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
rs.close
case 16
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�ӤӶ��٫�<img src='../hcjs/jhjs/001/6105.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�Ӷ��٫�' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�Ӷ��٫�',"&info(9)&",'�˹�',0,480,0,0,1000,1,20000,6105,80,6105,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
rs.close
case 17
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�Ӭï]����<img src='../hcjs/jhjs/001/6106.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�ï]����' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�ï]����',"&info(9)&",'�˹�',0,580,0,0,1400,1,20000,6106,100,6106,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
rs.close
case 18
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�ӷN������<img src='../hcjs/jhjs/001/6107.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�N������' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�N������',"&info(9)&",'�˹�',0,880,0,0,1800,1,20000,6107,120,6107,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
rs.close
case 19
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�ӥë���<img src='../hcjs/jhjs/001/6108.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�ë���' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�ë���',"&info(9)&",'�˹�',0,1480,0,0,2800,1,20000,6108,160,6108,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
rs.close
case 20
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�ӫC�ɬ޵P<img src='../hcjs/jhjs/001/2100.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�C�ɬ޵P' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�C�ɬ޵P',"&info(9)&",'����',0,500,0,0,1000,1,20000,2100,50,2100,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
end if
rs.close
case 21
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�ө����޵P<img src='../hcjs/jhjs/001/2101.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�����޵P' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�����޵P',"&info(9)&",'����',0,1500,0,0,3000,1,20000,2101,80,2101,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

end if
rs.close
case 22
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�Ӥ����޵P<img src='../hcjs/jhjs/001/2102.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�����޵P' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�����޵P',"&info(9)&",'����',0,2500,0,0,4000,1,20000,2102,140,2102,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

end if
rs.close
case 23
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�ӥͤƬ޵P<img src='../hcjs/jhjs/001/2103.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�ͤƬ޵P' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�ͤƬ޵P',"&info(9)&",'����',0,4500,0,0,6000,1,20000,2103,180,2103,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

end if
rs.close
case 24
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@�ӷ���<img src='../hcjs/jhjs/001/9100.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='����' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('����',"&info(9)&",'�ī~',0,100,0,10000000,0,1,20000,9100,0,9100,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

end if
rs.close
case 25
randomize()
rr=int(rnd*10000)
if rr=5 then
mess="����<font color=B0D0E0>"&info(0)&"</font>�����Ǫ��o��@���s�]<img src='../hcjs/jhjs/001/50000.gif' border=0  width=32 height=32 >"
rs.open "select id from ���~ where ���~�W='�s�]' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�s�]',"&info(9)&",'���~',0,100,0,0,0,1,10000000,50000,0,50000,"&aaao&")"
else
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

end if
rs.close
else

mess=info(0)&"�����Ǫ��o��@����l�A�@�p<img src='img/251.gif'>3�U��!"
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+30000,��O=��O-100 where id="&info(9)
end if
case else
mess=info(0)&"�����Ǫ��o��@����l�A�@�p<img src='img/251.gif'>3�U��!"
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+30000,��O=��O-100 where id="&info(9)

end select

Application.Lock
Application("ljjh_dj")=0
Application.UnLock
conn.close
set conn=nothing
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application.Lock
Application("ljjh_line")=line+1
Application.UnLock
for i=1 to 190
sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="����"
sd(195)="�j�a"
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="��"
sd(199)="<font color=B0D0E0>�i��������j</font>"&mess
sd(200)=session("nowinroom")
Application.Lock
Application("ljjh_sd")=sd
Application.UnLock
%>
