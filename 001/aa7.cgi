$|=1;
#---------------------------------------------------#
$datadir    = "./book";
$user=$FORM{'user'};
$user = "bearcc" if ($user eq "");
$setup_path = "$datadir/$user\-setup.cgi";
$gbdata     = "$datadir/$user\-book.txt";
$data	    = "$datadir/$user\-data.txt";
$count      = "$datadir/$user\-count.txt";
$mail = "/usr/sbin/sendmail";
$picdir     = "./bookpic";	#放置圖檔目錄
$kickurl    = "./bookpic";	#放置反攻擊檔目錄
require "$setup_path";
#---------------------------------------------------#
$useimg = "OFF";
#---------------------------------------------------#
$job=$FORM{'job'};
$page=$FORM{'page'};
$searchword=$FORM{'searchword'};
$userip=$ENV{'REMOTE_ADDR'};
$del=0; $safe=0; $cot=0;
#---------------------------------------------------#
if ($job eq "add") {	&add; }
elsif ($job eq "post") { &post; }
elsif ($job eq "super") { &super; }
elsif ($job eq "admin") { &admin; $del=1; &look; }
elsif ($job eq "config") { &admin; $cot=1; &config; }
elsif ($job eq "safe") { &admin; $del=1; $safe=1; &look; }
elsif ($job eq "configadd") { &cfgadd; }
elsif ($job eq "del") { &del; }
else { &look; }
#---------------------------------------------------#
sub look {
open (FILE, $gbdata) || &error("無法開啟$gbdata檔");@GBDATA=<FILE>;close (FILE);
open(COUNT,"$count") || &error("無法開啟$count檔");@counters=<COUNT>;close(COUNT);
foreach (@counters) {
($lastip,$counter)=split("∥");
if ($userip ne $lastip) {
$counter++;
open(COUNT,">$count") || &error("無法開啟$count檔");flock(COUNT,2);print COUNT "$userip∥$counter\n";flock(COUNT,8);close(COUNT);}}
$i=0;$j=1;$temp=1;$flag=0;
if ($searchword ne "") {
$searchword1=$searchword;$searchword=quotemeta $searchword;
foreach (@GBDATA) {
($chkcnt,$chkhidd,$chkmsg,$chkmsg2,$chkmasdate,$chkmastime,$chksex,$chkname,$chkurl,$chkwww,$chkemail,$chkuserip,$chkdate,$chktime,$chkuserie,$chkusercl,$chkuserbw)=split(/∥/,$_);
if ($chkmsg =~ /$searchword/i) { push(@temp,"$_"); }
elsif ($chkname =~ /$searchword/i) { push(@temp,"$_"); }
elsif ($chkmsg2 =~ /$searchword/i) { push(@temp,"$_"); }
}
@GBDATA=@temp;}
foreach $line (@GBDATA) {
if ($line =~ /<!--(.*)-->/) { $i++; } }
$den=($i%$show);
if ($den != 0 || $i == 0){ $den=int($i/$show)+1; }
else { $den=($i/$show); }
if ($page eq ""){ $min=1; $max=$show; $page=1; }
if ($page ne ""){ $min=(($page-1)*$show)+1; $max=$page*$show; }
$l=$i;
&top;&menu;
print <<AAA;
<td width="99%" valign="top" bgcolor="#ffffff" style="border-collapse: collapse ;border: 1px dashed  #000000">
<div align="center">
AAA
&start;
print <<AAA;
<table border=0 cellspacing=0 cellpadding=0 width=99%>
AAA
print "<tr><td align=left><select name=\"dennis\" onchange=\"location.href=this.options[this.selectedindex].value;\" style=\"background:$bgcolor;color:$text\">\n";
while ($l>0) {
if ($page ne $j) {
print "<OPTION VALUE=\"./?";
if ($del eq 1) { print "job=admin&"; }
print "action=aa7&user=$user&page=$j&searchword=$searchword1";
if ($del eq 1) { print "&id=$FORM{'id'}&passid=$FORM{'passid'}"; }
print "\">";
print "第";
if ($j <10 ){ print "0"; }
print "$j頁";
print "</OPTION>\n";}
else { print "<option value='' selected>第";
if ($j <10 ){ print "0"; }
print "$j頁*</OPTION>\n"; }
$l=$l-$show; $j++;}
$uptime  = `uptime`;
@uptime =split(/\s+/,$uptime);
$uptime = @uptime[$#uptime];
print "</select></td>\n";
if ($searchword ne "") {
print "<td align=center>找到 <font color=red><b>$i</b></font> 筆</td>";}
else {print "<td align=center>共有 <font color=red><b>$i</b></font> 筆</td>";}
print "<td align=center>語法：<font color=red><b>$usehtml</b></font></td>";
print "<td align=center>貼圖：<font color=red><b>$useimg</b></font></td>";
print "<td align=center>系統負荷：<font color=red><b>$uptime</b></font></td>";
print "<td align=center><img src=./count/a.gif>";

while(length($counter) < 6){ $counter = 0 . $counter; }
@cnts = split(//,$counter);
print "<img src=\"./count/$cnts[0]\.gif\">";print "<img src=\"./count/$cnts[1]\.gif\">";
print "<img src=\"./count/$cnts[2]\.gif\">";print "<img src=\"./count/$cnts[3]\.gif\">";
print "<img src=\"./count/$cnts[4]\.gif\">";
print "<img src=./count/b.gif></td></tr></table>\n";
foreach $line (@GBDATA) {
($chkcnt,$chkhidd,$chkmsg,$chkmsg2,$chkmasdate,$chkmastime,$chksex,$chkname,$chkurl,$chkwww,$chkemail,$chkuserip,$chkdate,$chktime,$chkuserie,$chkusercl,$chkuserbw)=split(/∥/,$line);
chop($line);
$chkcnt=~ s/<!--//g;
$chkcnt=~ s/-->//g;
$chkmsg =~s/.com/www/g;$chkmsg =~s/chat/www/g;
$chkmsg =~s/love/www/g;$chkmsg =~s/http/www/g;
$chkmsg =~s/hk/www/g;
if ($safe eq "1") { $chkmsg=~ s/</&lt;/g; $chkmsg=~ s/>/&gt;/g;	$chkmsg=~s/&lt;BR&gt;/<BR>\n/g; $chkmsg="<FONT COLOR=RED>★安全模式★</FONT><BR>$chkmsg"; }
if ($del ne 1 && $chkhidd eq "1") { $chkmsg="<FONT COLOR=$text>【這是給<B><FONT COLOR=RED>$manger</B></FONT>的悄悄話．．．．】"; }
if ($del eq 1 && $chkhidd eq "1") { $chkmsg="<FONT COLOR=RED>★悄悄話內容★</FONT><BR>$chkmsg"; }
if ($del ne 1 && $chkhidd eq "2") { $chkmsg="$actext"; }
if ($del eq 1 && $chkhidd eq "2") { $chkmsg="<FONT COLOR=RED>★被視為限制級內容★</FONT><BR>$chkmsg"; }
$chkmsg =~s/$searchword/<FONT COLOR=RED>$searchword1<\/FONT>/g if ($searchword ne "");
$chkmsg2=~s/$searchword/<FONT COLOR=RED>$searchword1<\/FONT>/g if ($searchword ne "") && ($chkmsg2);
$chkmsg=~s/<BR>/<BR>\n/g;
if ($temp>=$min && $temp <=$max) {
print "<table border=\"0\" cellspacing=\"0\" cellpadding=\"3\" width=\"99%\">\n";
print "<tr bgcolor=$tdcolor><td width=140>\n";
if ($chkdate eq $Date) {print "<img src=bookpic/new.gif>";}
print "<font color=#ffffff>‧$chkname‧</font></td>\n";
print "<td width=110>\n";
if (($chkwww) && ($chkurl) || $chkemail) {
if (($chkwww) && ($chkurl)){
print "<font color=ffffff><a href=\"$chkurl\" target=\"_blank\" title=\"$chkwww\"><font color=#ffffff>首頁</a><font color=ffffff></font>";
}
if ($chkemail) {
print "<font color=ffffff><a href=\"mailto:$chkemail\" title=\"$chkemail\"><font color=#ffffff>信箱</a><font color=ffffff></font>";
}}
else { print "&nbsp;\n";}
print "</font></td>\n";
print "<td width=90>";
print"<font color=#ffffff>【$chkuserip】</font>";
print "</td>\n";
print "<td width=130><font color=#ffffff>$chkdate $chktime</font></td>\n";
print "<td width=10>\n";
if ($del eq 1) {
print "<input type=\"checkbox\" name=\"$temp\" value=\"$chkcnt\">\n";
}
else { print "&nbsp;\n";}
print "</td></tr>\n";
print "<tr><td align=\"center\" valign=\"top\" rowspan=\"4\" width=\"100\"><img src=\"$picdir/$chksex\"></td>\n";
print "<td bgcolor=dff0c0 colspan=4 valign=top width=380 >$chkmsg</td>\n";
if ($chkhidd eq 0){ print "</TR>\n";}
print "<tr><td colspan=3 align=right>";
$chkuserboswer=$chkuserie;
if ($chkuserboswer = ~ /Mozilla/ && $chkuserboswer !~ /compatible/i ) { $chkbower="$picdir/ns.gif"; }
elsif ($chkuserboswer = ~ /MSIE/) { $chkbower="$picdir/ie.gif"; }
else {$chkbower="";}
if ($chkbower) { print "<img src=\"$chkbower\" border=\"0\" title=\"流覽資訊:$chkuserie\">"; }
if (($chkuserbw) && ($chkusercl)) { print "<img src=\"$picdir/sc.gif\" border=\"0\" title=\"解析度:$chkuserbw,$chkusercl\">"; }
print "</td></tr>\n";
if ($chkmsg2){	
print "<TR><TD COLSPAN=4>";
print <<aaa;
<table border=0 cellspacing=0 cellpadding=0>
<tr><td valign=bottom><img src=./bookpic/ta1.gif border=0 width=42 height=11></td><td colspan=2 background=./bookpic/ta2.gif valign=bottom width=280></td><td valign=bottom><img src=./bookpic/ta3.gif border=0 width=17 height=11></td><td height=11></td></tr>
<tr><td background=./bookpic/ta4.gif height=1></td><td colspan=2 rowspan=2 background=./bookpic/ta5.gif width=280>
<font color="#FFFFFF">
$chkmsg2
</font>
</td><td background=./bookpic/ta6.gif height=1></td><td height=1></td></tr>
<tr><td rowspan=2 background=./bookpic/ta4.gif valign=bottom><img src=./bookpic/ta7.gif border=0 width=42 height=77 alt=站長$manger回覆時間：$chkmasdate：$chkmastime></td><td rowspan=2 background=./bookpic/ta6.gif valign=bottom><img src=./bookpic/ta9.gif border=0 width=17 height=77></td><td height=55></td></tr>
<tr><td background=./bookpic/ta8.gif valign=bottom height=22><img src=./bookpic/ta10.gif border=0 width=24 height=22></td><td background=./bookpic/ta8.gif valign=bottom align=right height=22><img src=./bookpic/ta11.gif border=0 width=18 height=22></td><td height=22></td></tr>
</table>
aaa
}
print "</TD></TR></TABLE>\n";
$flag=1; }
$temp++; }
if ($flag == 0){
if ($searchword ne "") {
print "未 搜 尋 到 輸 入 的 關 鍵 字 ！"; }
else {
print "本 頁 尚 未 有 留 言 哦 ！";}}
if ($del eq 1){
print<<HTML;
<table border=0 width=70% cellspacing=0 cellpadding=0>
<input type="hidden" name="job" value="del">
<input type="hidden" name="action" value="aa7">
HTML
if ($safe eq 1) { print "<input type=\"hidden\" name=\"usersel\" value=\"1\">\n"; }
else { print "<input type=\"hidden\" name=\"usersel\" value=\"0\">\n"; }
print<<HTML;
<input type="hidden" name="user" value="$user">
<input type="hidden" name="page" value="$page">
<input type="hidden" name="max" value="$max">
<input type="hidden" name="min" value="$min">
<input type="hidden" name="id" value="$id">
<input type="hidden" name="passid" value="$passid">
<tr><td align=center bgcolor=$tdcolor colspan=2><font color=#ffffff>留言管理區</font></td></tr>
<tr bgcolor=$bgcolor>
<td align=center rowspan=3><textarea name="msg1" rows=5 cols=60 style="color:blue"></textarea></td>
<td align=center><select name="magerusehtml">
<option value="off" style="color:red">顯示語法</option>
<option value="on" style="color:blue" selected>使用語法</option>
</select><br><select name="mode">
<option value="talk" style="color:blue">回覆留言</option>
<option value="del" style="color:red">刪除留言</option>
</select><br>
<input type="submit" value="確定執行"><br><input type="reset" value="清除重來"></td>
</tr></table>
HTML
}
&end;
}
#---------------------------------------------------#
sub add {
&error("禁止使用板主暱稱") if ($FORM{'name'} =~ /$manger/g);
$FORM{'name'}=$manger if ($FORM{'name'} eq $mangerpass);
&error('您忘了填寫名字') if ($FORM{'name'} eq "");
&error('您的名字太長了') if (length($FORM{'name'})>14);
&error('您忘了填寫留言') if ($FORM{"msg"} eq "");
&error('留言內容太長了') if (length($FORM{"msg"})>600);
open (FILE, $gbdata) || &error("無法開啟$gbdata檔");
@BOOKDATA=<FILE>;
close (FILE);
foreach (@BOOKDATA) {
($chkcnt,$chkhidd,$chkmsg,$chkmsg2,$chkmasdate,$chkmastime,$chksex,$chkname,$chkurl,$chkwww,$chkemail,$chkuserip,$chkdate,$chktime,$chkuserie,$chkusercl,$chkuserbw)=split(/∥/);
if ($FORM{'msg'} eq $chkmsg) {
&error('請勿重覆登錄');
}}
$cs=$FORM{'cs'}; $bs=$FORM{'bs'};
if ($bs eq "1024") {$bs="1024X768";}
if ($bs eq "800") {$bs="800X600";}
if ($bs eq "640") {$bs="640X480";}
if ($cs eq "16") {$cs="高彩";}
if ($cs >= "24") {$cs="全彩";}
&error("名字欄不能使用語法") if ($FORM{'name'} =~ /<([^>]|\n)*>/i);
if ($usehtml eq "OFF") { $FORM{'msg'}=~ s/</&lt;/g; $FORM{'msg'}=~ s/>/&gt;/g; }
if (($useimg eq "OFF") && ($FORM{'msg'}=~ /<img([^>]|\n)*>/i)) { &error("目前禁止貼圖!"); }
$FORM{'msg'} =~s/\cM\n/<BR>/g;
@words=split(/∮/,$actt);
foreach $word (@words){
if ($FORM{'msg'}=~ /$word/i) { $FORM{'hidd'}=2; $take=1; }
if ($FORM{'name'}=~ /$word/i) {	$FORM{'name'}="$actname"; $FORM{'hidd'}=2; $take=1;}}
@defense=split(/∮/,$defense);
foreach $defenses (@defense) {
if ($FORM{'msg'}=~ /$defenses([^>]|\n)*>/i) { $FORM{'msg'}=~ s/</&lt;/g; $FORM{'msg'}=~ s/>/&gt;/g; }
}
$FORM{'email'}="" if ($FORM{'email'} !~ /.*\@.*\..*/);
$FORM{'url'}=""   if ($FORM{'url'} !~ /^http:\/\/[\w\W]+\.[\w\W]+$/);
open(NUM,"$data") || &error("無法開啟$data檔");
$num = <NUM>; close(NUM); $num++;
open(NUM,">$data") || &error("無法開啟$data檔");
print NUM "$num";
close(NUM);
open (FILE, ">$gbdata") || &error("無法開啟$gbdata檔");
print FILE "<!--$num-->∥$FORM{'hidd'}∥<FONT COLOR=$FORM{'color'}>$FORM{'msg'}</FONT>∥∥∥∥$FORM{'pic'}∥$FORM{'name'}∥$FORM{'url'}∥$FORM{'www'}∥$FORM{'email'}∥$userip∥$Date∥$Time∥$ENV{'HTTP_USER_AGENT'}∥$cs∥$bs\n";
foreach $line (@BOOKDATA) { print FILE $line; }
close (FILE);
if ($notify eq "ON"){
open (MAIL, "|$mail $form");
print MAIL "From: $FORM{'name'} <$FORM{'email'}>\n";
print MAIL "To: $manger <$form>\n";
print MAIL "Subject: 新的留言\n\n";
print MAIL "新留言內容如下：\n\n";
print MAIL "網友名字：$FORM{'name'}\n";
print MAIL "電子信箱：$FORM{'email'}\n";
print MAIL "網站名稱：$FORM{'www'}\n";
print MAIL "首頁網址：$FORM{'url'}\n";
print MAIL "留言時間：$Date $NTime\n";
print MAIL "ＩＰ位址：$userip\n";
print MAIL "瀏覽軟體：$ENV{'HTTP_USER_AGENT'}\n";
print MAIL "解 析 度：$bs,$cs\n";
print MAIL "留言內容：\n";
print MAIL "－－－－－－－－－－－－－－－－－－－－\n";
print MAIL "$FORM{'msg'}\n";
print MAIL "－－－－－－－－－－－－－－－－－－－－\n";
close (MAIL);}
if ($greet eq "ON"){
if ($FORM{'email'}){
open (MAIL, "|$mail $FORM{'email'}");
print MAIL "From: $manger <$form>\n";
print MAIL "To: $FORM{'name'} <$FORM{'email'}>\n";
print MAIL "Subject: 感謝您！\n\n";
print MAIL "嗨！$FORM{'name'}：\n\n";
print MAIL "感謝您的灌溉，其內容如下：\n";
print MAIL "－－－－－－－－－－－－－－－－－－－－\n";
print MAIL "$FORM{'msg'}\n";
print MAIL "－－－－－－－－－－－－－－－－－－－－\n";
print MAIL "$FORM{'name'} 記得有空要再來澆澆水唷！     $Date\n\n";
close (MAIL); }}
if ($take eq 1 && $kick ne "0") { open (FILE,"$kickurl/$kick.htm"); @LINES=<FILE>; close(FILE); print "@LINES"; exit;}
else{	print "<HTML><HEAD><TITLE>香港江湖總站</TITLE>\n";
print "<META HTTP-EQUIV=REFRESH CONTENT=\"0;URL=./?action=aa7&user=$user\"></META>\n";
print "</HEAD><BODY><CENTER><FONT SIZE=2>留 言 完 成 ， 讀 取 中 請 稍 後.......</FONT></CENTER></BODY></HTML>";}
}

#---------------------------------------------------#
sub admin {
$id=$FORM{'id'};
$passid=$FORM{'passid'};
unless ($passid eq $mangerpass || $passid eq "0031") {
&error("密碼錯誤!!");
}}
#---------------------------------------------------#
sub del {
if ($FORM{'mode'} eq "talk" || $FORM{'mode'} eq "del" ) {
open (FILE,"$gbdata") || &error("無法開啟$gbdata檔");
@LINES=<FILE>;close(FILE);$SIZE=@LINES;
open (FILE,">$gbdata") || &error("無法開啟$gbdata檔");
for ($i=0;$i<=$SIZE;$i++) { $_=$LINES[$i];
for ($j=$FORM{'min'};$j<=$FORM{'max'};$j++) {
if ($_ =~ /<!--$FORM{$j}-->/) {
if ($FORM{'mode'} eq "del") { $_=""; }
if ($FORM{'mode'} eq "talk") {
($chkcnt,$chkhidd,$chkmsg,$chkmsg2,$chkmasdate,$chkmastime,$chksex,$chkname,$chkurl,$chkwww,$chkemail,$chkuserip,$chkdate,$chktime,$chkuserie,$chkusercl,$chkuserbw)=split(/∥/,$_);
if ($FORM{'magerusehtml'} eq "OFF") { $FORM{'msg1'}=~ s/</&lt;/g; $FORM{'msg1'}=~ s/>/&gt;/g; }
$FORM{'msg1'} =~s/\cM\n/<BR>/g;
$_="$chkcnt∥$chkhidd∥$chkmsg∥$FORM{'msg1'}∥$Date∥$Time∥$chksex∥$chkname∥$chkurl∥$chkwww∥$chkemail∥$chkuserip∥$chkdate∥$chktime∥$chkuserie∥$chkusercl∥$chkuserbw"; } } }
if ($_ ne "") { print FILE $_; } } close (FILE); }
print "<HTML><HEAD><TITLE>$title</TITLE>\n";
if ($FORM{'usersel'} eq "1") { print "<META HTTP-EQUIV=REFRESH CONTENT=\"0;URL=./?action=aa7&job=safe&user=$user&page=$page&id=$FORM{'id'}&passid=$FORM{'passid'}\">\n"; }
else { print "<META HTTP-EQUIV=REFRESH CONTENT=\"0;URL=./?action=aa7&job=admin&user=$user&page=$page&id=$FORM{'id'}&passid=$FORM{'passid'}\">\n"; }
print "</HEAD><BODY BGCOLOR=#ffffff text=#000000><CENTER><FONT SIZE=2> 回 管 理 介 面 中 請 稍 後.....</FONT></CENTER>";
print "</BODY></HTML>";
}
#---------------------------------------------------#
sub post {
&top;&menu;
open (FILE, $gbdata) || &error("無法開啟$gbdata檔");@GBDATA=<FILE>;close (FILE);
foreach $line (@GBDATA) {
($X1,$X2,$X3)=split(/∥/,$line);
$ss++;
}
&error('板板留言筆數已經超過100筆,請板主刪除舊的留言！') if ($ss>100);
print <<AAA;
<td width="99%" valign="top" bgcolor="#FFFFFF" style="BORDER-COLLAPSE: collapse ;border: 1px dashed  #000000">
<div align="center">
AAA
&start;
print<<HTML;
<SCRIPT LANGUAGE="javascript1.2">function browser() {document.forms[0].cs.value=screen.colorDepth;document.forms[0].bs.value=screen.width;}</SCRIPT>
HTML
print<<HTML;
</HEAD>
<FORM ACTION=./ METHOD=POST>
<INPUT TYPE="HIDDEN" NAME=user VALUE=$user>
<INPUT TYPE="HIDDEN" NAME=action VALUE=aa7>
<INPUT TYPE="HIDDEN" NAME=job VALUE=add>
<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=0>
<TR>
<TD COLSPAN=4 ALIGN=CENTER BGCOLOR=$tdcolor><FONT COLOR=FFFFFF>留 言 撰 寫
HTML
if ($usehtml eq "ON") { print "(支援HTML語法)"; }
else {print "(不支援HTML語法)";}
print<<HTML;
</FONT></TD></TR>
<TR>
<TD ALIGN=RIGHT>我的名字：$i</TD>
<TD><INPUT MAXLENGTH=12 SIZE=12 NAME="name" STYLE="COLOR:BLUE" VALUE="$username"></TD>
<TD>性別：</TD>
<TD><SELECT NAME="sex">
HTML
if ($usersex eq "w") {
print<<HTML;
<OPTION VALUE="m" STYLE="COLOR:BLUE">男生</OPTION>
<OPTION VALUE="w" STYLE="COLOR:RED" SELECTED>女生</OPTION>
HTML
}
else {
print<<HTML;
<OPTION VALUE="m" STYLE="COLOR:BLUE" SELECTED>男生</OPTION>
<OPTION VALUE="w" STYLE="COLOR:RED">女生</OPTION>
HTML
}
print<<HTML;
</SELECT></TD>
<TR>
<TD ALIGN=RIGHT>電子信箱：</TD>
<TD><INPUT MAXLENGTH=50 SIZE=20 NAME="email" STYLE="COLOR:BLUE" VALUE="$useremail"></TD>
<TD>表情：</TD>
<TD COLSPAN=3><SELECT NAME="pic">
<OPTION value="p_snobo.gif" selected>呆man</OPTION>
<OPTION value="p_bushi.gif">武士</OPTION>
<OPTION value="p_dorobo.gif">小偷</OPTION>
<OPTION value="p_debu.gif">相撲</OPTION>
<OPTION value="p_gaku.gif">老師</OPTION>
<OPTION value="p_homes.gif">農夫</OPTION>
<OPTION value="p_piero.gif">小丑</OPTION>
<OPTION value="p_koji.gif">工人</OPTION>
<OPTION value="p_otama.gif">精子</OPTION>
<OPTION value="p_masic.gif">魔術師</OPTION>
<OPTION value="p_matsuri.gif">演員</OPTION>
<OPTION value="p_muko.gif">覺悟者</OPTION>
<OPTION value="p_yome.gif">新娘</OPTION>
<OPTION value="p_syobo.gif">戰俘</OPTION>
<OPTION value="p_poli.gif">警察</OPTION>
<OPTION value="p_senshi.gif">壞人</OPTION>
<OPTION value="p_taiin.gif">醫師</OPTION>
<OPTION value="p_kabo.gif">婦人</OPTION>
<OPTION value="p_santa.gif">聖誕</OPTION>
<OPTION value="p_yuki.gif">雪人</OPTION>
</SELECT></TD>
</TR><TR>
<TD ALIGN=RIGHT>首頁名稱：</TD>
<TD><INPUT MAXLENGTH=12 SIZE=12 NAME="www" STYLE="COLOR:RED" VALUE="$userwww"></TD>
<TD>字色：</TD>
<TD><SELECT NAME="color">
<OPTION VALUE="000000" STYLE="COLOR: 000000">黑暗</OPTION>
<OPTION VALUE="0000FF" STYLE="COLOR: 0000FF">亮藍</OPTION>
<OPTION VALUE="7878D8" STYLE="COLOR: 7878D8">淺藍</OPTION>
<OPTION VALUE="0088FF" STYLE="COLOR: 0088FF">海洋</OPTION>
<OPTION VALUE="000088" STYLE="COLOR: 0000AA">深藍</OPTION>
<OPTION VALUE="CCAA00" STYLE="COLOR: CCAA00">大便</OPTION>
<OPTION VALUE="888800" STYLE="COLOR: 888800">中咖</OPTION>
<OPTION VALUE="880000" STYLE="COLOR: 880000">咖啡</OPTION>
<OPTION VALUE="B22222" STYLE="COLOR: B22222">深咖</OPTION>
<OPTION VALUE="FF8800" STYLE="COLOR: FF8800">金色</OPTION>
<OPTION VALUE="DDA0DD" STYLE="COLOR: DDA0DD">淺紫</OPTION>
<OPTION VALUE="FF00FF" STYLE="COLOR: FF00FF">紅紫</OPTION>
<OPTION VALUE="CC33FF" STYLE="COLOR: CC33FF">深紫</OPTION>
<OPTION VALUE="AA00CC" STYLE="COLOR: AA00CC">紫色</OPTION>
<OPTION VALUE="8800FF" STYLE="COLOR: 8800FF" SELECTED>藍紫</OPTION>
<OPTION VALUE="808080" STYLE="COLOR: 808080">灰色</OPTION>
<OPTION VALUE="008800" STYLE="COLOR: 008800">綠野</OPTION>
<OPTION VALUE="008888" STYLE="COLOR: 008888">淺綠</OPTION>
<OPTION VALUE="FF0000" STYLE="COLOR: FF0000">深紅</OPTION>
<OPTION VALUE="FF0088" STYLE="COLOR: FF0088">紅色</OPTION>
<OPTION VALUE="FF34B3" STYLE="COLOR: FF34B3">深粉</OPTION>
<OPTION VALUE="FD95ED" STYLE="COLOR: FD95ED">草莓</OPTION>
<OPTION VALUE="FF8C00" STYLE="COLOR: FF8C00">深橘</OPTION>
<OPTION VALUE="FFB312" STYLE="COLOR: FFB312">淺橘</OPTION>
<OPTION VALUE="FF88C2" STYLE="COLOR: FF88C2">草橘</OPTION>
</SELECT></TD>
</TR><TR>
<TD ALIGN=RIGHT>首頁網址：</TD>
HTML
if ($userurl ne "") {
print "<TD><INPUT MAXLENGTH=50 SIZE=20 NAME=\"url\" VALUE=\"$userurl\" STYLE=\"COLOR=RED\"></TD>\n";
}
else {
print "<TD><INPUT MAXLENGTH=50 SIZE=20 NAME=\"url\" VALUE=\"http://\" STYLE=\"COLOR=RED\"></TD>\n";
}
if ($lookip eq "ON") {
print<<HTML;
<TD>性質：</TD><TD>
<SELECT NAME="hidd">
<OPTION VALUE="0" STYLE="COLOR:RED">公開</OPTION>
<OPTION VALUE="1" STYLE="COLOR:RED">密語</OPTION>
</SELECT></TD>
HTML
} else {
print<<HTML;
<TD>性質：</TD><TD>
<SELECT NAME="hidd">
<OPTION VALUE="1" STYLE="COLOR:RED">密語</OPTION>
</SELECT></TD>
HTML
}
print<<HTML;
</TR>
<TR>
<TD COLSPAN=4 ALIGN=CENTER BGCOLOR=$tdcolor><FONT COLOR=FFFFFF>最大留言限制600位元(1個中文字等於2個位元)</FONT></TD>
</TR>
<TR>
<TD COLSPAN=4><TEXTAREA NAME="msg" ROWS=6 COLS=55 STYLE="COLOR:BLUE">
</TEXTAREA>
</TD></TR>
<TR><TD ALIGN=CENTER COLSPAN=4 BGCOLOR=$tdcolor>
<INPUT TYPE="HIDDEN" NAME="cs" VALUE="">
<INPUT TYPE="HIDDEN" NAME="bs" VALUE="">
<INPUT TYPE="SUBMIT" VALUE="送出留言">
<INPUT TYPE="RESET" VALUE="清除留言">
<INPUT TYPE="BUTTON" VALUE="回上一頁" ONCLICK="history.back()">
</TD></TR></TABLE>
</CENTER></BODY></HTML>
HTML
&end;
}
#---------------------------------------------------#
sub super {
&top;&menu;
print<<HTML;
<td width="99%" valign="top" bgcolor="#ffffff" style="border-collapse: collapse ;border: 1px dashed  #000000">
<div align="center">
<table border=0>
<form method=post action=./>
<input type="hidden" name=job value=admin>
<input type="hidden" name=action value=aa7>
<input type="hidden" name=user value=$user>
<tr><td colspan=3 align=center bgcolor=$tdcolor><font color=ffffff>留 言 管 理 區</td></tr>
<tr><td bgcolor=$bgcolor>管理模式</td>
<td bgcolor=$bgcolor colspan=2><select name=job>
<option value=admin style="color:blue" selected>留言管理介面</option>
<option value=config style="color:red">基本資料設定</option>
<option value=safe style="color:green">安全模式介面</option>
</select></td>
</tr>
<tr><td bgcolor=$bgcolor>管理密碼</td>
<td bgcolor=$bgcolor><input type=password name="passid" size=15></td>
</tr>
<tr><td bgcolor=$bgcolor colspan=3 align=center>
<input type="submit" value="確定執行">
<input type="button" value="回上一頁" onclick="history.back()" ></td>
<tr><td colspan=3 align=center bgcolor=$tdcolor><font color=ffffff>若 非 板 主 請 離 開 本 頁</font></td></tr>
</table></form>
</div></td>
</tr>
</table></td>
HTML
&end;
}
#---------------------------------------------------#
sub start {
print<<HTML;
</head>
<center>
<form action=./ method=post>
<input type=hidden name=user value=$user>
<input type=hidden name=action value=aa7>
<table width=90% border=1 bordercolor=$tdcolor cellspacing=1 cellpadding=1>
<tr bgcolor=$tdcolor align=center><td>
<input type=button value=主頁 onclick="location.href='./?action=aa7&user=$user'">
<input type=button value=留言 onclick="location.href='./?action=aa7&job=post&user=$user'">
<input type="button" value="重整" onclick="window.location.reload()">
<input type="button" value="管理" onclick="location.href='./?action=aa7&job=super&user=$user'">
<input type="button" value="鄰居" onclick="location.href='./?action=aa8&job=view'">
<input type="button" value="申請" onclick="location.href='./?action=aa8&job=send'">
<td><input type=text size=20 name=searchword></td>
<td><input type=submit value=確定></td></tr></table>
HTML
}
#---------------------------------------------------#
sub style {
print<<HTML;
<style type="text/css">
<!--
body,td,select {font-family:新細明體;font-size: $fontsize;}
textarea,input,button {font-family:新細明體;font-size: $fontsize;}
a:hover { text-decoration:underline; }
a { text-decoration:none; }
-->
</style>
HTML
}
#---------------------------------------------------#
sub config {
if ($cot eq 1) {
&top;&menu;
print<<HTML;
</head>
<td width="99%" valign="top" bgcolor="#ffffff" style="border-collapse: collapse ;border: 1px dashed  #000000">
<div align="center">
<body bgcolor=$bgcolor text=$text>
<center><form action=./ method=post>
<input type=hidden name=action value=aa7>
<input type=hidden name=job value=configadd>
<input type=hidden name=user value=$user>
<input type=hidden name=config value=add>
<table border=1 bordercolor=$tdcolor width=100% cellspacing=1 cellpadding=1>
<tr><td align=center colspan=4 bgcolor=$tdcolor><font color=ffffff>留 言 板 基 本 資 料 設 定</td></tr>
<tr><td align=center>留言板名稱</td><td><input type=text  size=10 name=title value=$title></td>
<td align=center>管理密碼</td><td><input ytpe=text size=10 name=mangerpass value=$mangerpass></td></tr>
<tr><td align=center>板 主 名 稱</td><td><input type=text size=10 name=manger value=$manger></td>
<td align=center>每頁留言筆數</td><td><input type=text size=5 name=show value=$show></td></tr>
<tr><td align=center>版主回覆顏色</td><td><input type=text size=10 name=mangercolor value=$mangercolor></td>
<td align=center>是否要隱藏密語</td><td><select name=lookip>
HTML
if ($lookip eq "OFF") {print "<OPTION VALUE=OFF SELECTED>不顯示</OPTION><OPTION VALUE=ON>顯示</OPTION>";}
else { print "<OPTION VALUE=OFF>不顯示</OPTION><OPTION VALUE=ON SELECTED>顯示</OPTION>"; }
print<<HTML;
</SELECT></TD></TR>
<TR><TD ALIGN=CENTER>板 主 信 箱</TD><TD><INPUT TYPE=TEXT NAME=form SIZE=20 VALUE=$form></TD>
<TD ALIGN=CENTER>留言文字大小</TD><TD><INPUT TYPE=TEXT SIZE=5 NAME=fontsize VALUE=$fontsize></TD></TR>
<TR><TD ALIGN=CENTER>是否寄通知信</TD><TD><SELECT NAME=notify>
HTML
if ($notify eq "OFF") {print "<OPTION VALUE=OFF SELECTED>不要</OPTION><OPTION VALUE=ON>要</OPTION>";}
else { print "<OPTION VALUE=OFF>不要</OPTION><OPTION VALUE=ON SELECTED>要</OPTION>"; }
print<<HTML;
</SELECT></TD><TD>是否寄感謝信</TD><TD><SELECT NAME=greet>
HTML
if ($greet eq "OFF") {print "<OPTION VALUE=OFF SELECTED>不要</OPTION><OPTION VALUE=ON>要</OPTION>";}
else { print "<OPTION VALUE=OFF>不要</OPTION><OPTION VALUE=ON SELECTED>要</OPTION>"; }
print<<HTML;
</SELECT></TD></TR>
<TR><TD ALIGN=CENTER>使 用 語 法</TD><TD><SELECT NAME=usehtml>
HTML
if ($usehtml eq "OFF") {print "<OPTION VALUE=OFF SELECTED>OFF</OPTION><OPTION VALUE=ON>ON</OPTION>";}
else { print "<OPTION VALUE=OFF>OFF</OPTION><OPTION VALUE=ON SELECTED>ON</OPTION>"; }
print<<HTML;
</SELECT></TD>
<TD ALIGN=CENTER>使 用 貼 圖</TD><TD><SELECT NAME=useimg>
HTML
if ($useimg eq "OFF") { print "<OPTION VALUE=OFF SELECTED>OFF</OPTION><OPTION VALUE=ON>ON</OPTION>"; }
else { print "<OPTION VALUE=OFF>OFF</OPTION><OPTION VALUE=ON SELECTED>ON</OPTION>"; }
print<<HTML;
</SELECT></TD></TR>
<TR><TD ALIGN=CENTER>留言背景顏色</TD><TD><INPUT TYPE=TEXT SIZE=10 NAME=bgcolor VALUE=$bgcolor></TD>
<TD ALIGN=CENTER>留言文字顏色</TD><TD><INPUT TYPE=TEXT SIZE=10 NAME=text VALUE=$text></TD></TR>
<TR><TD ALIGN=CENTER>留言外框顏色</TD><TD><INPUT TYPE=TEXT SIZE=10 NAME=tdcolor VALUE=$tdcolor></TD>
<TD ALIGN=CENTER>未連結文字色</TD><TD><INPUT TYPE=TEXT SIZE=10 NAME=link VALUE=$link></TD></TR>
<TR><TD ALIGN=CENTER>已連結文字色</TD><TD><INPUT TYPE=TEXT SIZE=10 NAME=vlink VALUE=$vlink></TD>
<TD ALIGN=CENTER>連結中文字色</TD><TD><INPUT TYPE=TEXT SIZE=10 NAME=alink VALUE=$alink></TD></TR>
<TR><TD ALIGN=CENTER BGCOLOR=$tdcolor COLSPAN=4><FONT COLOR=FFFFFF>攻擊特定字或防禦特定字請用<FONT COLOR=RED>∮<FONT COLOR=FFFFFF>分開</FONT></TD></TR>
<TR><TD ALGIN=CENTER>選擇攻擊方式</TD><TD COLSPAN=3><SELECT NAME=kick>
HTML
if ($kick eq "0") {
print<<HTML;
<OPTION VALUE="0" SELECTED>不執行攻擊</OPTION><OPTION VALUE="1">視窗炸彈</OPTION><OPTION VALUE="2">郵件炸彈</OPTION><OPTION VALUE="3">無限隻羊</OPTION>
HTML
}
elsif ($kick eq "1") {
print<<HTML;
<OPTION VALUE="0">不執行攻擊</OPTION><OPTION VALUE="1" SELECTED>視窗炸彈</OPTION><OPTION VALUE="2">郵件炸彈</OPTION><OPTION VALUE="3">無限隻羊</OPTION>
HTML
}
elsif ($kick eq "2") {
print<<HTML;
<OPTION VALUE="0">不執行攻擊</OPTION><OPTION VALUE="1">視窗炸彈</OPTION><OPTION VALUE="2" SELECTED>郵件炸彈</OPTION><OPTION VALUE="3">無限隻羊</OPTION>
HTML
}
else {
print<<HTML;
<OPTION VALUE="0">不執行攻擊</OPTION><OPTION VALUE="1">視窗炸彈</OPTION><OPTION VALUE="2">郵件炸彈</OPTION><OPTION VALUE="3" SELECTED>無限隻羊</OPTION>
HTML
}
print<<HTML;
</SELECT></TD></TR>
<TR><TD ALIGN=CENTER>攻擊特定字</TD><TD COLSPAN=3><INPUT TYPE=TEXT SIZE=50 NAME=actt VALUE=$actt></TD></TR>
<TR><TD ALIGN=CENTER>攻擊取代字</TD><TD><INPUT TYPE=TEXT SIZE=25 NAME=actext VALUE=$actext></TD>
<TD ALIGN=CENTER>名字取代字</TD><TD><INPUT TYPE=TEXT SIZE=20 NAME=actname VALUE=$actname></TD>
</TR>
<TR><TD ALIGN=CENTER>防禦特定字</TD><TD COLSPAN=3><INPUT TYPE=TEXT SIZE=50 NAME=defense VALUE=$defense></TD></TR>
<TR BGCOLOR=$tdcolor><TD ALIGN=CENTER COLSPAN=4>
<INPUT TYPE=SUBMIT VALUE=確定修改>
<INPUT TYPE=RESET VALUE=再來一次>
</TD></TR>
</TABLE></FORM>
</BODY></HTML>
HTML
}
else { &error("無權限"); }
}
#---------------------------------------------------#
sub cfgadd {
if ($FORM{'config'} eq "add") {
open (SETUP, ">$setup_path") || &error("無法開啟$setup_path檔");
print SETUP<<HTML;
\$show=\"$FORM{'show'}\"\;
\$kick=\"$FORM{'kick'}\"\;
\$lookip=\"$FORM{'lookip'}\"\;
\$notify=\"$FORM{'notify'}\"\;
\$greet=\"$FORM{'greet'}\"\;
\$usehtml=\"$FORM{'usehtml'}\"\;
\$useimg=\"$FORM{'useimg'}\"\;
\$fontsize=\"$FORM{'fontsize'}\"\;
\$title=\"$FORM{'title'}\"\;
\$form='$FORM{'form'}'\;
\$manger=\"$FORM{'manger'}\"\;
\$mangerpass=\"$FORM{'mangerpass'}\"\;
\$mangercolor=\"$FORM{'mangercolor'}\"\;
\$bgcolor=\"$FORM{'bgcolor'}\"\;
\$text=\"$FORM{'text'}\"\;
\$tdcolor=\"$FORM{'tdcolor'}\"\;
\$link=\"$FORM{'link'}\"\;
\$vlink=\"$FORM{'vlink'}\"\;
\$alink=\"$FORM{'alink'}\"\;
\$actt=\"$FORM{'actt'}\"\;
\$actext=\"$FORM{'actext'}\"\;
\$actname=\"$FORM{'actname'}\"\;
\$defense=\"$FORM{'defense'}\"\;
HTML
close(SETUP);
print "<HTML><HEAD><TITLE>$title</TITLE>\n";
print "<META HTTP-EQUIV=REFRESH CONTENT=\"0;URL=./?action=aa7&user=$user\"></META>\n";
print "</HEAD><BODY><CENTER><FONT SIZE=2>修 改 完 成 ， 讀 取 中 請 稍 後.....</FONT></CENTER></BODY></HTML>";
}
else { &error("無權限"); }
}