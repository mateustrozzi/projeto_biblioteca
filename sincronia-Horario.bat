w32tm /config /syncfromflags:manual /manualpeerlist:0.pool.
tp.org
net stop w32time
net start w32time
w32tm /resync /rediscover
call msg.vbs
exit


