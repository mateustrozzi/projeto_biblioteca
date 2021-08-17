w32tm /config /syncfromflags:manual /manualpeerlist:0.pool.ntp.org
net stop w32time
net start w32time
w32tm /resync /rediscover
exit


