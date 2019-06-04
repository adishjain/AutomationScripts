import socket

original_list = [ip.strip() for ip in open('ip_list.csv', 'r').readlines()]
i=0
for a in original_list:
	i+=1
	try:
		socket.inet_aton(a)
	except socket.error:
		print(i," ",a)
