import os
import xlrd
from netmiko import ConnectHandler
import re
import time
import threading
from tqdm import tqdm

BASE_DIR = os.path.dirname(os.path.abspath(__file__))  #get the script's root path 

BASE_DIR_configuration = os.path.join(BASE_DIR, time.strftime("%Y%m%d%H%M%S", time.localtime()))
os.mkdir(BASE_DIR_configuration)

for file in os.listdir(BASE_DIR):
	if file == 'input_demo_v0.1.xls':
		file_name = file

wb = xlrd.open_workbook(filename=file_name)
sheet1 = wb.sheet_by_index(0)#get the sheet by index and it start with 0
#  sheet2 = wb.sheet_by_name('Devices')#get the sheet by name

device_infor_all = []
title_rows = sheet1.row_values(0)#get row values

for number_rows in range(1, sheet1.nrows):
	value_rows = sheet1.row_values(number_rows)
	device_infor_all.append(dict(zip(title_rows, value_rows)))

def devices_conn(device, position):
	cmd_set = wb.sheet_by_name(device['cmd_set']).col_values(0)

	cisco1 = { 
    "device_type": "cisco_ios" if device['ssh_telnet'] == "ssh" else "cisco_ios_telnet",
    "host": device['IP'],
    "username": device['Username'],
    "password": device['Password'],
    "secret": None if device['Enable'] == '' else device['Enable'],
	}

	device_directory = os.path.join(BASE_DIR_configuration, re.sub(r'[\\@#!|/]', '_', device['Hostname']) + '___' + device['IP'])

	os.mkdir(device_directory)

	with ConnectHandler(**cisco1) as net_connect:
		if device['Enable'] != '':
			net_connect.enable()
		#print(net_connect.find_prompt())
		th_name = threading.current_thread().getName()
		for command in tqdm(cmd_set, th_name, ncols = 100, position = int(position), leave = False):
			output = net_connect.send_command(command.strip())
			with open(os.path.join(device_directory, re.sub(r'[\\@#!|/]', '_', command) + '.txt'), 'wt') as f:
				f.write(output)



if __name__ == '__main__':
	start = time.perf_counter()
	thread_list = [threading.Thread(target=devices_conn, args=(device, device['ID'])) for device in device_infor_all]
		
	for t in thread_list:
		t.start()
	for t in thread_list:
		if t.is_alive():
			t.join()

	#print(threading.active_count())  
	print()
	print("all done: time", time.perf_counter() - start, "\n")

