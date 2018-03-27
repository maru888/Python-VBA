# coding: utf-8
import xlwings as xw
import sys
import time

def run():

	num_data = int(xw.Range("B2").value)

	for i in range(5, num_data + 5):        
		xw.Range("B" + str(i)).value = i - 4
		time.sleep(1)

