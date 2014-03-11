#!/usr/bin/python

'''
Author jsr
email  motorunner@yahoo.com
'''

import sys
from sar.parser import *

from tempfile import TemporaryFile
#for writing excel 95 formated files
from xlwt import Workbook


if len(sys.argv) != 2:
    print "usage sperf.py inputsarfile.txt"
    exit()

sarfile = sys.argv[1]
outxlfile = sarfile+'.xls'

sarparser = Parser( sarfile)

if sarparser.load_file() == True:
    print 'Load OK'
else:
    print 'Load FAIL'
    exit()

sar_dict = sarparser.get_sar_info()

#create a xl workbook
xlbook = Workbook()

#Write cpu or any generic header
#for which data_hash will be items for header Eg: { sys,usr,io,st } 
def write_datahead( data_hash, xl_sheet):
    c = 0
    xl_sheet.write( 0, 0, 'time')
    for data_items in data_hash:
        xl_sheet.write( 0, c+1, data_items)
        c = c+1

#Write Data lines for cpu
def write_datalines( data_hash, time, xl_sheet, row ):
    row = row -1
    for items in data_hash[time]:
        c = 1
        if items == 'all':
            for cpu_items in data_hash[time][items]:
                #print row, c, data_hash[time][items][cpu_items]
                xl_sheet.write( row, c, data_hash[time][items][cpu_items])
                c = c+1

#Write io or generic data
def write_data( data_hash, time, xl_sheet, row):
    row = row-1
    data = data_hash[time]
    c=1
    for items in data:
        #print row, c, data[items]
        xl_sheet.write( row, c, data[items])
        c=c+1
        
#Write memory Header
def write_memhead( mem_hash, time):
    data = mem_hash[time]
    c = 0
    for memitems in data:
        #print memitems
        xlmem_sheet.write( 0, c+1, memitems)
        c = c+1

#Create Summary Sheet 
xlsum_sheet = xlbook.add_sheet( 'summary')
filedate = sarparser.get_filedate()
xlsum_sheet.write( 0, 0, sarfile)
xlsum_sheet.write( 1, 0, filedate)

#Iterate throught the keys 
for key in sorted(sar_dict):
    if key == 'mem':
        print "Processing Memory..."
        mem_hash = sar_dict[key]
        xlmem_sheet = xlbook.add_sheet( key)

        firstmem  = mem_hash[mem_hash.keys()[0]]
        write_datahead( firstmem, xlmem_sheet)

        r = 1
        c = 0
        for time in sorted(mem_hash):
            xlmem_sheet.write( r, 0, time)
            r = r+1
            write_data( mem_hash, time, xlmem_sheet, r)
   
    if key == 'cpu':
        print "Processing cpu..."
        cpu_hash = sar_dict[key]
        xlcpu_sheet = xlbook.add_sheet( key)
        #get the first line to find cpu headers 
        firstele = cpu_hash[cpu_hash.keys()[0]]
        firstcpu = firstele[firstele.keys()[0]]
        #write data header
        write_datahead( firstcpu, xlcpu_sheet)

        r=1
        for time in sorted(cpu_hash.keys()):
            #print time, cpu_hash[time]
            xlcpu_sheet.write( r, 0, time)
            r = r+1
            write_datalines( cpu_hash, time, xlcpu_sheet, r)

    if key == 'io':
        print "Processing io..."
        xlio_sheet = xlbook.add_sheet( key)
        io_hash = sar_dict[key]
        firstio = io_hash[io_hash.keys()[0]]
        write_datahead( firstio, xlio_sheet)
        r = 1
        for time in sorted(io_hash.keys()):
            #print time, io_hash[time]
            xlio_sheet.write( r, 0, time)
            r = r+1
            write_data( io_hash, time, xlio_sheet, r)

    if key == 'swap':
        print "Processing swap..."
        swap_hash = sar_dict[key]
        xlswap_sheet = xlbook.add_sheet( key)
        fswap = swap_hash[swap_hash.keys()[0]]
        write_datahead( fswap, xlswap_sheet)
        r = 1
        for time in sorted(swap_hash.keys()):
            xlswap_sheet.write( r, 0, time)
            r = r+1
            write_data( swap_hash, time, xlswap_sheet, r)

    if key == 'prcsw':
        print "Processing proc cswch..."
        prcsw_hash = sar_dict[key]
        xlprcsw_sheet = xlbook.add_sheet( key)
        fprcsw = prcsw_hash[prcsw_hash.keys()[0]]
        write_datahead( fprcsw, xlprcsw_sheet)
        r = 1
        for time in sorted(prcsw_hash.keys()):
            xlprcsw_sheet.write( r, 0, time)
            r = r+1
            write_data( prcsw_hash, time, xlprcsw_sheet, r)

    if key == 'page':
        print "Processing pageing..."
        page_hash = sar_dict[key]
        print page_hash
        
        
xlbook.save(outxlfile)
xlbook.save(TemporaryFile())
