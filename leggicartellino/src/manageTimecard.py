#! /usr/bin/env python
# -*- coding: utf-8 -*-
# manageTimecard.py
#
# A simple tool to calculate hours and minutes from a time card


import calendar
import csv
import datetime

import os
from os import listdir

import sys
import xlrd
import re


__author__ = "Riccardo Del Gratta <riccardo.delgratta@ilc.cnr.it>"
__date__ = "$12 Dec, 2019 12:30 PM$"

# some utilities functions that I write here but they should be moved in a different file :) 

def get_month(month):
    month = datetime.date(1900, int(month), 1).strftime('%B')
    return month

def get_day_of_week(year, month, day):
    day = datetime.date(year, month, day)
    wod = calendar.day_name[day.weekday()]
    return wod[0:2]


def pivot_worked_time_single_file(indata):
    nums = list()
    dates = list()
    hours = list()
    ret = list()
    all=list()
    rcounter = 0
    ore = 0
    str_delimiter = '\t'
    check = 'Differenza orario standard'
    
    id = 1;
    with open(indata, mode='r') as tsv:
        # parse file name in search of Name, Year and Month

        for line in csv.reader(tsv, delimiter=str_delimiter):
            #print "CCC ", line[0], rcounter, id
           
            ore = 0
            
            if line[0] <> check:
                if (len(line[0]) > 0 and rcounter > 2): # si inizia dalla terza linea
                    mydate = line[0]
                    ret = parse_date(str(mydate))
                    wod=get_day_of_week(ret[0], ret[1], ret[2])
                    
                    if (len(line[4]) > 0): # si inizia dalla terza linea
                        ore = line[4]
                        hours.append(ore)
                    else:
                        hours.append('')

                    dates.append(wod)
                    nums.append(str(id))
                    id = id + 1        
            else:
                #print 'Uscita per',check,'a linea',rcounter+1, 'totale ore e minuti: ',totaleore, totalemin, totalemin/60
                #print 'Uscita per diff ',check, dates, nums
                #print str(str_delimiter.join(nums))
                #print str_delimiter.join(dates)
                all.append(hours)
                all.append(nums)
                all.append(dates)
                return all
            rcounter = rcounter + 1    
            

def csv_from_excel_single_file(indata):
    sheet='Cartellino'
    str_delimiter='\t'
    wb = xlrd.open_workbook(indata)
    sh = wb.sheet_by_name(sheet)
    i=indata.rfind('.')
    csv_filename=indata[0:i+1]+"csv"
    name=str(sh.row_values(0)[5])
    name=re.sub(r"\s+", '_', name)
    parse_file_name(indata)
    
    csv_file = open(csv_filename, 'w')
    wr = csv.writer(csv_file, delimiter=str_delimiter,quotechar='"', quoting=csv.QUOTE_MINIMAL)

    for rownum in range(sh.nrows):
        print 'line ', sh.row_values(rownum), rownum
        wr.writerow(sh.row_values(rownum))

    csv_file.close()
   

def csv_from_excel_file(flag, indata):
    # a list containing files
    files = list()
    str_delimiter = '\t'
    ret = list()
    # first of all is a file or a folder
    #print mode
    if '-f' in flag:
        #print 'folder'
        for f in listdir(indata):
            if f.endswith('.xls'):
                #print indata+"/"+f
                files.append(indata + "/" + f)
    elif '-i' in flag:
        if indata.endswith('.xls'):
            #print indata
            files.append(indata)
            
   
    for indata in files:
        print '\t\tconverting  ', indata, '\n'
        
        csv_from_excel_single_file(indata)
        
        #print'month year ',month, year
        
        
            

'''
sums up hours for a single file
@indata the single file
'''
def sum_worked_time_single_file(indata):
    rcounter = 0
    h1 = 0
    m1 = 0
    h2 = 0
    m2 = 0
    ore = 0
    diff = 0
    idxsep = 0
    idxmeno = 0
    meno = '-'
    sep = ':'
    str_delimiter = '\t'
    moltiplicatore = 1
    totaleore = 0
    totaleorediff = 0
    totalemin = 0
    totalemindiff = 0
    totaleorestandard = 0
    
    totaleminstandard = 0
    check = 'Differenza orario standard'
    ret = list()
    
    with open(indata, mode='r') as tsv:
        # parse file name in search of Name, Year and Month

        for line in csv.reader(tsv, delimiter=str_delimiter):

            h1 = 0
            m1 = 0
            h2 = 0
            m2 = 0
            ore = 0
            diff = 0
            idxsep = 0
            idxmeno = 0
            if line[0] <> check:
                if (len(line[4]) > 0 and rcounter > 2): # si inizia dalla terza linea
                    ore = line[4]
                    #cerco i :
                    idxsep = ore.find(sep)

                    h1 = ore[0:idxsep]
                    m1 = ore[idxsep + 1:len(ore)]
                    diff = line[5]
                    # cerco la stringa meno in diff
                    idxmeno = diff.find(meno)
                    if idxmeno <> -1:
                        moltiplicatore = -1
                        diff = diff[1:]
                        #print "diff1 ",diff,moltiplicatore
                    else:
                        moltiplicatore = 1

                    idxsep = diff.find(sep)
                    h2 = diff[0:idxsep]
                    m2 = diff[idxsep + 1:len(diff)]
                    #print "H", h2
                    if int(h2) > 0:
                        h2 = int(h2) * moltiplicatore
                        m2 = int(m2) * moltiplicatore
                        
                    else:
                        #print "A", h2, m2
                        m2 = int(m2) * moltiplicatore


                # print line[4], line[5],rcounter
                    #print "ore: ",int(h1), "minuti ",int(m1),"diff ore ",int(h2), "diff minuti", int(m2)," alla linea ",rcounter
                    totaleore = totaleore + int(h1)
                    totalemin = totalemin + int(m1)
                    totaleorediff = totaleorediff + int(h2)
                    totalemindiff = totalemindiff + int(m2)
                    #print "XXXX ", totaleorediff, totalemindiff
            else:
                #print 'Uscita per',check,'a linea',rcounter+1, 'totale ore e minuti: ',totaleore, totalemin, totalemin/60
                totaleore = totaleore + totalemin / 60
                totalemin = totalemin % 60
                totaleorediff = totaleorediff + totalemindiff / 60
                totalemindiff = totalemindiff % 60
                #totaleorediff and totalemindiff could be negative
                temp = totaleorediff * 60 + totalemindiff
                #print temp
                totaleorediff = temp / 60
                totalemindiff = temp % 60
                #print temp, temp/60, temp%60
                #print 'Uscita per diff ',check,'a linea',rcounter+1, 'totale ore e minuti diff : ',totaleorediff , totalemindiff #, totalemin/60
                #print 'Totale ',totaleore,':',totalemin
                
                #Ore Totali	Minuti Totali	Ore Differenza	Minuti Differenza	Ore Standard	Minuti Standard	SixtyBased	DecBased
                totaleorestandard = ((totaleore * 60 + totalemin)-(totaleorediff * 60 + totalemindiff)) / 60
                totaleminstandard = ((totaleore * 60 + totalemin)-(totaleorediff * 60 + totalemindiff)) % 60
                
                sixtybased = round(float(str(totaleore) + "." + str(totalemin)), 2)
                tempdec = int(round(float(totalemin) / 60, 2) * 100)
                decbased = round(float(str(totaleore) + "." + str(tempdec)), 2)
              
                ret.append(str(totaleore))
                ret.append(str(totalemin))
                ret.append(str(totaleorediff))
                ret.append(str(totalemindiff))
                ret.append(str(totaleorestandard))
                ret.append(str(totaleminstandard))
                ret.append(str(sixtybased))
                ret.append(str(decbased))

                return ret
            rcounter = rcounter + 1
'''
parses the date 
@mydate the date to parse
return a list with year, month, day
'''
def parse_date(mydate):
    #print 'BBB ',mydate #.strftime('%B')
    ret = list()
    i = mydate.rfind('/')
    year = mydate[i + 1:]
    mydate = mydate[0:i]
    i = mydate.rfind('/')
    month = mydate[i + 1:]
    mydate = mydate[0:i]
    i = mydate.rfind('/')
    day = mydate[i + 1:]
    
    ret.append(int(year))
    ret.append(int(month))
    ret.append(int(day))
   
    return ret            

'''
parses the file name to return name, year and month
'''
def parse_file_name(oldindata):
    ret = list()
    i = oldindata.rfind('/')
    oldindata = oldindata[i + 1:]
    i = oldindata.rfind('_')
    month = oldindata[i + 1:i + 3]
    oldindata = oldindata[0:i]
    i = oldindata.rfind('_')
    year = oldindata[i + 1:i + 5]
    name = oldindata[0:i]
    ret.append(name)
    ret.append(year)
    ret.append(get_month(month))
    return ret

'''
sums up the hours in the timecard
@flag -i for single file and -f for folder
@indata according to flag: file or folder
@outfile output file or standard output
@mode 1 for standard output 0 for file
'''
def sum_worked_time(flag, indata, outfile, mode):
    # a list containing files
    files = list()
    str_delimiter = '\t'
    ret = list()
    # first of all is a file or a folder
    #print mode
    if '-f' in flag:
        #print 'folder'
        for f in listdir(indata):
            if f.endswith('.csv'):
                #print indata+"/"+f
                files.append(indata + "/" + f)
    elif '-i' in flag:
        if indata.endswith('.csv'):
            #print indata
            files.append(indata)
            
    # well create the file in append
    # first of all, create the columns
    columns = ['Name', 'Year', 'Month', 'Total Hours', 'Total Minutes', 'Difference Hours', 'Difference Minutes', 'Due Hours', 'Due Minutes', 'SixtyBased', 'DecBased']
    	

    if mode == 1:
        sys.stdout.write(str_delimiter.join(columns) + '\n')
    elif mode == 0:    
        if os.path.isfile(outfile):
            os.remove(outfile)
            
        ofile = open(outfile, mode='a')
        ofile_writer = csv.writer(ofile, delimiter=str_delimiter, quotechar='"', quoting=csv.QUOTE_MINIMAL)
        ofile_writer.writerow(columns)
    
    
            
    # in files there is at least one file
    for indata in files:
        print '\t\tsumming up ', indata, '\n'
        ret = list()
        ret1 = list()
        ret2 = list()
        ret2 = parse_file_name(indata)
        
        #print'month year ',month, year
        
        ret1 = sum_worked_time_single_file(indata)
        ret.append(ret2[0]) #name
        ret.append(ret2[1]) #year
        ret.append(ret2[2]) #month
        for r in ret1:
            ret.append(r)
        
        #print ret
        if mode == 1:
            sys.stdout.write(str_delimiter.join(ret) + '\n')
        elif mode == 0: 
            ofile_writer.writerow(ret)
        
        
'''
pivots the timecard n horizontal way
adds weekdays as well
@flag -i for single file and -f for folder
@indata according to flag: file or folder
@outfile output file or standard output
@mode 1 for standard output 0 for file
'''
    
def pivot(flag, indata, outfile, mode):
    # a list containing files
    files = list()
    str_delimiter = '\t'
    # first of all is a file or a folder
    #print mode
    if '-f' in flag:
        #print 'folder'
        for f in listdir(indata):
            if f.endswith('.csv'):
                #print indata+"/"+f
                files.append(indata + "/" + f)
    elif '-i' in flag:
        if indata.endswith('.csv'):
            #print indata
            files.append(indata)
            
    # well create the file in append
    # first of all, create the columns
    columns = ['Name', 'Year', 'Month', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23'
    , '24', '25', '26', '27', '28', '29', '30', '31']
    	
#    with open(outfile, mode='a') as ofile:
#        ofile_writer = csv.writer(ofile, delimiter=str_delimiter, quotechar='"', quoting=csv.QUOTE_MINIMAL)
#
#        ofile_writer.writerow(columns)
    if mode == 1:
        sys.stdout.write(str_delimiter.join(columns) + '\n')
    elif mode == 0:    
        if os.path.isfile(outfile):
            os.remove(outfile)
            
        ofile = open(outfile, mode='a')
        ofile_writer = csv.writer(ofile, delimiter=str_delimiter, quotechar='"', quoting=csv.QUOTE_MINIMAL)
        ofile_writer.writerow(columns)    
            
    # in files there is at least one file
    for indata in files:

        ret1 = list()
        ret2 = list()
        ret2 = parse_file_name(indata)
        
        #print'month year ',month, year
        print '\t\tpivoting ', indata,'\n' 
        ret1 = pivot_worked_time_single_file(indata)
        
        for r1 in ret1:
            r1.insert(0,ret2[0]) #name
            r1.insert(1,ret2[1]) #year
            r1.insert(2,ret2[2]) #month
 
            if mode == 1:
                sys.stdout.write(str_delimiter.join(r1) + '\n')
            elif mode == 0: 
                ofile_writer.writerow(r1)        
            
'''
prints the help
'''
def print_help(name):
    print 'Usage: python ' + name + ' -a <attività> -<i|f> <singlefile|folder> [-o opt_output_file] '
    print '\n'
    print 'Please note that'
    print '\t1) -i anf -f flags are mutually exclusive:'
    print '\t\t use -i if you provide data using a file'
    print '\t\t use -f if you provide data using a folder containing more than one file'  
    print '\t\t **** It is allowed, however to use -f with a folder containing a single file ****'
    print '\n'
    print '\t2) inputfile(s) MUST be in csv format with tab separator and with the following name convention:'
    print '\t\t **** Name_Year_Month.csv ****'
    print '\t\t **** MUST contain NO spaces ****'
    print '\t\t **** Year is like YYYY (2019, 2018... ) ****'
    print '\t\t **** Month is like MM (01, 02, 11... ) ****'
    print '\n'
    print '\t3) -o is an optional parameter which identifies the output file.'
    print '\t\t -o MUST be provided with the name of the outfile to write'
    print '\t\t do not use -o flag if you want to see the results at standard output (display'
    print '\n'
    print '\t4)-a is the activity to perform (so far only sum and pivot)'
    print '\t\t use sum it you want to compute the total amount of time <Name> worked in <Year> and <Month>'
    print '\t\t use pivot it you want to pivot the timecard in horizontal lines: it is useful for the managements'
    print'\n'
    print'The following are valid examples'
    print '\tEx1: python ' + name + ' -a sum -i Myname_2018_07.csv'
    print '\tEx2: python ' + name + ' -a pivot -i Myname_2018_07.csv -o myoutfile'
    print '\tEx3 python ' + name + ' -a sum -f myfolder'
    print '\tEx4: python ' + name + ' -a pivot -f myfolder -o myoutfile'
    
'''
parse arguments and sets some values
@args the list of arguments passed to main (including the name of the program)
'''
            
def parse_args(args):
    printstd = 1
    mylen = 4
    ret = list()
    name = args[0]
    args = args[1:]
    for a in args:
        if '-o' in a:
            printstd = 0
            mylen = 6 
            break
    if len(args) != mylen:
        printHelp(name)
        sys.exit(-1)
    else:    
        act = args[1];
        file_flag = args[2]
        input_data = args[3]
        if mylen == 6:
            out_file = args[5]
        else:
            out_file = 'stdout'
    
    ret.append(act)
    ret.append(file_flag)
    ret.append(input_data)
    ret.append(out_file)
    ret.append(printstd)
    #print "XXX ",printstd
    return ret
    
'''
main function
'''
def main():
    
#    print 'Number of arguments:', len(sys.argv), 'arguments.'
#    print 'Argument List:', str(sys.argv)
    
    ret = parse_args(sys.argv)
    
    #get values from ret
    act = ret[0]
    file_flag = ret[1]
    input_data = ret[2]
    out_file = ret[3]
    mode = ret[4]
    
    #print ret
    
    if out_file == 'stdout':
        outfile = sys.stdout
    else:
        outfile = out_file
   
    print '\n---------------------------'
    print "\t Starting Activity  " + act + "\n "
    if (act.lower()) == "sum":
        sum_worked_time(file_flag, input_data, outfile, mode)
    elif (act.lower()) == "pivot":
        pivot(file_flag, input_data, outfile, mode) #(args[1], outfile,args[2], int(args[3]), int(args[4]), int(args[5]),int(args[6]))
    elif (act.lower()) == "convert":
        csv_from_excel_file(file_flag, input_data)
         #(args[1], outfile,args[2], int(args[3]), int(args[4]), int(args[5]),int(args[6]))
    else:
        print "Activity not in list: [somma|pivot|convert]"
        
    print "\t Activity  " + act + " Finished"
    if mode==0:
        print "\t output  in file " + outfile 
    elif mode==1:
        print "\t output in standard output"
    print '---------------------------'
   
 
if __name__ == '__main__':
    main()  