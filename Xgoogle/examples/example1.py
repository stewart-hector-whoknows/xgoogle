#!/usr/bin/python
#
# This program does a Google search for "quick and dirty" and returns
# 50 results. :)))
#

from xlrd import open_workbook
from xgoogle.search import GoogleSearch, SearchError


def googleFind(company, file):
    try:
        gs = GoogleSearch(company)
        gs.results_per_page = 10
        results = gs.get_results()
        for res in results:
            print res.title.encode('utf8')
            fh.write("<P>" + res.title.encode('utf8') + "</P>")         
            print res.desc.encode('utf8')
            fh.write("<P>" + res.desc.encode('utf8') + "</P>")                   
            print res.url.encode('utf8')
            fh.write('<a href="' + res.url.encode('utf8')+  '">' + res.title.encode('utf8') + '</a>')
      
            
            print
    except SearchError, e:
        print "Search failed: %s" % e

header = """
<HTML>
<HEAD>
<TITLE>My Web Page</TITLE>
</HEAD>
<BODY>
"""


footer = """

</BODY>
</HTML>
"""

fh = open("summary.html", "w")
fh.write(header)


wb = open_workbook('guangzhou_2013.xls')

#googleFind("test", fh)
for s in wb.sheets():
    print 'Sheet:',s.name
    for row in range(10): #s.nrows):
        #values = []
        #values.append(s.cell(row,1).value)
        val=  s.cell(row,1).value
        fh.write("<b>" + val + "</b>")
        googleFind(val, fh)
        fh.write("</BR>")     
        fh.write("</BR>")   
        fh.write("</BR>")      
        fh.write("</BR>")     
        fh.write("</BR>")   
        fh.write("</BR>")             
        #print values
        print s.cell(row,1).value


        
fh.write(footer)


