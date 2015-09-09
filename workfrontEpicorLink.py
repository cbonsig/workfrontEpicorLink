'''
workfrontEpicorLink.py
9 September 2015
cbonsig

home: 		some.linux.server:/var/python/workfront/[name of this version].py
			symlink to -> /var/python/workfront/workfrontEpicorLink.py
			sudo ln -s [name of ths version].py /var/python/workfront/workfrontEpicorLink.py
usage: 		python workfrontEpicorLink.py
daemon: 	sudo service workfront [start, stop] (see below)
log:		/var/logs/workfront.log

runs on some.linux.server, and listens on port 8071. responds to requests 
triggered by a Workfront "External Page" dashboard placed in a task detail view.
returns quotes, orders, jobs, and invoices related to the task, using the jiraCode (custom field)
or referenceNumber to match the task to Epicor entries.

syntax definition in Workfront:
Reporting > Dashboards > (name, i.e. "Epicor Task Integration") > Edit > + Add External Page >
Edit > URL = http://some.linux.server:8071/?type=[TASK,ISSUE,PROJ]&session={!$$SESSION}&object={!ID}
* For now, must have a different dashboard for each type - i.e. one for TASK, another for PROJ

this application waits for requests to arrive, then
1. parses the task object ID and session object from the GET request 
2. makes a connection to the MSSQLSERVER Epicor905 database using read-only access
3. calls the API and requests the referenceNumber and JIRA code related to the task object
4. queries the Epicor905 database to retreive quote, order, job, and invoice data
5. responds with HTML formatted tables to display data
6. formats with bootstrap CSS, served by some.linux.server and stored at /var/www/css, .../js, .../fonts

note:
in Chrome, and possibly other browsers, the embedded HTML in the Workfront dashboard is considered an "unsafe script"
and it prevented from loading. The workaround is to click the gray shielf icon at the right of the URL bar,
and click the button to "allow unsafe scripts". A better solution would be to revise this to serve HTTPS and 
configure a valid security certificate on the server.

references:
1. http://store.atappstore.com/index.php/executive-summary/
2. http://pymssql.org/en/stable/_mssql_examples.html
3. http://pymotw.com/2/BaseHTTPServer/
4. https://mkaz.com/2012/10/10/python-string-format/
5. https://developers.attask.com/api-docs/
6. https://developers.attask.com/wp-content/themes/revan/resources/downloads/python.zip
7. http://getbootstrap.com
8. http://www.tutorialrepublic.com/twitter-bootstrap-tutorial/bootstrap-tables.php
9. http://www.tutorialspark.com/twitterBootstrap/TwitterBootstrap_Collapsible_Accordion_Demo.php

future:
* http://www.cherrypy.org/
* http://www.zacwitte.com/using-ssl-https-with-cherrypy-3-2-0-example
* maybe add section for "Releases" (backlog)

service script
------------------------------------------------------
# workfront
description     "start workfront python http server."

start on startup
stop on shutdown 

console log

script
    exec python /var/python/workfront/workfrontEpicorLink.py >> /var/log/workfront.log 2>&1
end script
------------------------------------------------------

Revision notes
9 September 2015: post generic version of code to github
6 March 2015: Extended to work with ISSUE or TASK. Added graceful fail if in creation dashboard, or not triggered from issue or task.
9 March 2015: Extend to Project, using EpicorProjectCode custom field. Add collapsible panels for each table.
10 March 2015: Customize intro text at top for PROJ vs TASK. Filter Jobs for released=1.
11 March 2015: Added PN/Desc to Quote, Order, Jobs results (invoice SQL is from header, not detail ... so can't get to PN/Desc without rewriting)
12 March 2015: Revised SQL queries per Lance advice re: matching on Company
17 March 2015: Revised admin login credentials
18 March 2015: Resolved problem with rendering unicode characters in Part / job.

'''

from BaseHTTPServer import HTTPServer, BaseHTTPRequestHandler
from SocketServer import ThreadingMixIn
import threading
import urlparse
from api import StreamClient, ObjCode, AtTaskObject
import _mssql
import sys

reload(sys)  
sys.setdefaultencoding('utf8')

serverPort = 8071

class GetHandler(BaseHTTPRequestHandler):
	
	def do_GET(self):

		def do_FAIL():
			self.send_response(200)
			self.send_header("Content-type", "text/html")	
			self.end_headers()
			self.wfile.write("<!DOCTYPE html>\n")
			self.wfile.write("<html>\n")
			self.wfile.write("<head>\n")
			self.wfile.write("\t<meta>\n")
			self.wfile.write("\t<title>Fail</title>\n")
			self.wfile.write("\t<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n")
			self.wfile.write("\t<link rel=\"stylesheet\" type=\"text/css\" href=\"http://some.linux.server/css/bootstrap.min.css\">\n")
			self.wfile.write("</head>\n")
			self.wfile.write("<body>\n")
			self.wfile.write("\t<script src=\"http://code.jquery.com/jquery.min.js\"></script>\n")
			self.wfile.write("\t<script src=\"http://some.linux.server/js/bootstrap.min.js\"></script>\n")
			self.wfile.write("\n\t<div class=\"container\">\n")
			self.wfile.write("\n\t<h1>Sorry :(</h1>\n")
			self.wfile.write("\n\tThere's nothing to see here.\n")
			self.wfile.write("</body>\n")
			self.wfile.write("</html>\n")
			return

		# log the thread id
		message =  threading.currentThread().getName()
		
		# parse the URL
		parsed_path = urlparse.urlparse(self.path)	
		params = urlparse.parse_qs(parsed_path.query)
#		print params

		# authenticate to Workfront API
		client = StreamClient('https://DOMAIN.attask-ondemand.com/attask/api')
		client.login('ADMINUSERNAME','PASSWORD')
		
		# get the object type from the parameters
		objectTypeParam = params['type']
#		print objectTypeParam[0]

		thisObjectCode =  params['object']

		if (objectTypeParam[0] == 'TASK'):
			try:
				thisObject = AtTaskObject(client.get(ObjCode.TASK,thisObjectCode[0],{'referenceNumber','DE:JIRA'}))
				keyFieldQuote = 'qd.Character02' # QuoteDtl.Character02 is used for JIRA Link (Workfront Link)
				keyFieldOrder = 'od.Character02' # OrderDtl.Character02 is used for JIRA Link (Workfront Link)
				keyFieldJob = 'od.Character02' # JobHead.Character02 is used for JIRA Link (Workfront Link)
				keyFieldInvoice = 'od.Character02' # JobHead.Character02 is used for JIRA Link (Workfront Link)
			except:
				do_FAIL()
				return

		elif (objectTypeParam[0] == 'ISSUE'):
			try:
				thisObject = AtTaskObject(client.get(ObjCode.ISSUE,thisObjectCode[0],{'referenceNumber','DE:JIRA'}))
				keyFieldQuote = 'qd.Character02' # QuoteDtl.Character02 is used for JIRA Link (Workfront Link)
				keyFieldOrder = 'od.Character02' # OrderDtl.Character02 is used for JIRA Link (Workfront Link)
				keyFieldJob = 'od.Character02' # OrderDtl.Character02 is used for JIRA Link (Workfront Link)
				keyFieldInvoice = 'od.Character02' # OrderDtl.Character02 is used for JIRA Link (Workfront Link)
			except:
				do_FAIL()
				return

		elif (objectTypeParam[0] == 'PROJ'):
			try:
				thisObject = AtTaskObject(client.get(ObjCode.PROJECT,thisObjectCode[0],{'referenceNumber','DE:JIRA','DE:Epicor Code'}))
				keyFieldQuote = 'qh.ShortChar05' # QuoteHed.ShortChar05 is used for Project Code
				keyFieldOrder = 'oh.ShortChar05' # OrderHed.ShortChar05 is used for Project Code
				keyFieldJob = 'jh.ShortChar05' # JobHead.ShortChar05 is used for ProjectCode
				keyFieldInvoice = 'oh.ShortChar05' # OrderHed.ShortChar05 is used for ProjectCode
			except:
				do_FAIL()
				return

		else:
			do_FAIL()
			return

		# get the object code from the URL parser, and create an AtTask object for the related task
		
		
#		print thisObject

		# get the object type, referenceNumber, and JIRA code for the task
		objectType = thisObject.data.get('objCode')
		refNum = thisObject.referenceNumber
		jiraCode = thisObject.data.get('DE:JIRA','none')
		projectCode = thisObject.data.get('DE:Epicor Code','none')

#		print objectType
#		print refNum
#		print jiraCode
#		print projectCode

		# manage Epicor link transition from JIRA code to Workfront referenceNum
		# if the JIRA custom field is present, use its contents as the idCode
		# otherwise, use the referenceNumber
		if '-' in jiraCode:
			idCode = jiraCode
		else:
			idCode = refNum

		# if we're looking for a project, search on the Project Code, not the JIRA code / Workfront ID
		if (objectType == 'PROJ'):
			idCode = projectCode

		# if there is no project code defined, idCode will be set to 'none', so trap for that
		if (idCode == 'none'):
			do_FAIL()
			return

		# connect to MS SQL server
		conn = _mssql.connect(server='MSSQLSERVER.domain.tld',user='READONLYUSER',password='PASSWORD',database='Epicor905')

		# string containing SQL query for Quote details
		quoteSQL = """
		SELECT
		  LEFT(qh.ShortChar05, 5) AS ProjID,
		  qd.ProdCode,
		  qh.DateQuoted,
		  (CAST(qd.QuoteNum AS varchar) + ' / ' + CAST(qd.QuoteLine AS varchar)) AS QuoteLine,
		  (CAST(qd.SellingExpectedQty AS int)) AS Qty,
		  (CAST(qd.DocExtPriceDtl AS money)) AS LineCharges,
		  (CAST(ISNULL(aggregateMisc.MiscTotal, 0) AS money)) AS MiscCharges,
		  (CAST((ISNULL(SUM(aggregateMisc.MiscTotal), 0) + qd.DocExtPriceDtl) AS money)) AS Total,
		  (CAST(qd.PartNum AS varchar) + ' / ' + LEFT(CAST(qd.LineDesc AS varchar),50) ) AS PNDesc
		FROM Epicor905.dbo.QuoteDtl AS qd
		LEFT OUTER JOIN Epicor905.dbo.QuoteHed AS qh
		  ON qd.QuoteNum = qh.QuoteNum
		LEFT OUTER JOIN (SELECT
		  SUM(qm.DocMiscAmt) AS MiscTotal,
		  QuoteNum,
		  QuoteLine
		FROM Epicor905.dbo.QuoteMsc AS qm
		GROUP BY QuoteNum,
		         QuoteLine) AS aggregateMisc
		  ON (qd.QuoteNum = aggregateMisc.QuoteNum)
		  AND (qd.QuoteLine = aggregateMisc.QuoteLine)
		WHERE (%s LIKE \'%s%%\')
		GROUP BY qd.QuoteNum,
		         qd.QuoteLine,
		         qd.SellingExpectedQty,
		         qd.DocExtPriceDtl,
		         aggregateMisc.MiscTotal,
		         qh.ShortChar05,
		         qd.ProdCode,
		         qh.DateQuoted,
		         qd.PartNum,
		         qd.LineDesc
		ORDER BY qd.QuoteNum DESC, qd.QuoteLine ASC
		""" % (keyFieldQuote, idCode)

		# string containing SQL query for Order details
		orderSQL ="""
		SELECT
		  oh.OrderDate,
		  (CAST(od.OrderNum AS varchar) + ' / ' + CAST(od.OrderLine AS varchar) + ' / ' + oh.PONum) AS OrderLinePO,
		  (CAST(od.OrderQty AS int)) AS Qty,
		  (CAST(od.DocExtPriceDtl AS money)) AS LineCharges,
		  (CAST(ISNULL(aggregateMisc.MiscTotal, 0) AS money)) AS MiscCharges,
		  (CAST((ISNULL(SUM(aggregateMisc.MiscTotal), 0) + od.DocExtPriceDtl) AS money)) AS Total,
		  (CAST(od.PartNum AS varchar) + ' / ' + LEFT(CAST(od.LineDesc AS varchar),50) ) AS PNDesc
		FROM Epicor905.dbo.OrderDtl AS od
		LEFT OUTER JOIN (SELECT
		  SUM(om.DocMiscAmt) AS MiscTotal,
		  OrderNum,
		  OrderLine
		FROM Epicor905.dbo.OrderMsc AS om
		GROUP BY OrderNum,
		         OrderLine) AS aggregateMisc
		  ON (od.OrderNum = aggregateMisc.OrderNum)
		  AND (od.OrderLine = aggregateMisc.OrderLine)
		INNER JOIN OrderHed AS oh
		  ON oh.Company = od.Company
		  AND oh.OrderNum = od.OrderNum
		WHERE (%s LIKE \'%s%%\')
		GROUP BY od.OrderNum,
		         oh.PONum,
		         od.OrderLine,
		         od.OrderQty,
		         od.DocExtPriceDtl,
		         aggregateMisc.MiscTotal,
		         oh.OrderDate,
		         od.PartNum,
		         od.LineDesc
		ORDER BY od.OrderNum DESC, od.OrderLine ASC
		""" % (keyFieldOrder, idCode)

		# string containing SQL query for Job details
		jobSQL = """
		SELECT
		  jh.CreateDate,
		  od.OrderNum,
		  jh.JobNum,
		  jh.ProdQty,
		  jh.QtyCompleted,
		  jh.DueDate,
		  jh.JobCompletionDate,
		  (CAST(od.PartNum AS varchar) + ' / ' + LEFT(CAST(od.LineDesc AS varchar),50) ) AS PNDesc
		FROM Epicor905.dbo.OrderDtl AS od
		INNER JOIN Epicor905.dbo.JobProd AS jp
		  ON od.Company = jp.Company
		  AND od.OrderNum = jp.OrderNum
		  AND od.OrderLine = jp.OrderLine
		INNER JOIN Epicor905.dbo.JobHead AS jh
		  ON jp.Company = jh.Company
		  AND jp.JobNum = jh.JobNum
		WHERE UPPER(%s) LIKE \'%s\'
		  AND jh.JobReleased = 1
		ORDER BY jh.CreateDate DESC
		""" % (keyFieldJob, idCode)

		# string containing SQL query for Invoice details
		invoiceSQL = """
		SELECT
		  InvoiceNum,
		  InvoiceDate,
		  DocInvoiceAmt,
		  OpenInvoice,
		  OrderNum,
		  PONum
		FROM Epicor905.dbo.InvcHead ih
		WHERE ih.OrderNum IN (SELECT DISTINCT
		  od.ORDERNUM
		FROM Epicor905.dbo.OrderDtl od
		LEFT OUTER JOIN Epicor905.dbo.OrderHed oh
		  ON od.Company = oh.Company
		  AND od.OrderNum = oh.OrderNum
		WHERE %s LIKE \'%s\')
		ORDER BY ih.InvoiceDate DESC
		""" % (keyFieldInvoice, idCode)

		# send HTTP 200 OK status code, headers
		self.send_response(200)
		self.send_header("Content-type", "text/html")	
		self.end_headers()

		# write HTML top matter
		self.wfile.write("<!DOCTYPE html>\n")
		self.wfile.write("<html>\n")
		self.wfile.write("<head>\n")
		self.wfile.write("\t<meta>\n")
		self.wfile.write("\t<title>Epicor Integration</title>\n")
		self.wfile.write("\t<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n")
		self.wfile.write("\t<link rel=\"stylesheet\" type=\"text/css\" href=\"http://some.linux.server/css/bootstrap.min.css\">\n")
		self.wfile.write("</head>\n")
		self.wfile.write("<body>\n")
		self.wfile.write("\t<script src=\"http://code.jquery.com/jquery.min.js\"></script>\n")
		self.wfile.write("\t<script src=\"http://some.linux.server/js/bootstrap.min.js\"></script>\n")

		self.wfile.write("\n\t<div class=\"container\">\n")

		if (objectType == 'PROJ'):
			self.wfile.write("\n\t<h2>Epicor Reports for Project {0}</h2>\n".format(idCode))
			self.wfile.write("\n\t<p>In Workfront, the Project Code is selected in \
				Project Details > Standard Project Form > Epicor Code. In Epicor, the Project code is \
				selected from the dropdown menu on Epicor Quotes, Orders, and Jobs. This report shows Epicor\
				details matching the selected Project Code. Click on each panel to display details.</p>\n")
		else:
			self.wfile.write("\n\t<h2>Epicor Reports for Task {0}</h2>\n".format(idCode))
			self.wfile.write("\n\t<p>This report provides details related to the Task ID that is entered in the \
				\"JIRA Link\" field on Epicor quotes and orders. For old entries, the Task ID is the \
				JIRA code (\"PROTO-123\"). For new entries, the Task ID is the Workfront Task Reference Number \
				(\"12345\"). Click on each panel to hide details.</p>\n")

		self.wfile.write("\n\t\t<div class=\"panel-group\" id=\"accordian\">\n")

		# execute SQL query for quote detail
		conn.execute_query(quoteSQL)

		# create an empty list for rows of strings to be joined later
		tempList = []

		# create table for quote details
		self.wfile.write("\n\t\t\t<div class=\"panel panel-primary\">\n")
		self.wfile.write("\n\t\t\t\t<div class=\"panel panel-heading\">\n")
		self.wfile.write("\t\t\t\t\t<h3 class=\"panel-title\">\n")
		self.wfile.write("\t\t\t\t\t\t<a data-toggle=\"collapse\" data-parent=\"#accordion\" href=\"#accordionOne\">\n")
		self.wfile.write("\t\t\t\t\t\t\tQuote Details\n")
		self.wfile.write("\t\t\t\t\t\t</a>\n")
		self.wfile.write("\t\t\t\t\t</h3>\n")
		self.wfile.write("\t\t\t\t</div> <!-- close panel-heading -->\n")
		# for PROJ reports, there's usually lots of data, so start with the panels collapsed
		# for TASK/ISSUE reports, there is usually less data, so start with the panels expanded
		if (objectType == 'PROJ'):
			self.wfile.write("\n\t\t\t\t<div id=\"accordionOne\" class=\"panel-collapse collapse\">\n")
		else:
			self.wfile.write("\n\t\t\t\t<div id=\"accordionOne\" class=\"panel-collapse collapse in\">\n")
		self.wfile.write("\t\t\t\t\t<div class=\"panel-body\">\n")
		self.wfile.write("\t\t\t\t\t\t<table class=\"table table-striped\">\n")
		tempList.append('\t\t\t\t\t\t\t<tr><thead>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th>Date Quoted</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th>Quote / Line</th>\n')
#		tempList.append('\t\t\t\t\t\t\t\t<th>Project Code</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th>PN / Description</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th>Product Group</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th class=\"text-right\">Quantity</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th class=\"text-right\">Line Charges</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th class=\"text-right\">Misc Charges</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th class=\"text-right\">Total Value</th>\n')
		tempList.append('\t\t\t\t\t\t\t</tr></thead>\n')
		for row in conn:
			tempList.append('\t\t\t\t\t\t\t\t<tr>\n')
			try:
				dateString = (row['DateQuoted']).strftime('%d %b %Y')
			except:
				# dateString = str(None) did this cause a problem?
				dateString = 'None'
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-left\">', dateString, '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td>', row['QuoteLine'], '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td>', row['PNDesc'], '</td>\n'))
#			tempList.extend(('\t\t\t\t\t\t\t\t<td>', row['ProjID'], '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td>', row['ProdCode'], '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-right\">', "{:,.0f}".format(row['Qty']), '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-right\">', "{:,.0f}".format(row['LineCharges']), '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-right\">', "{:,.0f}".format(row['MiscCharges']), '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-right\">', "{:,.0f}".format(row['Total']), '</td>\n'))
			tempList.append('\t\t\t\t\t\t\t\t</tr>\n')
		tempList.append('\t\t\t\t\t\t</table>\n')
		tempList.append('\t\t\t\t\t</div> <!-- close panel body -->\n')
		tempList.append('\t\t\t\t</div> <!-- close accordianOne -->\n')
		tempList.append('\t\t\t</div> <!-- close panel-primary -->\n')

		# join rows of list into a new strting, and write it to the page
		quoteString = ''.join(tempList).encode('utf-8')
		
		try:
			self.wfile.write(quoteString)
		except:
			self.wfile.write('sorry, i choked on something i found in the list of quotes\n')

		self.wfile.write("\t\t\t<p></p>\n")

		#repeat to order details
		conn.execute_query(orderSQL)
		tempList = []
		self.wfile.write("\n\t\t\t<div class=\"panel panel-primary\">\n")
		self.wfile.write("\n\t\t\t\t<div class=\"panel panel-heading\">\n")
		self.wfile.write("\t\t\t\t\t<h3 class=\"panel-title\">\n")
		# for PROJ reports, there's usually lots of data, so start with the panels collapsed
		# for TASK/ISSUE reports, there is usually less data, so start with the panels expanded
		self.wfile.write("\t\t\t\t\t\t<a data-toggle=\"collapse\" data-parent=\"#accordion\" href=\"#accordionTwo\">\n")
		self.wfile.write("\t\t\t\t\t\t\tOrder Details\n")
		self.wfile.write("\t\t\t\t\t\t</a>\n")
		self.wfile.write("\t\t\t\t\t</h3>\n")
		self.wfile.write("\t\t\t\t</div> <!-- close panel-heading -->\n")
		if (objectType == 'PROJ'):
			self.wfile.write("\n\t\t\t\t<div id=\"accordionTwo\" class=\"panel-collapse collapse\">\n")
		else:
			self.wfile.write("\n\t\t\t\t<div id=\"accordionTwo\" class=\"panel-collapse collapse in\">\n")
		self.wfile.write("\t\t\t\t\t<div class=\"panel-body\">\n")
		self.wfile.write("\t\t\t\t\t\t<table class=\"table table-striped\">\n")
		tempList.append('\t\t\t\t\t\t\t<tr><thead>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th>Order Date\n')
		tempList.append('\t\t\t\t\t\t\t\t<th>SO / Line / PO</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th>PN / Description</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th class=\"text-right\">Quantity</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th class=\"text-right\">Line Charges</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th class=\"text-right\">Misc Charges</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th class=\"text-right\">Total Value</th>\n')
		tempList.append('\t\t\t\t\t\t\t</tr></thead>\n')
		for row in conn:
			tempList.append('\t\t\t\t\t\t\t<tr>\n')
			try:
				dateString = (row['OrderDate']).strftime('%d %b %Y')
			except:
				dateString = str(None)
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-left\">', dateString, '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td>', row['OrderLinePO'], '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td>', row['PNDesc'], '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-right\">', "{:,.0f}".format(row['Qty']), '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-right\">', "{:,.0f}".format(row['LineCharges']), '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-right\">', "{:,.0f}".format(row['MiscCharges']), '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-right\">', "{:,.0f}".format(row['Total']), '</td>\n'))
			tempList.append('\t\t\t\t\t\t\t</tr>\n')
		tempList.append('\t\t\t\t\t\t</table>\n')
		tempList.append('\t\t\t\t\t</div> <!-- close panel body -->\n')
		tempList.append('\t\t\t\t</div> <!-- close accordianTwo -->\n')
		tempList.append('\t\t\t</div> <!-- close panel-primary -->\n')

		orderString = ''.join(tempList).encode('utf-8')
		self.wfile.write(orderString)

		self.wfile.write("\t\t\t<p></p>\n")


		# repeat for Job details
		conn.execute_query(jobSQL)
		tempList = []
		self.wfile.write("\n\t\t\t<div class=\"panel panel-primary\">\n")
		self.wfile.write("\n\t\t\t\t<div class=\"panel panel-heading\">\n")
		self.wfile.write("\t\t\t\t\t<h3 class=\"panel-title\">\n")
		self.wfile.write("\t\t\t\t\t\t<a data-toggle=\"collapse\" data-parent=\"#accordion\" href=\"#accordionThree\">\n")
		self.wfile.write("\t\t\t\t\t\t\tJob Details\n")
		self.wfile.write("\t\t\t\t\t\t</a>\n")
		self.wfile.write("\t\t\t\t\t</h3>\n")
		self.wfile.write("\t\t\t\t</div> <!-- close panel-heading -->\n")
		if (objectType == 'PROJ'):
			self.wfile.write("\n\t\t\t\t<div id=\"accordionThree\" class=\"panel-collapse collapse\">\n")
		else:
			self.wfile.write("\n\t\t\t\t<div id=\"accordionThree\" class=\"panel-collapse collapse in\">\n")
		self.wfile.write("\t\t\t\t\t<div class=\"panel-body\">\n")
		self.wfile.write("\t\t\t\t\t\t<table class=\"table table-striped\">\n")
		tempList.append('\t\t\t\t\t\t\t<tr><thead>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th>Created</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th>Job Number</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th>SO Number</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th>PN / Description</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th class=\"text-right\">Start Qty</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th class=\"text-right\">Complete Qty</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th class=\"text-right\">Due Date</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th class=\"text-right\">Job Completion Date</th>\n')
		tempList.append('\t\t\t\t\t\t\t</tr></thead>\n')
		for row in conn:
			tempList.append('\t\t\t\t\t\t\t<tr>\n')
			try:
				dateString = (row['CreateDate']).strftime('%d %b %Y')
			except:
				dateString = 'None'
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-left\">', dateString, '</td>\n'))
			try:
				jobString = "%s" % row['JobNum']
				jobString = jobString.encode('utf-8')
			except:
				jobString = 'Unicode'
				print "I just choked on a Unicode character in Job number field"
				print row['JobNum']
			tempList.extend(('\t\t\t\t\t\t\t\t<td>', jobString, '</td>\n'))
			try:
				orderString = "%d" % row['OrderNum']
			except:
				orderString = "Choked on Unicode"
				print "I just choked on a Unicode character in Order number field"
			tempList.extend(('\t\t\t\t\t\t\t\t<td>', orderString, '</td>\n'))
			try:
				pnString = "%s" % row['PNDesc']
			except:
				pnString = "Choked on Unicode"
				print "I just choked on a Unicode character in PNDesc field"
			tempList.extend(('\t\t\t\t\t\t\t\t<td>', str(row['PNDesc']), '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-right\">', "{:,.0f}".format(row['ProdQty']), '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-right\">', "{:,.0f}".format(row['QtyCompleted']), '</td>\n'))
			try:
				dateString = (row['DueDate']).strftime('%d %b %Y')
			except:
				dateString = 'None'
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-right\">', dateString, '</td>\n'))
			try:
				dateString = (row['JobCompletionDate']).strftime('%d %b %Y')	
			except:
				dateString = 'None'
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-right\">', dateString, '</td>\n'))
			tempList.append('\t\t\t\t\t\t\t</tr>\n')
		tempList.append('\t\t\t\t\t\t</table>\n')
		tempList.append('\t\t\t\t\t</div> <!-- close panel body -->\n')
		tempList.append('\t\t\t\t</div> <!-- close accordianThree -->\n')
		tempList.append('\t\t\t</div> <!-- close panel-primary -->\n')

		#newTempList = unicode(tempList, 'utf-8')
		#jobString = u' '.join(tempList).encode('utf-8')
		jobString = u' '.join(tempList).strip()
		
		self.wfile.write(jobString)

		self.wfile.write("\t\t\t<p></p>\n")


		# repeat for Invoice details
		conn.execute_query(invoiceSQL)
		tempList = []
		self.wfile.write("\n\t\t\t<div class=\"panel panel-primary\">\n")
		self.wfile.write("\n\t\t\t\t<div class=\"panel panel-heading\">\n")
		self.wfile.write("\t\t\t\t\t<h3 class=\"panel-title\">\n")
		self.wfile.write("\t\t\t\t\t\t<a data-toggle=\"collapse\" data-parent=\"#accordion\" href=\"#accordionFour\">\n")
		self.wfile.write("\t\t\t\t\t\t\tInvoice Details\n")
		self.wfile.write("\t\t\t\t\t\t</a>\n")
		self.wfile.write("\t\t\t\t\t</h3>\n")
		self.wfile.write("\t\t\t\t</div> <!-- close panel-heading -->\n")
		if (objectType == 'PROJ'):
			self.wfile.write("\n\t\t\t\t<div id=\"accordionFour\" class=\"panel-collapse collapse\">\n")
		else:
			self.wfile.write("\n\t\t\t\t<div id=\"accordionFour\" class=\"panel-collapse collapse in\">\n")
		self.wfile.write("\t\t\t\t\t<div class=\"panel-body\">\n")
		self.wfile.write("\t\t\t\t\t\t<table class=\"table table-striped\">\n")
		tempList.append('\t\t\t\t\t\t\t<tr><thead>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th>Invoice Number</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th>Invoice Date</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th>Sales Order</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th>Customer PO</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th class=\"text-right\">Open</th>\n')
		tempList.append('\t\t\t\t\t\t\t\t<th class=\"text-right\">Invoice Amount</th>\n')
		tempList.append('\t\t\t\t\t\t\t</tr></thead>\n')
		for row in conn:
			tempList.append('\t\t\t\t\t\t\t<tr>\n')
			tempList.extend(('\t\t\t\t\t\t\t\t<td>', str(row['InvoiceNum']), '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td>', (row['InvoiceDate']).strftime('%d %b %Y'), '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td>', str(row['OrderNum']), '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td>', str(row['PONum']), '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-right\">', "{:,.0f}".format(row['OpenInvoice']), '</td>\n'))
			tempList.extend(('\t\t\t\t\t\t\t\t<td class=\"text-right\">', "{:,.0f}".format(row['DocInvoiceAmt']), '</td>\n'))
			tempList.append('\t\t\t\t\t\t\t</tr>\n')
		tempList.append('\t\t\t\t\t\t</table>\n')
		tempList.append('\t\t\t\t\t</div> <!-- close panel body -->\n')
		tempList.append('\t\t\t\t</div> <!-- close accordianThree -->\n')
		tempList.append('\t\t\t</div> <!-- close panel-primary -->\n')

		invString = ''.join(tempList).encode('utf-8')
		
		self.wfile.write(invString)

		# close accordian panel group
		self.wfile.write("\t\t</div> <!-- close panel-group -->\n")
		
		# footnotes
		self.wfile.write("\n\t\t\t<h3>Footnotes</h3>\n")
		self.wfile.write("\t\t\t<ul>\n")
		self.wfile.write("\t\t\t\t<li>These data are current as of last night.</li>\n")
		self.wfile.write("\t\t\t\t<li>The reference number for this %s is: %s</li>\n" % (objectType, refNum) )
		self.wfile.write("\t\t\t\t<li>The JIRA Link for this %s is: %s</li>\n" % (objectType, jiraCode) )
		self.wfile.write("\t\t\t\t<li>The Project Code for this %s is: %s</li>\n" % (objectType, projectCode) )
		self.wfile.write("\t\t\t\t<li>Quote query searched Epicor field %s for: %s</li>\n" % (keyFieldQuote, idCode) )
		self.wfile.write("\t\t\t\t<li>Order query searched Epicor field %s for: %s</li>\n" % (keyFieldOrder, idCode) )
		self.wfile.write("\t\t\t\t<li>Order query searched Epicor field %s for: %s</li>\n" % (keyFieldJob, idCode) )
		self.wfile.write("\t\t\t\t<li>Order query searched Epicor field %s for: %s</li>\n" % (keyFieldInvoice, idCode) )
		self.wfile.write("\t\t\t\t<li>Served by %s</li>\n" % message)
		self.wfile.write("\t\t\t</ul>\n")

		# close container, body, and html
		self.wfile.write("\n\t</div> <!-- outer container -->\n")
		self.wfile.write("</body>\n</html>\n")

		# close the MSSQL connection
		conn.close()

		return


class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    """Handle requests in a separate thread."""

if __name__ == '__main__':
	server = ThreadedHTTPServer(('0.0.0.0', serverPort), GetHandler)
	print 'Listening on port %d, use <Ctrl-C> to stop' % serverPort
	server.serve_forever()
