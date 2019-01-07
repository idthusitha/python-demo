#!/usr/bin/python
#
# Small script to show PostgreSQL and Pyscopg together
#python /rezsystem/python/OAH_REPORTS.py 'MAY-17-2016 00:00:00' 'JUL-11-2016 23:59:59'

import psycopg2
import sys
import pprint
import smtplib
import datetime
import csv

def main():
	conn_string = "host='134.213.217.98' port='5432' dbname='PGRS1' user='postgres' password='res13pg'"
	#conn_string = "host='162.13.185.21' port='5432' dbname='PGRS1' user='postgres' password='res13pg'"
	#conn_string = "host='192.168.1.246' port='5433' dbname='PGDEV' user='postgres' password='res13pg'"
	# print the connection string we will use to connect
	print "===>Connecting to database->%s" % (conn_string)
	
	current_time = datetime.datetime.strftime(datetime.datetime.now(), '%Y-%m-%d-%H:%M:%S')
	file_name_save_sales = 'OAH_sales_report_'+current_time+'.csv'
	file_name_sales = '/rezsystem/portal_static_data/OAH_sales_report_'+current_time+'.csv'

	file_name_save_refund = 'OAH_refund_report_'+current_time+'.csv'
	file_name_refund = '/rezsystem/portal_static_data/OAH_refund_report_'+current_time+'.csv'

	file_name_save_profit_and_loss = 'OAH_profit_and_loss_report_'+current_time+'.csv'
	file_name_profiandloss = '/rezsystem/portal_static_data/OAH_profit_and_loss_report_'+current_time+'.csv'

	print "===>"+file_name_sales
	print "===>"+file_name_refund
	print "===>"+file_name_profiandloss

 	# report parameters
	# from date
	from_date = (datetime.datetime.now() - datetime.timedelta(days=7)).strftime("'%b-%d-%Y 00:00:00'")
	# to date
	to_date = datetime.datetime.now().strftime("'%b-%d-%Y 23:59:59'")

	try:
		from_date = "'"+sys.argv[1]+"'"
	except:
		print "from date not set"

	try:
		to_date = "'"+sys.argv[2]+"'"
	except:
		print "to date not set"

	print "from date ===>"+from_date
	print "to date ===>"+to_date

	# conn.cursor will return a cursor object, you can use this cursor to perform queries
	# curso = conn.cursor()
	 
	# execute our Query
	# Need to set follwing path to posgres usser access
	# query_string_salesr = "select omanairholidays_admin.fn_tour_sale_report_query('refcursor');fetch all in refcursor;"
	# query_string_refund = "select omanairholidays_admin.fn_tour_refund_report_query('refcursor');fetch all in refcursor;"
	# query_string_profiandloss = "select omanairholidays_admin.fn_profi_and_loss_report_query('refcursor');fetch all in refcursor;"

	# retrieve the records from the database
	# records = cursor.fetchall()
	
	# print out the records using pretty print
	# note that the NAMES of the columns are not shown, instead just indexes.
	# for most people this isn't very useful so we'll show you how to return
	# columns as a dictionary (hash) in the next example.
	#pprint.pprint(records)

	###################################################Save file here
	# sales reprot	
	# get a connection, if a connect cannot be made an exception will be raised here
	conn = psycopg2.connect(conn_string)
	curso = conn.cursor()
	curso.execute("select omanairholidays_admin.fn_tour_sale_report_query('refcursor'," + from_date + "," + to_date + ");fetch all in refcursor;")
	with open(file_name_sales, "wb") as csv_file_sales:
		csv_writer_sales = csv.writer(csv_file_sales, delimiter ='~',quotechar =',',quoting=csv.QUOTE_MINIMAL)
		csv_writer_sales.writerow([i[0] for i in curso.description]) # write headers
		csv_writer_sales.writerows(curso)	
	
	curso.close()
	del curso

	# closing the db connection 
	conn.close()

	# refund reprot	
	# get a connection, if a connect cannot be made an exception will be raised here
	conn = psycopg2.connect(conn_string)
	curso = conn.cursor()
	curso.execute("select omanairholidays_admin.fn_tour_refund_report_query('refcursor'," + from_date + "," + to_date + ");fetch all in refcursor;")
	with open(file_name_refund, "wb") as csv_file_refund:
		csv_writer_refund = csv.writer(csv_file_refund, delimiter ='~',quotechar =',',quoting=csv.QUOTE_MINIMAL)
		csv_writer_refund.writerow([i[0] for i in curso.description]) # write headers
		csv_writer_refund.writerows(curso)

	curso.close()
	del curso

	# closing the db connection 
	conn.close()

	# profitandloss reprot	
	# get a connection, if a connect cannot be made an exception will be raised here
	conn = psycopg2.connect(conn_string)
	curso = conn.cursor()
	curso.execute("select omanairholidays_admin.fn_profit_and_loss_report_query('refcursor'," + from_date + "," + to_date + ");fetch all in refcursor;")
	with open(file_name_profiandloss, "wb") as csv_file_profiandloss:
		csv_writer_profiandloss = csv.writer(csv_file_profiandloss, delimiter ='~',quotechar =',',quoting=csv.QUOTE_MINIMAL)
		csv_writer_profiandloss.writerow([i[0] for i in curso.description]) # write headers
		csv_writer_profiandloss.writerows(curso)

	curso.close()
	del curso

	# closing the db connection 
	conn.close()

	####################################################sending mail
	# Import the email modules we'll need
	from email.mime.text import MIMEText
	from email.mime.application import MIMEApplication
	from email.mime.multipart import MIMEMultipart
	from smtplib import SMTP
	
	fromaddress = 'omanairholidays@rezgateway.com'
	me =  'thusitha@rezgateway.com'
	cc = ["thusitha@rezgateway.com", "chamidu@rezgateway.com","suranga@rezgateway.com"]
	to = ["thusitha@rezgateway.com","thusitha@rezgateway.com"]
	msg = MIMEMultipart()

	# me == the sender's email address
	# to == the recipient's email address
	msg['Subject'] = 'OAH Tour salse / Refund / Profit and loss REPORTS CSV - ' + from_date + ' - ' + to_date
	msg['From'] = me
	msg['To'] = ",".join(to)
	msg['Cc'] = ",".join(cc)

	# This is the textual part:
	part = MIMEText("Hi All,\n\nPlease find attached csv files herewith. Please use ~ for Seperator other option for read CSV file.\n\n\nRegards,\nThusitha Dissanayaka\n\n\n\nThis is an auto generated e-mail.")
	msg.attach(part)
	 
	# This is the binary part(The Attachment):
	# attachement 1
	part_sales = MIMEApplication(open(file_name_sales,"rb").read())
	part_sales.add_header('Content-Disposition', 'attachment', filename=file_name_save_sales)
	msg.attach(part_sales)

	# attachement 2
	part_refund = MIMEApplication(open(file_name_refund,"rb").read())
	part_refund.add_header('Content-Disposition', 'attachment', filename=file_name_save_refund)
	msg.attach(part_refund)

	# attachement 3
	part_profiandloss = MIMEApplication(open(file_name_profiandloss,"rb").read())
	part_profiandloss.add_header('Content-Disposition', 'attachment', filename=file_name_save_profit_and_loss)
	msg.attach(part_profiandloss)
	
	# Send the message via our own SMTP server, but don't include the
	# envelope header.
	s = smtplib.SMTP('192.168.1.229',25)
	s.sendmail(me, to+cc, msg.as_string())
	s.quit()
	print "Successfully sent email"

 
if __name__ == "__main__":
	main()

