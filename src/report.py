#!/usr/bin/python
#
# Small script to show PostgreSQL and Pyscopg together
#

import psycopg2
import sys
import pprint
import smtplib
import datetime
import csv

def main():
	conn_string = "host='192.168.1.246' port='5433' dbname='PGDEV' user='postgres' password='res13pg'"
	# print the connection string we will use to connect
	print "===>Connecting to database->%s" % (conn_string)
	
	current_time = datetime.datetime.strftime(datetime.datetime.now(), '%Y-%m-%d-%H:%M:%S')
	file_name_save = 'OAH_report_'+current_time+'.csv'
	file_name = '/rezsystem/portal_static_data/OAH_report_'+current_time+'.csv'
	print "===>"+file_name
 
	# get a connection, if a connect cannot be made an exception will be raised here
	conn = psycopg2.connect(conn_string)
	 
	# conn.cursor will return a cursor object, you can use this cursor to perform queries
	cursor = conn.cursor()
	 
	# execute our Query
	# Need to set follwing path to posgres usser access
	query_string = "select rezbase_v3_admin.fn_tour_sale_report_query('refcursor');fetch all in refcursor;"
	#query_string = "COPY (select * from rezbase_v3_admin.screen) TO '/rezsystem/portal_static_data/csv/OAH_tour_salse_report_c.csv' DELIMITER '~' CSV HEADER;"
	cursor.execute(query_string)
	 
	# retrieve the records from the database
	#records = cursor.fetchall()
	
	# print out the records using pretty print
	# note that the NAMES of the columns are not shown, instead just indexes.
	# for most people this isn't very useful so we'll show you how to return
	# columns as a dictionary (hash) in the next example.
	#pprint.pprint(records)

	print "test ===>"+file_name
	###################################################Save file here	
	with open(file_name, "wb") as csv_file:
		csv_writer = csv.writer(csv_file)
		csv_writer.writerow([i[0] for i in cursor.description]) # write headers
		csv_writer.writerows(cursor)	
	
	####################################################sending mail
	# Import the email modules we'll need
	from email.mime.text import MIMEText
	from email.mime.application import MIMEApplication
	from email.mime.multipart import MIMEMultipart
	from smtplib import SMTP
	
	me =  'thusitha@rezgateway.com'
	to = ["thusitha@rezgateway.com", "chamidu@rezgateway.com"]
	msg = MIMEMultipart()

	# me == the sender's email address
	# to == the recipient's email address
	msg['Subject'] = 'OAH Sales report %s' % file_name_save
	msg['From'] = me
	msg['To'] = ",".join(to)

	# This is the textual part:
	part = MIMEText("Hi All,\n\nI have attached relevent csv file in this mail.\n\n\nRegards,\nThusitha Dissanayaka")
	msg.attach(part)
	 
	# This is the binary part(The Attachment):
	part = MIMEApplication(open(file_name,"rb").read())
	part.add_header('Content-Disposition', 'attachment', filename=file_name_save)
	msg.attach(part)
	
	# Send the message via our own SMTP server, but don't include the
	# envelope header.
	s = smtplib.SMTP('192.168.1.229',25)
	s.sendmail(me, to, msg.as_string())
	s.quit()
	print "Successfully sent email"

 
if __name__ == "__main__":
	main()

