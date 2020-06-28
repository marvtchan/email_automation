# email_class.py

import os
import smtplib
import imghdr
from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText  # Added
from email.mime.image import MIMEImage


class Email:
	def __init__(self, email, password):
		self.email = email 
		self.password = password


	def send_email(self, name, email_to, order_number, new_path):
		# Create the root message and fill in the from, to, and subject headers
		msgRoot = MIMEMultipart('related')
		msgRoot['Subject'] = 'Here is your proof! #' + str(order_number)
		msgRoot['From'] = self.email
		msgRoot['To'] = str(email_to)
		msgRoot.preamble = 'This is a multi-part message in MIME format.'

		# Encapsulate the plain and HTML versions of the message body in an
		# 'alternative' part, so message agents can decide which they want to display.
		msgAlternative = MIMEMultipart('alternative')
		msgRoot.attach(msgAlternative)

		msgText = MIMEText('This is the alternative plain text message.')
		msgAlternative.attach(msgText)

		html = ("""<!DOCTYPE html>
			<html>
				<body>

					<br><p1>Hey %s! </p1><br>

					<br><p2>Attached to this email are the proof(s) of your Star Map(s). Let us know how it looks. If it looks good,</p2> 
					<b><p2> reply to this email: APPROVED.</p2></b><br>
		

					<br><p3>If we don't hear from you within 3 days, we will automatically push it into production. </p3> <br>
					
					
					<br><p4>(FYI -- the quicker we hear back from you, the sooner production will start)</p4><br>

					<br><p5>Best,</p5><br>

					<p6>Tyler</p6>

					<br><p7>Cosmic Prints</p7><br>
					
						<br><img src="cid:image1" width="592.4" height="704.2"><br>  
					</center>

				</body>	
			</html>""" )

		# We reference the image in the IMG SRC attribute by the ID we give it below
		msgText = MIMEText(html %(str(name)), 'html')
		msgAlternative.attach(msgText)

		# This example assumes the image is in the current directory
		star_map = open(new_path + '/' + str(order_number) + '.jpg', 'rb')
		msgImage = MIMEImage(star_map.read())
		msgImage.add_header("Content-ID", '<image1>')
		msgImage.add_header("Content-Disposition", "attachment; filename= %s" % str(order_number))
		msgRoot.attach(msgImage)
		star_map.close()


		with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
		# with smtplib.SMTP('localhost', 1025) as smtp:
			smtp.login(self.email, self.password)
			smtp.send_message(msgRoot)
