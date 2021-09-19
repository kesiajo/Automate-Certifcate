import pandas as pd 
from io import BytesIO
import os
from PIL import Image, ImageDraw, ImageFont
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


"""
	* Function Name: get_attachment()
	* Input:  image,  filename
        * Output: msg
        * Logic: attaches the file with the MIMEMultipart for created instance msg.
"""
def get_attachment(img, filename):
	bytes = BytesIO()
	img.save( bytes, format='JPEG' )
	msg = MIMEBase( filename, 'jpeg' )
	msg.set_payload( bytes.getvalue() )
	encoders.encode_base64( msg )
	msg.add_header( 'Content-Disposition', 'attachment', filename = filename )
	return msg


"""
        * Function Name: send_mail()
        * Input:  userid,  name, email, filename
        * Logic: Sends mail with attached certificate file by creating an SMTP session
"""
def send_mail(userid, name, email, filename):
        mail_content = "Hi " + name.title() + """,\n
        Thank you for participating in LinkedIn Branding Session organized by IEEE MBCET SB.\nPlease find your attached certificate.
        \nAwaiting your active participation in our upcoming events. 
	\nThanks & Regards,\nKesia Mary Joies \nChair | IEEE MBCET SB	
	"""
        senderAddress = 'kesiajoies@gmail.com'
        senderPswd = '******'
        receiverAddress = email

        # Setup the MIME
        message = MIMEMultipart()
        message['From'] = senderAddress
        message['To'] = receiverAddress
        # Subject line
        message['Subject'] = 'LinkedIn Branding Participation Certificate'   
        # The body and the attachments for the mail
        message.attach(MIMEText(mail_content, 'plain'))
        image = Image.open(filename)
        attachment = get_attachment(image,filename)
        message.attach(attachment)
        # Create SMTP session for sending the mail
        # Use Gmail with port
        session = smtplib.SMTP('smtp.gmail.com', 587) 
        # Enable security
        session.starttls() 
        # Login with mail_id and password
        session.login(senderAddress, senderPswd) 
        text = message.as_string()
        session.sendmail(senderAddress, receiverAddress, text)
        session.quit()
        print(f"{userid+1}\t Mail Sent\t {name}\t {receiverAddress}")


def main():
	# Reads excel sheet with the names and emails of attendees
	data = pd.read_excel(r'Linkedin Branding (Responses).xlsx')  #File name of the excel sheet
	total = len(data.index)

	for i in range(total):
                name = str.upper(data.at[i,'Name']) 
                email = data.at[i, 'Email']
                # Opens certificate template image
                image = Image.open('certificate.jpg') # Filename of certificate template
                # Font used for name
                bytes_font = "SF-Compact-Display-Heavy.otf" # Font used
                font = ImageFont.truetype(bytes_font, size = 260, encoding = "unic")
                (x, y) = (3505, 2623) # Coordinates of centre
                x -= (len(name)/2)* 125 
                color = 'rgba(6, 147, 164, 255)' # Text color
                img_draw = ImageDraw.Draw(image)
                img_draw.text((x, y), name, fill = color, font = font)
                filename = 'LinkedIn_' + str(name) + '.jpg'
                image.save(filename)
                send_mail(i, name, email, filename)


if __name__  == '__main__':
	main()
