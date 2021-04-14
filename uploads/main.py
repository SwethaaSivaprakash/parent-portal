import xlrd
from flask import *
from string import Template
import os
import pandas as pd
from werkzeug.utils import secure_filename
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
ALLOWED_EXTENSIONS = {'cvs', 'dif', 'ods', 'ots', 'tsv', 'xlm','xls','xlsb','xlsm','xlsx','xlt','xltm','xltx'}
app=Flask(__name__)
app.secret_key = "kannihilators"
UPLOAD_FOLDER = 'D:\\auxano_connections\\uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
uname=""
coursename=""
total=0

@app.errorhandler(404)
def page_not_found(name):
	with app.app_context():
		return render_template("404.html")
def get_contacts(filename,a1):
	names = []
	emails = []
	attendance=[]
	marks=[]
	try:
		wb = xlrd.open_workbook(filename) #opens he excel sheet
		sheet = wb.sheet_by_index(0) #gets the first sheet
		sheet.cell_value(1,1)
		total=sheet.nrows-1
		for i in range(sheet.nrows-1):
			if (a1=='CAT 1'):
				marks.append(sheet.cell_value(i+1,5))
			elif (a1=='CAT 2'):
				marks.append(sheet.cell_value(i+1,6))
			else:		
				marks.append(sheet.cell_value(i+1,7))
			
			names.append(sheet.cell_value(i+1,2))	
			emails.append(sheet.cell_value(i+1,4))
			attendance.append(sheet.cell_value(i+1,10))
			cname=sheet.cell_value(i+1,0)
			
		if (names==[] or emails==[] or marks==[] or attendance==[] or cname==" "):
			raise Exception("Content not found")
		return names,emails,attendance,marks,cname,total
	except FileNotFoundError as e:
		with app.app_context():
			return render_template("FNF.html",errormsg=str(e))
	except Exception as e:
		with app.app_context():
			return render_template("Error.html",errormsg=str(e))
def read_template(filename):
	try:
		with open(filename, 'r', encoding='utf-8') as template_file:
			template_file_content = template_file.read()
			return Template(template_file_content)
	except FileNotFoundError as e:
		with app.app_context():
			return render_template("FNF.html",errormsg=str(e))
@app.route('/mail',methods=['GET', 'POST'])			
def sendmail(**nthng):
	try:
		if request.method == "POST":
			select = request.form.get('comp_select')
			a=str(select)
			print(a)
			f = request.files['file']  
			f1=f.filename
			print(f1)
			password = request.form["username"]
			with app.app_context():
				count=0
				names,emails,attendance,marks,cname,tot = get_contacts(f1,a)  # read contacts
				port = 465  # For SSL
				# Create a secure SSL context
				context = ssl.create_default_context()
				context = ssl.create_default_context()
				server = smtplib.SMTP('smtp.gmail.com')
				server.connect('smtp.gmail.com', '587')
				server.ehlo()
				server.starttls()
				server.login("auxanoconnections@gmail.com", password)
				for name, email,attendance,marks in zip(names, emails,attendance,marks):
					msg = MIMEMultipart()       # create a message
					# add in the actual person name to the message template
					if (marks<25 and attendance<75.00):
						message_template=read_template('messagelowattendancemark.txt')
						message = message_template.substitute(PERSON_NAME=name.title(),COURSE_NAME=cname,ATTENDANCE=attendance,MARKS=marks)
					elif (marks<25):
						message_template=read_template('messagelowmark.txt')
						message = message_template.substitute(PERSON_NAME=name.title(),COURSE_NAME=cname,ATTENDANCE=attendance,MARKS=marks)
					elif (attendance<75.00):
						message_template=read_template('messagelowattendance.txt')
						message = message_template.substitute(PERSON_NAME=name.title(),COURSE_NAME=cname,ATTENDANCE=attendance,MARKS=marks)
					else:
						message_template=read_template('message.txt')
						message = message_template.substitute(PERSON_NAME=name.title(),COURSE_NAME=cname,ATTENDANCE=attendance,MARKS=marks)
					# setup the parameters of the message
					msg['From']="auxanoconnections@gmail.com"
					msg['To']=email
					msg['Subject']="This is TEST"
					# add in the message body
					msg.attach(MIMEText(message, 'plain'))
					# send the message via the server set up earlier.
					server.send_message(msg)
					count+=1
					del msg	 
				server.quit()
				print(count,tot)
				if (count==tot):
					return render_template("Success.html")
				else:
					return "Mail not sent"
		return render_template("mail1.html")	
	except Exception as  e:
		with app.app_context():
			return render_template("Error.html",errormsg=str(e))			
@app.route('/upload/<filename>')
def uploaded_file(filename):
    return send_from_directory(os.path.dirname(UPLOAD_FOLDER),filename)
#To check the file format
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

#Main page for uploading
@app.route("/upload", methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        #print(request.files['file'])
        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            file.save(os.path.join(app.config['UPLOAD_FOLDER'],file.filename))
            return redirect(url_for('uploaded_file',filename=file.filename))
            #data_xls = pd.read_excel(file)
            
            #return data_xls.to_html()
        else:
            submit_name = file.filename
            #if '.' in submit_name and submit_name.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS:
            #    filename = secure_filename(submit_name)
            #    form.file_upload.data.save('uploads/' + filename)
             #   return redirect('home')
        #else:
            flash('File (%s) is not an accepted format' % submit_name)
            print (submit_name)
    return render_template('base.html')
    

@app.route("/export", methods=['GET'])
def export_records():
    return 

@app.route('/suggestions', methods=['GET', 'POST'])
def suggestions(**nthng):
	if request.method == "POST":
		mail_id = request.form["text"]
		password = request.form["text1"]
		message = request.form["text2"]
		port = 465  # For SSL
		# Create a secure SSL context
		context = ssl.create_default_context()
		context = ssl.create_default_context()
		msg = MIMEMultipart()       # create a message
		msg['From']=mail_id
		msg['To']="auxanoconnections@gmail.com"
		msg['Subject']="Suggestions->Auxano_connections"
		server = smtplib.SMTP('smtp.gmail.com')
		server.connect('smtp.gmail.com', '587')
		server.starttls()
		server.login(msg['From'], password)
		# add in the message body
		msg.attach(MIMEText(message, 'plain'))
		# send the message via the server set up earlier.
		server.sendmail(msg['From'],msg['To'],msg.as_string())
		server.quit()
	return render_template('suggestions.html')
@app.route('/',methods=['GET', 'POST'])
def login():
	try:
		count=0
		if request.method == 'POST':
			uname = request.form['nm']
			password=request.form['pw']
			print(password)
			with app.app_context(): 
				df = (r'D:\auxano_connections\Faculty.xlsx')
				wb = xlrd.open_workbook(df) #opens the excel sheet
				sheet = wb.sheet_by_index(0) #gets the first sheet
				for x in range(sheet.nrows-1):
					un=sheet.cell_value(x+1,5)
					passw=sheet.cell_value(x+1,6)
					print(passw)
					if (uname == un and password==passw):
						count+=1
						cid=sheet.cell_value(x+1,4).split()
						return render_template('course.html',cid=cid)
				if count==0:
					return render_template('UsernotFound.html')
	except FileNotFoundError:
		print("File not found!")
	except Exception as e:
		print(e)
	if "course" in request.form:
		return render_template('option.html')
	elif "mail" in request.form:	
		return render_template('mail1.html')
	elif "upload" in request.form:	
		return render_template('base.html')	
	else:
		return render_template('login.html')
if __name__=='__main__':
	app.debug=True
	app.run() 