import xlwt 
from flask import *
from flask import Flask, render_template, redirect, url_for, request, session, flash, app, Blueprint, jsonify
from string import Template
from xlwt import Workbook 
app = Flask(__name__)
@app.route('/',methods=['GET','POST'])
def signingup(): 
		if request.method=='POST':
			username=request.form.get('username')
			print(username)
			password=request.form.get('password')
			email=request.form.get('email')
			courses=request.form.get('courseId')
			print(courses)
			with app.app_context():
				wb = Workbook()
				sheet1 = wb.add_sheet('Sheet 1')
				row = 0
				col = 0
				data=[]
				count=0
				data.append(username)
				data.append(password)
				data.append(email)
				data.append(courses)
				print(data)
				for i in  data:
					sheet1.write(row, col+count, i) 
					count+=1  
				wb.save('Faculty.xls')
		return render_template("signup.html")
if __name__=='__main__':
	app.debug=True
	app.run()