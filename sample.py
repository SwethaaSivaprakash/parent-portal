from flask import *  
app = Flask(__name__)  
 
@app.route('/', methods = ['GET','POST'])  
def upload():
	if request.method == 'POST':  
		f = request.files['file']  
		f1=f.filename
		return f1    
	return render_template("sample.html")  
 
  
if __name__ == '__main__':  
    app.run(debug = True)  