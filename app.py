from flask import Flask , render_template, request,redirect,url_for
import xlrd
from flask_mail import Mail, Message

   
app= Flask(__name__)
 
app.config['MAIL_SERVER']='smtp.gmail.com'
app.config['MAIL_PORT'] = 465
app.config['MAIL_USERNAME'] = 'digiat590@gmail.com'
app.config['MAIL_PASSWORD'] = '18me1a0590'
app.config['MAIL_USE_TLS'] = False
app.config['MAIL_USE_SSL'] = True
mail = Mail(app)


path="cybersecurity.xlsx"
wb=xlrd.open_workbook(path)
ws=wb.sheet_by_index(0)
row=(ws.nrows)
col=(ws.ncols)
data=[[ws.cell_value(r,c)for c in range(col)]for r in range(row)]
data=data[1:]

path1="ai&ds.xlsx"
wb1=xlrd.open_workbook(path1)
ws1=wb1.sheet_by_index(0)
row1=(ws1.nrows)
col1=(ws1.ncols)
data1=[[ws1.cell_value(r,c)for c in range(col1)]for r in range(row1)]
data1=data1[1:]

path2="demo.xlsx"
wb2=xlrd.open_workbook(path2)
ws2=wb2.sheet_by_index(0)
row2=(ws2.nrows)
col2=(ws2.ncols)
data3=[[ws2.cell_value(r,c)for c in range(col2)]for r in range(row2)]
data3=data3[1:]

# home url
@app.route("/")
@app.route("/home")
def home():
    return render_template("index.html")

@app.route("/login",methods=["GET"])
def login():
    return render_template("about.html")

@app.route("/contact")
def contact():
    return render_template("contact.html")

@app.route("/login",methods=["POST"])
def login_post():
    id=request.form.get("id")
    id=id.upper()
    
    if id=="20MEIA46":
        return redirect(url_for("portal",branch="Cyber security",link="cyber"))
    elif id=="20MEIA45":
        return redirect(url_for("portal",branch="AI & DS",link="ai"))
    elif id=="18ME1A05":
        return redirect(url_for("portal",branch="Testing Demo",link='demo'))
    else :
        mess="oops, check the Branch ID is wrong"
        return render_template("about.html",Message=mess)

@app.route("/portal",methods=["GET"])
def portal():
    branch=request.values.get("branch")
    return render_template("symptoms.html",branch=branch)

@app.route("/portal",methods=["POST"])
def portal_1():
    mail=request.values.get("mail",default="null")
    link=request.values.get("link")
    return redirect(url_for(link, mail=mail))

@app.route("/cyber",methods=["GET"])
def cyber() :
    mail=request.values.get("mail",default="null")
    return render_template("attendence.html",data=data,mess="Cyber Security",mail=mail)

@app.route("/cyber",methods=["POST"])
def cyber1() :
    listed=[]
    b=['sameenamz@gmail.com']
    default='ABSENT'
    message ="Absent's list \n"
    for i in data :
        a=request.form.get(str(i[0]), default)
        if a=='ABSENT':
            listed.append(str(i[0]))
    if mail!="null" :
        b.append(request.form.get("mail"))
    msg = Message( 
            'Absent list ', 
            sender ='digiat590@gmail.com', 
             recipients = b
           ) 
    
    msg.body = 'Absents list '
    for i in listed :
        msg.body+=" \n"+i


    mail.send(msg)
    count=len(listed)

   
    return render_template("blog.html",count=count)

@app.route("/ai",methods=["GET"])
def ai() :
    
    
    data2=data1[::-1]
    return render_template("attendence.html",data=data1,data1=data2)

@app.route("/ai",methods=["POST"])
def ai1() :
    listed=[]
    b=['kallakiran1974@gmail.com']
    default='ABSENT'
    message ="Absent's list \n"
    for i in data :
        a=request.form.get(str(i[0]), default)
        if a=='ABSENT':
            listed.append(str(i[0]))
    if mail!="null" :
        b.append(request.form.get("mail"))
    msg = Message( 
            'Absent list ', 
            sender ='digiat590@gmail.com', 
             recipients = b
           ) 
    
    msg.body = 'Absents list '
    for i in listed :
        msg.body+=" \n"+i


    mail.send(msg)
    count=len(listed)
   
    return render_template("blog.html",count=count)

# testing part

@app.route("/demo",methods=["GET"])
def demo() :
    
    
    data2=data3[::-1]
    return render_template("attendence.html",data=data3,data1=data2)

@app.route("/demo",methods=["POST"])
def demo_one() :
    listed=[]
    b=['gopireddy590@gmail.com']
    default='ABSENT'
    message ="Absent's list \n"
    for i in data :
        a=request.form.get(str(i[0]), default)
        if a=='ABSENT':
            listed.append(str(i[0]))
    if mail!="null" :
        b.append(request.form.get("mail"))
    msg = Message( 
            'Absent list ', 
            sender ='digiat590@gmail.com', 
             recipients = b
           ) 
    
    msg.body = 'Absents list '
    for i in listed :
        msg.body+=" \n"+i


    mail.send(msg)
    count=len(listed)
   
    return render_template("blog.html",count=count)



if __name__ == "__main__" :
    app.run(debug=True)