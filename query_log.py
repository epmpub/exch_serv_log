import datetime
from flask import Flask, render_template, request
from flask_cors import CORS, cross_origin
import subprocess

app = Flask(__name__)
cors = CORS(app, resources={r"/api/*": {"origins": "*"}})
cors = CORS(app)

def run(cmd):
    completed = subprocess.run(["powershell", "-Command", cmd], capture_output=True)
    return completed






@app.route('/api/query', methods=['GET', 'POST'])
@cross_origin()
def query():
    sender = request.json["sender"]
    recipients = request.json["recipients"]
    server = request.json["server"]
    eventId = request.json["eventId"]
    start = request.json["start"]
    end = request.json["end"]

    filter = f"|Select-Object Sender,Recipients,Timestamp,MessageSubject"
    date = str(datetime.datetime.now()).replace(' ','_').replace(':','_').replace('.','_')
    logpath = f"c:\\logs\log_{date}"
    tee = f"| tee -FilePath {logpath}.log"
    query_command = f"Get-MessageTrackingLog -Server {server} -Sender {sender} -Recipients {recipients} -Eventid {eventId} -Start '{start}' -End '{end}' {filter} {tee}"
    print(query_command)

    query_info = run(query_command)
    if query_info.returncode != 0:
        print("An error occured: %s", query_info.stderr)
    else:
        print("query_info command executed successfully!")

    print((query_info.stderr))


    tempList = query_info.stdout.split(b'\r\n')
    output_list = []
    for item in tempList:
        #清理输出中不需要的字符
        output_list.append(bytes.decode(item).replace('-',''))

    # return render_template('query.html', say=output_list, to=request.form['sender'])
    return {'output':output_list}


@app.route('/api/ExchServer',methods=['GET','POST'])
@cross_origin()
def ExchServer():
    sender = request.json['sender']
    query_command = f"(get-mailbox {sender}).ServerName"
    query_info = run(query_command)
    if query_info.returncode != 0:
        print("An error occured: %s", query_info.stderr)
    else:
        print("query_info command executed successfully!")

    print((query_info.stderr))


    tempList = query_info.stdout.split(b'\r\n')
    output_list = []
    for item in tempList:
        output_list.append(bytes.decode(item))
    return {
        "output":output_list
    }
