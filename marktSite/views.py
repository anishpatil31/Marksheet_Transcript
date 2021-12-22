from django.shortcuts import render, redirect
from .forms import *
from .models import *
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import csv


def HomePage(request):
    if request.method == 'POST':
        form = Markform(request.POST, request.FILES)

        if form.is_valid():
            form.save()
            if request.POST.get('Roll'):
                return redirect('Proj1')
            if request.POST.get('Concise'):
                return redirect('GCM')
            if request.POST.get('Send'):
                return redirect('send')
    else:
        form = Markform()
    return render(request, 'marktSite/marktSite.html', {'form': form, 'check': 0})


def send(request):
    path = r"marksheets"
    path2 = ""
    response = r"media/file/responses.csv"
    dict = {}
    with open(os.path.join(path2, response), 'r') as file:
        reader = csv.reader(file)
        heading_of_columns = next(reader)
        for row in reader:
            roll = row[6]
            if roll not in dict:
                dict[roll] = []
            dict[roll].append(row[1])
            dict[roll].append(row[4])

    for excel_name in os.listdir(path):
        name = excel_name.split('.')[0]
        for email in dict[name]:
            fromaddr = "testcspy@gmail.com"
            toaddr = email
            msg = MIMEMultipart()
            msg['From'] = fromaddr
            msg['To'] = toaddr
            msg['Subject'] = "CS384 Marks"
            body = "PFA marks for the CS384 Quiz"
            msg.attach(MIMEText(body, 'plain'))
            filename = name + "_marksheet.xlsx"
            path_name = path + "/" + excel_name
            attachment = open(path_name, "rb")
            p = MIMEBase('application', 'octet-stream')
            p.set_payload((attachment).read())
            encoders.encode_base64(p)
            p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
            msg.attach(p)
            s = smtplib.SMTP('smtp.gmail.com', 587)
            s.starttls()
            s.login(fromaddr, "akashlord")
            text = msg.as_string()
            s.sendmail(fromaddr, toaddr, text)
            s.quit()
    return redirect('Home')
    # return render(request, 'marktSite/marktSite.html', {'form' : form, 'check':0})













