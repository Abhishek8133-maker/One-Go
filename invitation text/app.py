import os
from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
import random
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from datetime import datetime

app = Flask(__name__)

sentences = [
    "Join us for our next lecture on {topic} on {date} at {time}. The classroom link is: {link}.",
    "Attend our upcoming lecture on {topic} on {date} at {time}. Access the class here: {link}.",
    "Don't miss our lecture on {topic} on {date} at {time}. Join us using this link: {link}.",
    "Mark your calendar for a lecture on {topic} on {date} at {time}. Connect via: {link}.",
    "We invite you to our lecture on {topic} on {date} at {time}. Click here to join: {link}.",
    "Get ready for a lecture on {topic} on {date} at {time}. Use this link to access the class: {link}.",
    "Join us for an insightful lecture on {topic} on {date} at {time}. Classroom link: {link}.",
    "Be part of our upcoming lecture on {topic} on {date} at {time}. Join using: {link}.",
    "Save the date for a lecture on {topic} on {date} at {time}. Access the session here: {link}.",
    "You're invited to a lecture on {topic} on {date} at {time}. Join with this link: {link}.",
    "Don't miss our special lecture on {topic} on {date} at {time}. Click here to attend: {link}.",
    "Prepare yourself for a transformative lecture on {topic} on {date} at {time}. Join here: {link}.",
    "Our upcoming lecture on {topic} on {date} at {time} is a must-attend event. Join via: {link}.",
    "You're cordially invited to our lecture on {topic} on {date} at {time}. Use this link to join: {link}.",
    "Elevate your understanding with our lecture on {topic} on {date} at {time}. Join with: {link}.",
    "We're hosting a lecture on {topic} on {date} at {time}. Connect through this link: {link}.",
    "Learn more about {topic} in our upcoming lecture on {date} at {time}. Join via: {link}.",
    "Expand your expertise by attending our lecture on {topic} on {date} at {time}. Classroom link: {link}.",
    "Join us on {date} at {time} for a lecture on {topic}. Access the class using: {link}.",
    "Mark your schedule for our lecture on {topic} on {date} at {time}. Click to join: {link}."
]

@app.route('/', methods=['GET', 'POST'])
def index():
    invitation = ''
    names = []
    uploaded = False
    error = ''

    if request.method == 'POST':
        if 'upload' in request.form:
            file = request.files.get('file')
            if file and file.filename.endswith('.xlsx'):
                filename = secure_filename(file.filename)
                upload_path = os.path.join('uploads', filename)
                if not os.path.exists('uploads'):
                    os.makedirs('uploads')
                file.save(upload_path)
                
                # Read the file and extract names
                students = pd.read_excel(upload_path)
                names = students[['Name']].to_dict(orient='records')
                uploaded = True
            else:
                error = 'Invalid file format.'
        elif 'generate' in request.form:
            date = request.form.get('date')
            time = request.form.get('time')
            topic = request.form.get('topic')
            link = request.form.get('link')
            custom_message = request.form.get('message')  # Custom message (optional)

            try:
                formatted_date = datetime.strptime(date, '%Y-%m-%d').strftime('%B %d, %Y')
            except ValueError:
                formatted_date = date

            invitation = ""
            
            if date or time or topic or link:
                sentence = random.choice(sentences)
                invitation = sentence.format(date=formatted_date, time=time or '', topic=topic or '', link=link or '')

            if custom_message:
                if invitation:
                    invitation += f"\n\n{custom_message}"
                else:
                    invitation = f"\n{custom_message}"

            if not invitation:
                invitation = "No details provided."

    return render_template('index.html', invitation=invitation, names=names, uploaded=uploaded, error=error)

def read_students():
    # Assuming 'students.xlsx' will be used to get student details
    students = pd.read_excel('students.xlsx')
    return students[['Name', 'Email']]

@app.route('/send_email', methods=['POST'])
def send_email():
    invitation_template = request.form.get('invitation')

    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    smtp_user = 'shreyas.johri11@gmail.com'
    smtp_password = 'zmdi fwpq mqsl czyx'
    from_email = 'shreyas.johri11@gmail.com'
    
    # Ensure file is available and path is correct
    file_path = 'uploads/students.xlsx'
    if not os.path.exists(file_path):
        return jsonify({'status': 'error', 'message': 'Student file not found.'})

    students = pd.read_excel(file_path)

    subject = 'Lecture Invitation from InTrainTech'
    successes = 0
    errors = []

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)

            for index, row in students.iterrows():
                student_name = row['Name']
                to_email = row['Email']
                
                personalized_invitation = f"Hi {student_name},\n\n{invitation_template}"

                msg = MIMEText(personalized_invitation)
                msg['Subject'] = subject
                msg['From'] = from_email
                msg['To'] = to_email

                try:
                    server.sendmail(from_email, to_email, msg.as_string())
                    successes += 1
                except Exception as e:
                    errors.append(f"Failed to send email to {to_email}: {str(e)}")

        if errors:
            return jsonify({'status': 'error', 'message': f"Errors occurred: {', '.join(errors)}"})
        return jsonify({'status': 'success', 'message': f'Emails sent successfully to {successes} students!'})

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)})

if __name__ == '__main__':
    app.run(debug=True)
