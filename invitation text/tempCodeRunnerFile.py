from flask import Flask, render_template, request, jsonify
import random
import smtplib
from email.mime.text import MIMEText
import pandas as pd

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

def read_students():
    students = pd.read_excel('students.xlsx')  
    return students[['Name', 'Email']]  

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        date = request.form.get('date')
        time = request.form.get('time')
        topic = request.form.get('topic')
        link = request.form.get('link')
        text = request.form.get('text')
        

        sentence = random.choice(sentences)
        invitation = sentence.format(date=date or '', time=time or '', topic=topic or '', link=link or '' )

        return render_template('index.html', invitation=invitation , text = text)

    return render_template('index.html', invitation='')

@app.route('/send_email', methods=['POST'])
def send_email():
    invitation_template = request.form.get('invitation')
    message = request.form.get('text')
    print(message)

    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    smtp_user = 'shreyas.johri11@gmail.com'
    smtp_password = 'zmdi fwpq mqsl czyx'
    from_email = 'shreyas.johri11@gmail.com'
    
    students = read_students()

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
                
                if invitation_template:
                    personalized_invitation = f"Hi {student_name},\n\n{invitation_template}"
                else:
                    personalized_invitation = f"Hi {student_name},\n\n{message}"

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
