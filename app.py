from flask import Flask, render_template, request, redirect, url_for
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

app = Flask(__name__)
app.debug = True

@app.route("/")
def form():
    return render_template("form.html")

@app.route("/home")
def member_home():
    return render_template("home.html")

@app.route("/result", methods=['GET', 'POST'])
def result():
    if request.method == 'POST':
        # Check if file is uploaded
        if 'file' not in request.files:
            return "Error: No file uploaded."

        f = request.files['file']

        # Check if file is an Excel file
        if not f.filename.endswith('.xlsx'):
            return "Error: Invalid file format. Please upload an Excel file."

        try:
            df = pd.read_excel(f)
            # Process the file data here
            print("File is read successfully!")

            def send_mail(i):
                try:
                    fromaddr = "nirmalasreemv@gmail.com"
                    toaddr = i[7]
                    msg = MIMEMultipart()
                    msg['From'] = fromaddr
                    msg['To'] = toaddr
                    msg['Subject'] = "Result Declaration"
                    body = "This is to inform to Roll No:" + str(i[0]) + "\nName:" + str(i[1]) + "\nYour obtained result is:\n\nOOAD:" + str(
                        i[2]) + "\nES:" + str(i[3]) + "\nWT:" + str(i[4]) + "\nCN:" + str(i[5]) + "\nAI:" + str(i[6]) + "\n\nRegards HOD\n"
                    msg.attach(MIMEText(body, 'plain'))

                    # Create and configure the SMTP server
                    s = smtplib.SMTP('smtp.gmail.com', 587)
                    s.starttls()

                    # Use the application-specific password for authentication
                    app_password = "iziyhuomuckascqa"  # Replace with your application-specific password
                    s.login(fromaddr, app_password)

                    # Send the email
                    s.send_message(msg)
                    s.quit()

                    print('Mail sent to ' + toaddr)
                except Exception as e:
                    print("An error occurred while sending the email:", e)

            df = df.to_numpy()

            for i in df:
                send_mail(i)

            return redirect(url_for('result'))  # Redirect to the /result URL
        except pd.errors.ExcelFileNotFound:
            return "Error: File not found."
        except pd.errors.ParserError:
            return "Error: Unable to read the file. Please make sure it is a valid Excel file."

    return render_template("result.html")

if __name__ == "__main__":
    app.run()
