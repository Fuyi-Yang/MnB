import smtplib


sender_email = "#put your email here"
rec_email = [] #put receiver emails here]
title = "put title"
context = "Put Message"
message = "Subject: {}\n\n{}".format("title", "context")

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(sender_email, "put app password here")
print("Login success")
server.sendmail(sender_email, rec_email, message)
server.quit()
print("Email has been sent to", rec_email)