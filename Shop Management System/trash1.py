
import smtplib
def mailboy(*arg):
    server = smtplib.SMTP_SSL('64.233.184.108', 465)
    server.login("romanogers99@gmail.com", "USEANDTHROW")
    server.sendmail(arg[0], arg[1],arg[2])
    server.quit()
