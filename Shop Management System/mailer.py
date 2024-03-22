
import smtplib
def mailboy(*arg):
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.login("romanogers99@gmail.com", "USEANDTHROW")
    server.sendmail('romanoger99@gmail.com', arg[1],arg[2])
    server.quit()
