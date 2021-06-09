from datetime import date
import win32com.client as win32
import csv

outlook = win32.Dispatch('outlook.application')


def sendEmail(userName, userRole, userEmail):
    string = 'https://docs.google.com/forms/d/e/1FAIpQLSf5F8Q21kYZt0ORwlepmss1bT8mjt9LZ7Uf-vKI2tHJ70sWiQ/viewform?'
    todays_date = date.today()
	
    #Name
    string += "entry.2005620554=" + userName
    #Year
    string += "&entry.1045781291_year=" + str(todays_date.year)
    #Month
    string += "&entry.1045781291_month=" + str(todays_date.month)
    #Day
    string += "&entry.1045781291_day=" + str(todays_date.day)
    #Role
    string += "&entry.1065046570.other_option_response=" + userRole + "&entry.1065046570=__other_option__"
    string += "&entry.1065046570=" + userRole
    #New Symptoms
    string += "&entry.1166974658=No"
    #Close Contact
    string += "&entry.839337160=No"
    #Travel
    string += "&entry.2062676151=No"

    mail = outlook.CreateItem(0)
    mail.To = userEmail
    mail.Subject = 'COVID Reminder Email'
    mail.Body = 'This form is pre-filled with your information.  For each COVID question, the default answer will be NO.  If you need to answer yes to one, or all, of the COVID questions, simply change the answer to Yes and submit. \n\nHere is the link to today\'s COVID Survey:\n\n ' + string + '\n\n     ~Brought to you by Tim Bomers.  This is an automatic email'
    mail.Send()
    


with open('employee_file.csv', mode='r') as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=',')
    for row in csv_reader:
        sendEmail(row[0], row[1], row[2])
