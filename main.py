import pandas as pd
import datetime
import smtplib

# enter your authentication details
GMAIL_ID = ''
GMAIL_PSWD = ''

def sendEmail(to, sub, msg):
    
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PSWD)    
    s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")    
    s.quit()

if __name__ == "__main__":      
    df = pd.read_excel('data.xlsx')
    today = datetime.datetime.now().strftime("%d-%m")
    year = datetime.datetime.now().strftime("%Y")    
    
    writeInd = []
    for index, item in df.iterrows():        
        bday = item['Birthday'].strftime("%d-%m")
        if today == bday and year not in str(item["Year"]):
            sendEmail(item['Email'], "Happy Birthday", item["Message"])
            writeInd.append(index)    
    
    for i in writeInd:        
        yr = df.loc[i, 'Year']
        df.loc[i, 'Year'] = str(yr) + ',' + str(int(year)+1)
    
    df.to_excel('data.xlsx')    
    print("How many new records you would like to enter?")
    n=int(input())
    for _ in range(n):
        print("Enter Name: ")
        name=input()
        print("Enter DOB (yyyy-mm-dd): ")
        dob=input()
        print("Enter their Email: ")
        mail=input()
        print("Enter your Message: ")
        msg=input()        
        print("Enter Year: ")
        yrr=input()
        
        dict={'Name':name, 'Birthday':dob, 'Email':mail, 'Message':msg, 'Year':yrr}
        df=df.append(dict, ignore_index=True)
        df.to_excel('data.xlsx')
        
    #print(df)
