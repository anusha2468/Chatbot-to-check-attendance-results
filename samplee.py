import telebot
import pandas as pd
import ast
import time

attendanceFlag=0
resultsFlag=0

data=pd.read_excel("results.xlsx")
data=data.iloc[:,:].values
data=list(data)
Rollno=[]
for i in data:
    Rollno.append(i[0])
data1=pd.read_excel("attendance.xlsx")
data1=data1.iloc[:,:].values
data1=list(data1)

token='6850899605:AAGrlnczGMNMGjeEX2hjPSM2vGTGZZK8z74'
bot = telebot.TeleBot(token, parse_mode=None)

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
  bot.reply_to(message, "Welcome to students result and attendance check, what do you want to know? \n 1. /attendance \n 2. /results")

@bot.message_handler(commands=['results'])
def handle_results_request(message):
    global resultsFlag
    bot.reply_to(message, "Please enter the student ID to which you want to check results:")
    resultsFlag=1
@bot.message_handler(commands=['attendance'])
def handle_attendance_request(message):
    global attendanceFlag
    bot.reply_to(message, "Please enter the student ID for which you want to check attendance:")
    attendanceFlag=1
  
@bot.message_handler(regexp="[a-zA-Z0-9_]")
def handle_message(message):
    global attendanceFlag,resultsFlag
    message=str(message)
    k=ast.literal_eval(message)
    #print(k,type(k))
    chat_id=(k['from_user']['id'])
    student_id=k['text'].upper()
    #print(m)
    flag=0
    print(Rollno)

    for i in Rollno:
      if(i==(student_id) and attendanceFlag==1):
        attendanceFlag=0
        flag=1
        rollindex=Rollno.index(i)
        atten=data1[rollindex][1]
        if(atten>90):
           feed="very good ! keep it up"
        elif(atten>=75 and atten<=90):
           feed="good!"
        elif(atten>=65 and atten<75):
           feed="need to improve!"
        else:
           feed="very poor! likely to be detained"
        bot.send_message(int(chat_id),'Hi ! Nice meeting to you.Your Registered Number is '+ str(i) +'\nYour overall attendance is '+str(atten)+'% \nAccording to your attendance you are '+str(feed))
      if(i==(student_id) and resultsFlag==1):
        message=str(message)
        k=ast.literal_eval(message)
        chat_id=(k['from_user']['id'])
        student_id=k['text']
        flag=0
        for i in Rollno:
          if(i==(student_id) and resultsFlag==1):
            resultsFlag=0
            flag=1
            rollindex=Rollno.index(i)
            att=data[rollindex][1]
            agg=data[rollindex][2]
            bot.send_message(int(chat_id),'Hey,Nice meeting to you.Your Registered Numbers is ' + str(i) + '\nyour academic percentage is '+str(att)+'% \nyour aggregate is '+str(agg) + '%\n\nThank You')
        if(flag==0):
          bot.reply_to(message,"Sorry ! student ID is not present")

         
def listener(messages):
  for m in messages:
    print(str(m))

bot.set_update_listener(listener)

while True:
    try:
        bot.polling(none_stop=True, timeout=120)
    except Exception as e:
        print(f"Error: {e}")
        time.sleep(10)