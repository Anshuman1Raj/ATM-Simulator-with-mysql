import smtplib
import random
from win32com.client import Dispatch
import mysql.connector
from mysql.connector import Error


a = random.randint(100000, 999999)

myemail = "your_email@example.com"
password = "your_password"

subject = "Your Verification Code"
body = f"Your verification code is {a}."
msg = f"Subject: {subject}\n\n{body}"

db_host = "localhost"
db_name = "your_database_name"
db_user = "your_username"
db_password = "your_password"

try:
    connection = mysql.connector.connect(
        host=db_host,
        database=db_name,
        user=db_user,
        password=db_password
    )

    if connection.is_connected():
        print("Successfully connected to the database")


    connection_smtp = smtplib.SMTP("smtp.gmail.com", port=587)
    connection_smtp.starttls()
    connection_smtp.login(user=myemail, password=password)
    connection_smtp.sendmail(from_addr=myemail, to_addrs="Email You want to send otp", msg=msg)
    connection_smtp.quit()
    print("Email sent successfully!")

    speak = Dispatch("SAPI.SpVoice")
    speak.Voice = speak.GetVoices().Item(1)

    def taking_input():
        speak.Speak("Enter your account number : ")
        Account_Number = int(input("Enter your account number : "))

        cursor = connection.cursor()
        cursor.execute("SELECT * FROM accounts WHERE account_number = %s", (Account_Number,))
        account_data = cursor.fetchone()

        if account_data:  
            Account_Number, atm_pin, account_type, phone_numbers, account_balance = account_data

            speak.Speak("Enter your atm pin : ")
            Atm_Pin = int(input("Enter your atm pin : "))

            if Atm_Pin == atm_pin:
                speak.Speak("Enter your account type : ")
                Account_Type = input("Enter your account type : ")

                if Account_Type == account_type:
                    speak.Speak("Enter your mobile number : ")
                    Mobile_Number = int(input("Enter your Mobile Number : "))

                    if Mobile_Number == phone_numbers:
                        print(account_balance)
                        speak.Speak("Enter the amount you want to withdraw")
                        Amount = int(input("Enter the amount you want to withdraw :"))

                        if account_balance > 0:
                            speak.Speak("Enter your otp ")
                            otp = int(input("Enter your otp : "))

                            if otp == a:
                                if (account_balance - Amount) > 0:
                                    cursor.execute("UPDATE accounts SET account_balance = account_balance - %s WHERE account_number = %s", (Amount, Account_Number))
                                    connection.commit() 

                                    updated_balance = account_balance - Amount
                                    speak.Speak(f"You have Credited by {Amount} and your updated Balance is {updated_balance}")
                                    print(f"You have Credited by {Amount} and your updated Balance is {updated_balance}")
                                else:
                                    speak.Speak("There is insufficient amount in Your Account")
                                    print("There is insufficient amount in Your Account")
                            else:
                                speak.Speak("You have Entered Wrong otp ")
                                print("You have Entered Wrong otp")
                        else:
                            speak.Speak("There is insufficient amount in Your Account")
                            print("There is insufficient amount in Your Account")
                    else:
                        speak.Speak("You have Enter wrong mobile number ")
                        print("You have Enter wrong mobile number ")
                else:
                    speak.Speak("You have Enter wrong Account number ")
                    print("You have Enter wrong Account number ")
        else:
            speak.Speak("Account not found")
            print("Account not found")

    taking_input()

except Error as e:
    print(f"An error occurred: {e}")

finally:
    if connection.is_connected():
        connection.quit() 