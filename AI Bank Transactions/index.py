import pandas as pd
import pyttsx3
import speech_recognition as sr
from datetime import datetime, timedelta

# Load the banking dataset
df = pd.read_excel("transactions.xlsx")

# Initialize pyttsx3 engine
engine = pyttsx3.init()

# Function to speak the response
def speak(text):
    engine.say(text)
    engine.runAndWait()

# Function to listen to the user's voice input
def listen():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        audio = recognizer.listen(source)
    try:
        text = recognizer.recognize_google(audio)
        print(f"You said: {text}")
        return text.lower()
    except sr.UnknownValueError:
        print("Sorry, I couldn't understand.")
        return ""

# 1. What was my last transaction amount and date?
def get_last_transaction():
    last_transaction = df.iloc[-1]
    response = f"Your last transaction was {last_transaction['Transaction Narrative']} on {last_transaction['Date']}, amount {last_transaction['Debit'] or last_transaction['Credit']}."
    speak(response)

# 2. How much money did I spend in the last 7 days?
def get_spend_last_7_days():
    seven_days_ago = datetime.now() - timedelta(days=8)
    recent_transactions = df[df['Date'] >= seven_days_ago]
    total_spent = recent_transactions['Debit'].sum()
    speak(f"You spent a total of {total_spent} dollars in the last 7 days.")

# 3. What is my current account balance?
def get_account_balance():
    balance = df['Credit'].sum() - df['Debit'].sum()
    speak(f"Your current balance is {balance} dollars.")

# 4. Tell me my last 3 transactions.
def get_last_n_transactions(n=3):
    last_n = df.tail(n)
    response = "Here are your last " + str(n) + " transactions: "
    for _, row in last_n.iterrows():
        response += f"{row['Transaction Narrative']} on {row['Date']}, amount {row['Debit'] or row['Credit']}. "
    speak(response)

# 5. Did I receive any salary this month?
def check_salary():
    this_month = datetime.now().month
    salary_transactions = df[(df['Date'].dt.month == this_month) & df['Transaction Narrative'].str.contains("Salary", case=False)]
    if not salary_transactions.empty:
        total_salary = salary_transactions['Credit'].sum()
        speak(f"Your salary this month was {total_salary} dollars.")
    else:
        speak("You haven't received any salary this month.")


# 6. How much did I spend on shopping?
def shopping_expenses():
    shopping_transactions = df[df['Transaction Narrative'].str.contains("Shopping", case=False)]
    total_shopping = shopping_transactions['Debit'].sum()
    speak(f"You spent a total of {total_shopping} dollars on shopping.")

# 7. When was the last time I paid my electricity bill?
def last_electricity_bill():
    electricity_transactions = df[df['Transaction Narrative'].str.contains("Electricity", case=False)]
    if not electricity_transactions.empty:
        last_electricity = electricity_transactions.iloc[-1]
        speak(f"Your last electricity bill payment was on {last_electricity['Date']}.")
    else:
        speak("You haven't paid an electricity bill recently.")

# 8. How much did I transfer to John this month?
def transfer_to_john():
    this_month = datetime.now().month
    john_transactions = df[(df['Date'].dt.month == this_month) & df['Transaction Narrative'].str.contains("John", case=False)]
    total_transferred = john_transactions['Debit'].sum()
    speak(f"You transferred {total_transferred} dollars to John this month.")

# 9. What was my last grocery expense?
def last_grocery_expense():
    grocery_transactions = df[df['Transaction Narrative'].str.contains("Grocery", case=False)]
    if not grocery_transactions.empty:
        last_grocery = grocery_transactions.iloc[-1]
        speak(f"Your last grocery expense was {last_grocery['Debit']} on {last_grocery['Date']}.")
    else:
        speak("You haven't made any grocery purchases recently.")

# Main loop to listen for commands
while True:
    command = listen()
    if "last transaction" in command:
        get_last_transaction()
    elif "last 7 days" in command:
        get_spend_last_7_days()
    elif "account balance" in command:
        get_account_balance()
    elif "last three transactions" in command:
        get_last_n_transactions(3)
    elif "salary" in command:
        check_salary()
    elif "shopping" in command:
        shopping_expenses()
    elif "electricity bill" in command:
        last_electricity_bill()
    elif "transfer to john" in command:
        transfer_to_john()
    elif "grocery" in command:
        last_grocery_expense()
    elif "exit" in command:
        speak("Goodbye!")
        break