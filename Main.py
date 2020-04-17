Bank_name = input("Enter your bank name: ")
Bank_name = Bank_name.upper()
print(Bank_name)

if Bank_name == "HDFC":
    print("Calling HDFC Bank script")
    exec(open("./HDFC_Bank_statement_PDF.py").read())
elif Bank_name == "ICICI":
    print("Calling ICICI Bank script")
    exec(open("./ICICI_Bank_statement_PDF.py").read())
else:
    print("This Bank passer is not available.")