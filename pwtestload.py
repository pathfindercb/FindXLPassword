import msoffcrypto
# opening the file in read mode
my_file = open("pwtest.txt", "r")
  
# reading the file
data = my_file.read()
  
# replacing end splitting the text 
# when newline ('\n') is seen.
datalist = data.split("\n")
i = 0
for testPW in datalist:
    i = i + 1
    encrypted =open("pwtest.xlsx", "rb")
    file = msoffcrypto.OfficeFile(encrypted)   
    try:
        file.load_key(password=testPW)   
        with open("decrypted.xlsx", "wb") as f:
            file.decrypt(f)
        encrypted.close()    
    except:
        print (str(i) + " " + testPW)
    else:
        print("Found! " + testPW)
        break

    
my_file.close()