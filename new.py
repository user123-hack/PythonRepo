var = int(input("Enter Your Choice\n1. Reverse Number\n2. Palindrome\n3. Forming a number\n4. Exit\n"))
if(var == 1):
        number = int(input("enter the number \t"))
        rev = 0
        while(number > 0):
                remainder = number % 10  
                rev = (rev * 10) + remainder  
                number = number // 10  

        print("The reverse number is : {}".format(rev))

if(var == 2):
        s = input("Enter the String\n")
        ear = s[::-1]
        if(s == ear):
                print("Is Palindrome")
        else:
                print("Isn't")        

if(var == 3):
        num = int(input("Enter the number"))
        last_dig = num % 10
        first = num // 10
        first = first / 10   


        print(first)