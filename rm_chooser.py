import xlrd

class UserDetails():
    def __init__(self):
        self.name = ""
        self.number = ""
        self.email = ""
        self.gender = ""
        self.food = ""
        self.o_food = "y"
        self.cooking = ""
        self.party = ""
        self.smoke = ""
        self.drink = ""
        self.o_drink = ""
        self.o_smoke = ""
        self.flatmate = ""
        self.roommate = ""
        self.comment = ""

user = UserDetails()
temp = UserDetails()


def accept_details():
    print("Beginning Questionaire for Roommate Finder. Please answer the following questions")
    user.gender = raw_input("Gender? Enter 'm' or 'f'")
    print("For all the following questions, please answer with 'y' for YES or 'n' for NO.\n Do NOT give any other answer.\n Do not use capital letters")  
    user.food = raw_input("Are you a vegetarian? ")
    if user.food == "y":
        user.o_food = raw_input("Is it okay if roommate is non-vegetarian? ")
    user.cooking = raw_input("Can you cook? ")
    user.party = raw_input("Do you have a problem with a party person? ")
    user.smoke = raw_input("Do you smoke? ")
    user.o_smoke = raw_input("Is it okay if roommate smokes?")
    user.drink = raw_input("Do you drink? ")
    user.o_drink = raw_input("Is it okay if roommate drinks?")
    user.flatmate = raw_input("Are you looking for a flatmate?(Different rooms, same flat)")
    user.roommate = raw_input("Are you looking for a roommate?(Same rooms)")
    print("Thank you for answering all those....detecting people with the same preferences as you")
    
def compare(path,no_fields):
    book = xlrd.open_workbook(path)
    f_sheet = book.sheet_by_index(0)
    num = 1
    flag = 1
    for i in range(1,no_fields):
        flag = 1
        temp.name = f_sheet.cell(i,1).value
        temp.number = f_sheet.cell(i,2).value
        temp.email = f_sheet.cell(i,3).value
        temp.gender = f_sheet.cell(i,4).value
        if ((user.gender == "m") and (temp.gender == "Female")) or ((user.gender == 'f') and (temp.gender == 'Male')):
            flag = 0
            
        temp.food = f_sheet.cell(i,10).value
        if (user.o_food == "n") and (temp.food == "Non-Veg"):
            flag = 0
            
        temp.cooking = f_sheet.cell(i,12).value
        temp.party = f_sheet.cell(i,13).value
        if user.party == "y" and temp.party == "Yes":
            flag = 0
        temp.smoke = f_sheet.cell(i,14).value
        temp.drink = f_sheet.cell(i,15).value
        temp.o_drink = f_sheet.cell(i,16).value
        temp.o_smoke = f_sheet.cell(i,17).value
        if ((user.o_smoke == "n") and (temp.smoke == "Yes" or temp.smoke == "Maybe")) or ((user.smoke == "y") and (temp.smoke == "No, I am not.")):
            flag = 0
        if ((user.o_drink == "n") and (temp.drink == "Yes" or temp.drink == "Maybe")) or ((user.drink == "y") and (temp.drink == "No, I am not.")):
            flag = 0
        temp.flatmate = f_sheet.cell(i,23).value
        temp.roommate = f_sheet.cell(i,24).value
        temp.comment = f_sheet.cell(i,25).value
        if flag == 1:
            display(num)
            num = num + 1
    if num == 1:
        print "Sorry, we do not have mythical beings present"

def display(num):
    print "Details of person number ",num
    print "Name : ",temp.name
    print "Number : ",temp.number
    print "Email : ", temp.email
    print "Food Preference :", temp.food
    print "Can cook :",temp.cooking
    print "Does party : ",temp.party
    print "Does smoke : ",temp.smoke
    print "Does drink : ",temp.drink
    print "Looking for flatmate : ",temp.flatmate
    print "Looing for roomamte : ",temp.roommate
    print "Any comments : ",temp.comment
    print ""
    print ""




print("Please download the excel sheel from the group, and place it in the same folder as this file. \nThen, rename it as 'input.xlsx' before running this program")
print("Note : This does not match people with opposite gender")
accept_details()
path = "input.xlsx"
no_fields = input("Enter the number of Rows in the file : ")
compare(path,no_fields)



