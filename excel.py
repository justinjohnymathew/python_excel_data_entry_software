import openpyxl
import os
print("EXCEL DATA ENTERING SOFTWARE")
wb=openpyxl.load_workbook('git_workshop.xlsx')
sheet=wb['Sheet1']
sheet['A1']='Name'
sheet['B1']='Semester'
sheet['C1']='Branch&class'
sheet['D1']='Email'
ch='Y'
i=2
os.system("clear")
while(ch.upper()=='Y'):
    name=input("Name:")
    sheet[str('A'+str(i))]=name
    semester=input("Semester(1,2,3,4,5,6,7,8):")
    sheet[str('B'+str(i))]=str('S'+str(semester))
    branch_no=input("Branch \n(1)CSE\n(2)ECE:\n")
    if branch_no==1:
        branch='CSE'
    else:
        branch="ECE"
    clas=input("Division[A OR B]:")
    sheet[str('C'+str(i))]=str(branch+'-'+clas.upper())
    email=input("EMAIL:")
    sheet[str('D'+str(i))]=str(email)
    i+=1
    wb.save('git_workshop.xlsx')
    os.system("clear")
    ch=input("Do You Want to add More Data(Y/N):")
