import requests,xlsxwriter,time,random,os,re,sys,glob,csv
from bs4 import BeautifulSoup
from xlsxwriter.workbook import Workbook

option=int(input("Choose Mode [ 1. Single Roll Number | 2. Multiple Roll Number | 3. Import From File ] [ Enter 1 | 2 | 3 ] : "))
if option==1:
        a=int(input("Enter Roll Number\t\t\t\t\t\t\t\t\t\t\t  : "))
        b=a+1
        c=b-a
elif option==2:
        a=int(input("Enter Starting Roll Number\t\t\t\t\t\t\t\t\t\t  : "))
        b=int(input("Enter Ending Roll Number\t\t\t\t\t\t\t\t\t\t  : "))+1
        c=b-a
elif option==3:
        txtfile=input("Enter Roll Number Text File Name [ Type Excluding .TXT ]\t\t\t\t\t\t  : ")
        if os.path.isfile(txtfile+'.txt'):
                lines = tuple(open(txtfile+'.txt', 'r'))
                with open(txtfile+'.txt') as f:
                        length=len(f.readlines())
                a=0
                b=length
                c=b
                roll_number=0
        else:
                print("\n"+txtfile+".txt doesn't exist! Please enter correct file name.")
                sys.exit(0)
else:
        print("Incorrect choice")
        sys.exit(0)


sem_no=int(input("Enter Semester Number [ 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 ]\t\t\t\t\t\t\t  : "))
if sem_no<1 or sem_no>8:
        print("\nEnter a valid semester number.")
        sys.exit(0)        
fileextension=input("Enter Output File Type [ .CSV | .XLSX ]\t\t\t\t\t\t\t\t\t  : ")
fileextension=re.sub('[\s+]', '', fileextension)
fileextension=fileextension.upper()
if fileextension==".CSV" or fileextension==".XLSX":
        p=1
else:
        print("\nSorry, your chosen file type isn't supported. Choose any of the two file type supported by this program")
        sys.exit(0)
filename=input("Enter Output "+fileextension.upper()+" File Name\t\t\t\t\t\t\t\t\t\t  : ")
filename=re.sub('[\s+]', '', filename)
if filename=="":
        print("Blank file name! Enter proper file name.")
        sys.exit(0)
outfile = open(filename+".csv", 'w')
res_str=""
first_run=0
counter=0
p=0
time.sleep(1)
sys.stdout.write("\nConnecting to WBUT Server...")


if (sem_no%2):
        url="http://wbutech.net/show-result1516.php"
        ref_url="http://wbutech.net/result_odd1516.php"
else:
        url="http://wbutech.net/show-result_even.php"
        ref_url="http://wbutech.net/result_even.php"
for roll_number in range(a,b):
        if option==3:
                roll_number=lines[roll_number]
        time.sleep(random.randint(4,12))
        payload={
                'semno':sem_no,
                'rectype':'1',
                'rollno':roll_number
        }
        headers={
                'Host':'wbutech.net',
                'Origin':'http://wbutech.net',
                'Referer':'http://wbutech.net/result_even.php',
                'User-Agent':'Mozilla/5.0 (Windows NT 10.0; <64-bit tags>) AppleWebKit/<WebKit Rev> (KHTML, like Gecko) Chrome/<Chrome Rev> Safari/<WebKit Rev> Edge/<EdgeHTML Rev>.<Windows Build>'
        }
        r = requests.post(url, data=payload,headers=headers)
        if(r.url==ref_url):
                continue
        soup=BeautifulSoup(r.text, "html.parser")
        lblContent=soup.find(id="lblContent")
        data = []
        student_details = []
        sem_marks = []
        table = lblContent.find('table')
        table_body = table.find('tbody')
        rows = table_body.find_all('tr')
        for row in rows:
            cols = row.find_all('th')
            cols = [ele.text.strip() for ele in cols]
            student_details.append([ele for ele in cols if ele])
        table = table.find_next('table')
        table_body = table.find('tbody')
        rows = table_body.find_all('tr')
        for row in rows:
            cols = row.find_all('td')
            cols = [ele.text.strip() for ele in cols]
            data.append([ele for ele in cols if ele])
        table = table.find_next('table')
        table_body = table.find('tbody')
        rows = table_body.find_all('tr')
        for row in rows:
            cols = row.find_all('td')
            cols = [ele.text.strip() for ele in cols]
            sem_marks.append([ele for ele in cols if ele])

        if(first_run==0):
                first_run=1
                dept=student_details[0][0]
                dept=dept.split()
                print("\n\n"+sem_marks[1][0][22:]+" | "+ dept[5][1:-1] + " | "+dept[2]+" YEAR\n")
                res_str=("ROLL,NAME,")
                for i in range(1,len(data)-1):
                        res_str+=data[i][0]+","
                res_str+=("SGPA-O")
                outfile.write(res_str)
        res_str="\n"
        student_name=student_details[1][0][7:]
        student_roll=student_details[1][1][11:]
        res_str+=student_roll+","+student_name+","
        for i in range(1,len(data)-1):
                res_str+=data[i][2]+","
        sgpa_odd=sem_marks[0][0][42:]
        res_str+=(sgpa_odd)
        outfile.write(res_str)
        print("Retrieved "+str(counter+1)+" out of "+str(c)+" roll numbers successfully.")
        counter+=1
outfile.close()
if fileextension==".xlsx" or fileextension==".XLSX":
    workbook = Workbook(filename+fileextension, {'strings_to_numbers':  True})
    c = workbook.add_format({'font_name': 'Segoe UI', 'align': 'center'})
    b = workbook.add_format({'font_name': 'Segoe UI', 'align': 'center','bold': True})
    g = workbook.add_format({'color': 'green', 'font_name': 'Segoe UI', 'align': 'center','bold': True})
    n = workbook.add_format({'color': 'red', 'font_name': 'Segoe UI', 'align': 'center','bold': True})
    o = workbook.add_format({'color': 'orange', 'font_name': 'Segoe UI', 'align': 'center','bold': True})
    worksheet = workbook.add_worksheet('Makaut WB Result')
    worksheet.set_column('A:A', 16)
    worksheet.set_column('B:B', 28)
    for csvfile in glob.glob(os.path.join('.', '*.csv')):
        with open(csvfile, 'r') as f:
            reader = csv.reader(f)
            for r, row in enumerate(reader):
               for c, col in enumerate(row):
                       if r==0:
                               worksheet.write(r,c,col,b)
                       else:
                               if col=="P " or col=="P":
                                       worksheet.write(r,c,col,g)
                               elif col=="X":
                                       worksheet.write(r,c,col,n)
                               elif col=="XP":
                                       worksheet.write(r,c,col,o)
                               else:
                                       worksheet.write(r, c, col,c)
os.remove(filename+".csv")
print("\nGenerated "+str(counter)+" row(s) successfully in "+filename+fileextension)
workbook.close()
print("\nOpening "+filename+fileextension+" ...")    
os.startfile(filename+fileextension)
