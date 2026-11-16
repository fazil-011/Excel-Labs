import random, csv, datetime as dt, os
random.seed(42)
D='Allfiles/Datasets/'
def w(name, header, rows):
    with open(D+name,'w',newline='',encoding='utf-8') as f:
        c=csv.writer(f); c.writerow(header); c.writerows(rows)
first=['Aarav','Diya','Rohan','Ananya','Vikram','Sneha','Karan','Priya','Arjun','Meera','Rahul','Kavya','Imran','Fatima','Suresh','Lakshmi','Nikhil','Pooja','Aditya','Ishita','Manoj','Divya','Farhan','Zoya','Harsh','Neha','Sanjay','Tanvi','Yash','Riya']
last=['Sharma','Patel','Mehta','Iyer','Singh','Das','Nair','Reddy','Kumar','Gupta','Khan','Joshi','Rao','Verma','Shetty','Menon','Pillai','Bose','Naidu','Hegde']
cities=['Bengaluru','Mumbai','Delhi','Hyderabad','Chennai','Pune']
def name(): return random.choice(first)+' '+random.choice(last)
def rdate(a,b): return a+dt.timedelta(days=random.randint(0,(b-a).days))
# 1 students
rows=[]
for i in range(1,101):
    n=name(); rows.append([f'S{i:03d}',n,random.choice(['Male','Female']),random.choice(['BCA','BE-CSE','BE-ECE','BSc Data Science']),random.choice([1,2,3,4,5,6]),random.choice(cities),rdate(dt.date(2002,1,1),dt.date(2006,12,31)).isoformat(),'9'+str(random.randint(100000000,999999999)),n.lower().replace(' ','.')+'@college.edu',random.randint(35,98),random.randint(35,98),random.randint(35,98)])
w('students.csv',['StudentID','StudentName','Gender','Course','Semester','City','DateOfBirth','Phone','Email','MathsMarks','ProgrammingMarks','DatabaseMarks'],rows)
# 2 employees
desg={'Sales':[('Sales Executive',32000),('Senior Sales Executive',48000),('Sales Manager',85000)],'HR':[('HR Executive',30000),('HR Manager',80000)],'IT':[('Software Engineer',60000),('Senior Software Engineer',95000),('IT Manager',140000)],'Finance':[('Accountant',38000),('Finance Analyst',55000),('Finance Manager',110000)],'Operations':[('Operations Executive',33000),('Operations Manager',90000)],'Marketing':[('Marketing Executive',34000),('Marketing Manager',88000)]}
emps=[]
for i in range(150):
    d=random.choice(list(desg)); ds,base=random.choice(desg[d])
    emps.append([f'E{1001+i}',name(),d,ds,int(base*random.uniform(0.9,1.25)/500)*500,rdate(dt.date(2012,1,1),dt.date(2024,9,30)).isoformat(),random.choice(cities)])
w('employees.csv',['EmployeeID','EmployeeName','Department','Designation','Salary','JoiningDate','Location'],emps)
# 3 sales
prods=[('Laptop','Electronics',55000,44000),('Smartphone','Electronics',22000,17500),('Headphones','Electronics',3500,2200),('Monitor','Electronics',12000,9000),('Office Chair','Furniture',7500,5200),('Standing Desk','Furniture',18000,12500),('Bookshelf','Furniture',6500,4300),('Notebook Pack','Office Supplies',250,140),('Printer Paper','Office Supplies',450,320),('Pen Set','Office Supplies',180,90),('Stapler','Office Supplies',220,130),('Whiteboard','Office Supplies',2200,1500)]
sp=[('Aarav Sharma','North'),('Diya Patel','North'),('Rohan Mehta','South'),('Ananya Iyer','South'),('Vikram Singh','East'),('Sneha Das','East'),('Karan Nair','West'),('Priya Reddy','West')]
sales=[]
for i in range(1200):
    p=random.choice(prods); s=random.choice(sp)
    sales.append([f'ORD-{10001+i}',rdate(dt.date(2024,1,1),dt.date(2024,12,31)),s[1],s[0],p[1],p[0],random.randint(1,10),p[2],p[3],random.choice([0,0,0,5,10,15])])
sales.sort(key=lambda r:(r[1],r[0]))
for n,r in enumerate(sales): r[0]=f'ORD-{10001+n}'
H=['OrderID','OrderDate','Region','Salesperson','Category','Product','Quantity','UnitPrice','UnitCost','DiscountPct']
w('sales_transactions.csv',H,[[r[0],r[1].isoformat()]+r[2:] for r in sales])
for m in range(1,13):
    w(f'monthly_sales/sales_2024_{m:02d}.csv',H,[[r[0],r[1].isoformat()]+r[2:] for r in sales if r[1].month==m])
w('salespeople.csv',['Salesperson','Manager','Team','JoinYear'],[[a,{'North':'Ravi Kapoor','South':'Latha Krishnan','East':'Subir Ghosh','West':'Meenal Desai'}[b],b+' Team',random.randint(2015,2022)] for a,b in sp])
tg=[]
for a,b in sp:
    tg.append([a]+[random.choice([80,90,100,110,120])*1000 for _ in range(12)])
w('sales_targets_wide.csv',['Salesperson','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'],tg)
# 4 performance
perf=[]
rg=random.Random(7)
for e in emps[:60]:
    t=rg.choice([400000,500000,600000,800000]); perf.append([e[0],e[1],t,int(t*rg.uniform(0.6,1.4)),rg.randint(76,100),rg.choice([2,3,3,4,4,4,5,5])])
w('employee_performance.csv',['EmployeeID','EmployeeName','SalesTarget','SalesAchieved','AttendancePct','ManagerRating'],perf)
# 6 customers raw
cr=[]
def messy(n):
    k=random.randint(0,3)
    n={0:n.upper(),1:n.lower(),2:n,3:n.swapcase() if False else n.title()}[k]
    return random.choice(['',' ','  '])+n.replace(' ',random.choice([' ','  ']))+random.choice(['',' '])
for i in range(1,81):
    n=name(); ph=str(random.randint(7000000000,9999999999))
    cr.append([f'C{i:04d}',messy(n),n.lower().replace(' ','.')+random.choice(['@gmail.com','@yahoo.com','@outlook.com']) if random.random()<.5 else n.upper().replace(' ','.')+'@GMAIL.COM',random.choice([ph,ph[:5]+' '+ph[5:],'+91-'+ph]),random.choice(cities),rdate(dt.date(2021,1,1),dt.date(2024,10,31)).isoformat(),f'{random.choice(["ELEC","FURN","OFFC"])}-{random.randint(1,9999):04d}-{random.choice(["IN","US","UK"])}'])
w('customers_raw.csv',['CustomerID','FullName','Email','Phone','City','SignupDate','ProductCode'],cr)
# 7 messy
clean=[]
for i in range(1,151):
    n=name(); clean.append([f'M{i:04d}',n,random.choice(cities),rdate(dt.date(2024,1,1),dt.date(2024,12,31)),random.choice(['Electronics','Furniture','Office Supplies']),random.randint(1,10),random.randint(500,50000),n.lower().replace(' ','.')+'@mail.com','9'+str(random.randint(100000000,999999999))])
w('customer_sales_clean_reference.csv',['OrderID','CustomerName','City','OrderDate','Category','Quantity','Amount','Email','Phone'],[[r[0],r[1],r[2],r[3].isoformat()]+r[4:] for r in clean])
mess=[]; stats=dict(blank_city=0,blank_qty=0,badname=0,baddate=0,badamt=0,variant=0)
cv={'Bengaluru':['Bangalore','bengaluru','BENGALURU'],'Mumbai':['Bombay','mumbai','MUMBAI '],'Delhi':['New Delhi','delhi','DELHI'],'Hyderabad':['hyderabad','Hyd'],'Chennai':['chennai','Madras'],'Pune':['pune','PUNE']}
for r in clean:
    r=list(r); r[3]=r[3].isoformat()
    if random.random()<.3: r[1]=random.choice([' '+r[1].upper(),r[1].lower()+'  ',r[1].replace(' ','  ')]); stats['badname']+=1
    if random.random()<.35: r[2]=random.choice(cv[r[2]]); stats['variant']+=1
    if random.random()<.06: r[2]=''; stats['blank_city']+=1
    if random.random()<.05: r[5]=''; stats['blank_qty']+=1
    if random.random()<.07: r[3]=random.choice(['31/02/2024','N/A','2024-13-05','tomorrow','00-00-0000']); stats['baddate']+=1
    elif random.random()<.2:
        d=dt.date.fromisoformat(r[3]); r[3]=d.strftime('%d-%b-%Y')
    if random.random()<.08: r[6]=random.choice(['Rs. '+str(r[6]),str(r[6])+' INR','five thousand','$'+str(r[6])]); stats['badamt']+=1
    mess.append(r)
dups=random.sample(mess,10); mess+= [list(d) for d in dups]; random.shuffle(mess)
w('customer_sales_messy.csv',['OrderID','CustomerName','City','OrderDate','Category','Quantity','Amount','Email','Phone'],mess)
print('messy rows',len(mess),stats,'dups 10')
# 8 inventory
inv=[]
sup=['TechSource Pvt Ltd','OfficeMart','FurniWorld','BrightSupplies']
for i,p in enumerate(prods*3+[('USB Cable','Electronics',299,150)]*0):
    pass
names=[(p[0]+(' '+v if v else ''),p[1],p[2]) for p in prods for v in ['','Pro','Lite']]
for i,(n,c,pr) in enumerate(names[:36]):
    inv.append([f'P{101+i}',n,c,random.choice(sup),random.randint(0,200),random.choice([10,20,30,50]),int(pr*random.uniform(.8,1.3))])
w('inventory.csv',['ProductID','ProductName','Category','Supplier','StockQty','ReorderLevel','UnitPrice'],inv)
# 13 HR
roles={'Sales':['Sales Executive','Sales Representative','Manager'],'Research & Development':['Research Scientist','Laboratory Technician','Manager'],'Human Resources':['HR Executive','Recruiter','Manager']}
hr=[]
for i in range(1,301):
    d=random.choices(list(roles),[.45,.45,.10])[0]; inc=random.randint(2500,19000); ot=random.choice(['Yes','No','No']); js=random.randint(1,4)
    p=0.08+(0.12 if ot=='Yes' else 0)+(0.08 if js<=2 else 0)+(0.08 if inc<4500 else 0)
    hr.append([i,random.randint(21,58),'Yes' if random.random()<p else 'No',d,random.choice(roles[d]),random.choice(['Male','Female']),inc,random.randint(0,25),js,random.choice([3,3,3,4]),ot,random.randint(1,29)])
w('hr_sample.csv',['EmployeeNumber','Age','Attrition','Department','JobRole','Gender','MonthlyIncome','YearsAtCompany','JobSatisfaction','PerformanceRating','OverTime','DistanceFromHome'],hr)
# 14 finance
fin=[]; accts=[('Product Sales','Revenue',900000),('Service Revenue','Revenue',300000),('Cost of Goods Sold','COGS',520000),('Salaries','Opex',280000),('Rent','Opex',60000),('Marketing','Opex',70000),('Utilities','Opex',18000),('Other Expenses','Opex',25000)]
for m in range(1,13):
    for a,t,b in accts:
        bud=int(b*(1+0.01*m)/1000)*1000; fin.append([dt.date(2024,m,1).isoformat(),a,t,bud,int(bud*random.uniform(.88,1.14))])
w('finance_monthly.csv',['Month','Account','AccountType','Budget','Actual'],fin)
# 15 ecommerce
ep=[('85123A','White Hanging Heart Holder',2.95),('71053','Metal Lantern',3.39),('84406B','Cream Cupid Hearts Coat Hanger',2.75),('22423','Regency Cakestand 3 Tier',12.75),('21730','Glass Star Frosted Light',4.25),('47566','Party Bunting',4.95),('20725','Lunch Bag Red Retrospot',1.65),('22720','Set of 3 Cake Tins',4.95),('23084','Rabbit Night Light',2.08),('22197','Popcorn Holder',0.85),('84879','Assorted Colour Bird Ornament',1.69),('22086','Paper Chain Kit 50s Christmas',2.55)]
cust=[(10000+i,random.choices(['United Kingdom','Germany','France','India','USA'],[.5,.15,.12,.12,.11])[0]) for i in range(150)]
ec=[]; inv_no=500000
for _ in range(500):
    inv_no+=1; c=random.choice(cust); d=rdate(dt.date(2024,1,1),dt.date(2024,12,31))
    for p in random.sample(ep,random.randint(1,5)):
        ec.append([inv_no,d.isoformat(),c[0],c[1],p[0],p[1],random.randint(1,24),p[2]])
ec.sort(key=lambda r:r[1])
w('ecommerce_orders.csv',['InvoiceNo','InvoiceDate','CustomerID','Country','StockCode','Description','Quantity','UnitPrice'],ec)
# stats
tot=sum(r[6]*r[7]*(1-r[9]/100) for r in sales); cost=sum(r[6]*r[8] for r in sales)
print('sales rows',len(sales),'rev',round(tot,2),'profit',round(tot-cost,2),'AOV',round(tot/len(sales),2))
from collections import defaultdict
for key,idx in [('region',2),('sp',3),('cat',4),('prod',5)]:
    a=defaultdict(float);b=defaultdict(float)
    for r in sales: a[r[idx]]+=r[6]*r[7]*(1-r[9]/100); b[r[idx]]+=r[6]*r[7]*(1-r[9]/100)-r[6]*r[8]
    print(key,{k:(round(a[k]),round(b[k])) for k in a})
mm=defaultdict(float)
for r in sales: mm[r[1].month]+=r[6]*r[7]*(1-r[9]/100)
print('month',{k:round(v) for k,v in sorted(mm.items())})
print('emp salary sum',sum(e[4] for e in emps),'IT',sum(1 for e in emps if e[2]=='IT'))
print('perf achieved>=target',sum(1 for p in perf if p[3]>=p[2]),'of',len(perf))
print('HR attr',sum(1 for h in hr if h[2]=='Yes'),'of 300; avg income',round(sum(h[6] for h in hr)/300))
a=sum(f[3] for f in fin);b=sum(f[4] for f in fin);print('fin budget',a,'actual',b)
r_=sum(f[4] for f in fin if f[2]=='Revenue');c_=sum(f[4] for f in fin if f[2]=='COGS');o_=sum(f[4] for f in fin if f[2]=='Opex');print('fin rev',r_,'cogs',c_,'opex',o_,'np',r_-c_-o_)
print('ecom rows',len(ec),'rev',round(sum(r[6]*r[7] for r in ec),2),'invoices',len({r[0] for r in ec}))
print('students avg maths',round(sum(r[9] for r in [[*x] for x in rows]) ,0) if False else '')
