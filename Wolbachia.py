# -*- coding: utf-8 -*-
import xlsxwriter,math
xlsnum=0
workbook = xlsxwriter.Workbook('Wolbachia.xlsx')
worksheet = workbook.add_worksheet()
alar=[.5,.2,.46,.78]
'''i=0
while i<=1:
    alar.append(float(i))
    print(i)
    i=float(i)+float(10**-2)
'''#pler=[]
bler=[.8,1,0,.2,.5,0,0.1]
element=0
genNum=input("Nesil Sayisi:")
alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ"
def xlscol(num):
    if num<len(alphabet):
        return alphabet[num]
    else:
        for k in range(len(alphabet)):
            for i in range(len(alphabet)):
                if num<len(alphabet)*(i+2):
                    return str(alphabet[i])+str(alphabet[num-((i+1)*len(alphabet))])
                '''elif num<len(alphabet)*(i+2)+(k+2)*len(alphabet):
                    return str(alphabet[k])+str(alphabet[i])+str(alphabet[num-((i+1)*len(alphabet))])'''
while True:
    if element==len(alar):
        break;
    data = []
    data2 = []
    a=alar[element]
    p=.90
    #b=bler[element]
    b=1-a
    def f(gen):
        if gen==0:
            return float(a)
        else:
            return "%.53f" %((float(p)*n1)/(1-(p*n1)+((p**2)*(n1)**2)))
    def f2(gen):
        if gen==0:
            return float(a)
        elif gen==1:
            return (float(a)*float(p))/(1-(float(p)*float(b))+((float(p)**2)*float(a)*float(b)))
        else:
            return "%.53f" %((float(p)*n1)/(1-(p*n1)+((p**2)*(n1)**2)))
    for i in range(genNum+1):
        print(str(i)+". Nesil :"+str(f(i)))
        data.append([str(i),float(f(i))])
        n1=float(f(i))
    for i in range(genNum+1):
        print(str(i)+". Nesil (f2):"+str(f2(i)))
        data2.append([str(i),float(f2(i))])
        n1=float(f2(i))
    row = 1
    col = (xlsnum*4)+1
    worksheet.write(row, col, 'a:'+str(a))
    worksheet.write(row, col+1, 'p:'+str(p))
    worksheet.write(row, col+2, 'b=a')
    row+=1
    worksheet.write(row, col, 'Nesil')
    worksheet.write(row, col+1, 'f(WolbachiaFemale)')
    row+=1
    for gen, value in (data):
        worksheet.write(row, col,gen)
        worksheet.write(row, col + 1, value)
        row += 1
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({'values': '=Sheet1!$'+xlscol(col+1)+'$4:$'+xlscol(col+1)+'$'+str(4+genNum)})
    chart.set_title({'name': 'a:'+str(a)+' p:'+str(p)+' b=a'})
    worksheet.insert_chart(str(xlscol((len(alar)*4)+1))+str((element*15)+1), chart)
    row+=2
    worksheet.write(row, col, 'a:'+str(a))
    worksheet.write(row, col+1, 'p:'+str(p))
    worksheet.write(row, col+2, 'b:'+str(b))
    row+=1
    worksheet.write(row, col, 'Nesil')
    worksheet.write(row, col+1, 'f(WolbachiaFemale)')
    row+=1
    for gen, value in (data2):
        worksheet.write(row, col,gen)
        worksheet.write(row, col + 1, value)
        row += 1
    chart = workbook.add_chart({'type': 'line'})
    chart.add_series({'values': '=Sheet1!$'+xlscol(col+1)+'$'+str(4+genNum+5)+':$'+xlscol(col+1)+'$'+str(4+(genNum*2)+5)})
    chart.set_title({'name': 'a:'+str(a)+' p:'+str(p)+' b:'+str(b)})
    worksheet.insert_chart(str(xlscol((len(alar)*4)+10))+str((element*15)+1), chart)
    xlsnum+=1
    element+=1
workbook.close()
