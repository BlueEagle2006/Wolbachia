import xlsxwriter,math
import plotly.express as px
import pandas as pd
def xlscol(num): # The function that turns column numbers into xlsx understandable format [for example 0 -> A and 55 -> BD]
    alphabet="ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    if num<len(alphabet):
        return alphabet[num]
    else:
        for k in range(len(alphabet)):
            for i in range(len(alphabet)):
                if num<len(alphabet)*(i+2):
                    return str(alphabet[i])+str(alphabet[num-((i+1)*len(alphabet))])
def f(x,a,d,w,b,fn): # The Wolbachia Percentage Calculator Function
    x  = float(x)
    a  = float(a)
    d  = float(d)
    w  = float(w)
    b  = float(b)
    fn = float(fn)
    if x==0:
        return a
    elif x==1:
        return (d*a) / (1-w*b+d*w*a*b)
    else:
        return (d*fn) / (1-w*fn+d*w*(fn**2))

#=====The Values Are Here======
list_a = [0.45,0.45]
list_b = [0.07,0.1]
list_d = [0.74]
list_w = [1]
max_x  = 30
#==============================

for i in range(len(list_a)):
    a  = list_a[i]
    b  = list_b[i]
    for j in range(len(list_d)):
        d  = list_d[j]
        w  = list_w[j]
        fn = 0
        xlsx_data = []
        xlsx_data_for_graph = []
        id_data   = []
        # Calculation of the "Wolbachia Percentage" until given generation value
        for x in range(max_x):
            #print(str(x)+". Generation : "+str(f(x,a,d,w,b,fn)))
            xlsx_data.append([x,f(x,a,d,w,b,fn)])
            xlsx_data_for_graph.append(f(x,a,d,w,b,fn))
            id_data.append(x)
            fn = f(x,a,d,w,b,fn)
        # Creation of excel file
        workbook = xlsxwriter.Workbook("a:"+str(a)+", d:"+str(d)+", w:"+str(w)+", b:"+str(b)+".xlsx")
        worksheet = workbook.add_worksheet()
        row = 1
        col = 1
        worksheet.write(row, col, 'Generation')
        worksheet.write(row, col+1, 'Wolbachia')
        row+=1
        for x, fx in (xlsx_data):
            worksheet.write(row, col,x)
            worksheet.write(row, col + 1, fx)
            row += 1
        workbook.close()
        # Creation of the graph
        #DO NOT DELETE!!!
        df = pd.DataFrame(dict(
            Generation = id_data,
            Wolbachia  = xlsx_data_for_graph
        ))
        fig = px.line(df, x="Generation", y="Wolbachia", title="a:"+str(a)+", d:"+str(d)+", w:"+str(w)+", b:"+str(b))
        fig.write_image("a:"+str(a)+", d:"+str(d)+", w:"+str(w)+", b:"+str(b)+".png")
