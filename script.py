import turtle as tr
import random
import openpyxl as xlp

lwb=xlp.load_workbook("Dataset.xlsx")
lwk=lwb.get_sheet_by_name("Your_sheet")

mxr=lwk.max_row

tr.speed(0)
tr.color("white")
tr.right(180)
tr.forward(500)
tr.right(90)
tr.forward(400)
tr.right(90)

ln=950
br=800

l=[]
for zm in range(2,mxr+1):
    l.append(lwk["Your_column"+str(zm)].value)

area=ln*br
avar=area/len(l)
avar=int(avar)

fcs=[]
for i in range(1, avar + 1):
    if avar % i == 0:
        if i not in fcs:
            fcs.append(i)
            fcs.append(avar//i)

x=fcs[len(fcs)-1]
y=fcs[len(fcs)-2]

print(x)
print(y)

ppl=950//x
ppc=800//y
crp=0
ccp=0
for qw in range(0,len(l)):
    if crp==ppl:
        tr.right(90)
        tr.forward(y)
        tr.right(90)
        tr.forward(x*ppl)
        tr.right(180)
        crp=0
        ccp+=1
    else:
        tr.color(l[qw])
        tr.fillcolor(l[qw])
        tr.begin_fill()
        tr.forward(x)
        tr.right(90)
        tr.forward(y)
        tr.right(90)
        tr.forward(x)
        tr.right(90)
        tr.forward(y)
        tr.right(90)
        tr.forward(x)
        crp=crp+1
        if crp==ppl-1 and ccp==ppc-1:
            tr.forward(x)
            tr.right(90)
            tr.forward(y)
            tr.right(90)
            tr.forward(x)
            tr.right(90)
            tr.forward(y)
            tr.right(90)
            tr.forward(x)
        tr.end_fill()

tr.hideturtle()
tr.done()
