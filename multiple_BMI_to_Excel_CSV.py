import xlsxwriter as exc
import PySimpleGUI as sg
import convert as con


sg.theme('DarkAmber')
def works(data,workbook,worksheet,file):
    judul = ["Name","Height (cm)","Weight (cm)","BMI","Status"]
    worksheet.write(0,0,judul[0])
    worksheet.write(0,1,judul[1])
    worksheet.write(0,2,judul[2])
    worksheet.write(0,3,judul[3])
    worksheet.write(0,4,judul[4])

    row=1
    col=0
    for nama,tb,bb,bmi,stat in data:
        worksheet.write(row,col,nama)
        worksheet.write(row,col+1,tb)
        worksheet.write(row,col+2,bb)
        worksheet.write(row,col+3,bmi)
        worksheet.write(row,col+4,stat)
        row+=1

    workbook.close()

    layout = [
        [sg.Text('\n'*4)],
        [sg.Text('\n\t       Convert file to CSV ?  ',)],
        [sg.Text('\n\n\t '),sg.Submit("Yes",size=(8,2)), sg.Cancel('No',size=(8,2))]
    ]
    window = sg.Window('BMI to Excel',layout, size = (350,400))
    event,values = window.read()
    if event=='Yes':
        con.convert(file)
        window.close()
    else:
        window.close()
    sg.popup("Done")
 
def inp(n,workbook,worksheet,file):
    data=[0]*n
    count=0
    for i in range(n):
        layout = [
            [sg.Text("\n")],
            [sg.Text("Data"),sg.Text(i+1),sg.Text('\n')],
            [sg.Text('Name         : '),sg.Input(),sg.Text('\n')],
            [sg.Text('Height (cm) : '),sg.Input(),sg.Text('\n')],
            [sg.Text('Weight (kg) : '),sg.Input(),sg.Text('\n')],
            [sg.Text('\n'*2)],
            [sg.Text('\t'),sg.Submit(size=(8,2)), sg.Cancel('Exit',size=(8,2))]

        ]
        window = sg.Window('BMI to Excel',layout, size = (350,400))
        event,values = window.read() 
        window.close()
        if event==sg.WIN_CLOSED or event=="Exit":
            window.close() 
            return inp_awal()
            break;
        elif event=="Submit":
            count+=1
            data[i]=[]*len(values)
            for j in range(len(values)):
                if j==1 or j==2:
                    data[i].append(int(values[j]))
                else:
                    data[i].append(values[j])
            bmi=int(values[2])/((int(values[1])/100)*(int(values[1])/100))
            data[i].append(bmi)
            if bmi<18.5:
                data[i].append("Underweight")
            elif 18.5<=bmi<25:
                data[i].append("Normal")
            elif 25<=bmi<40:
                data[i].append("Overweight")
            elif bmi>=40:
                data[i].append("Obesse")

    if count==n:
        return works(data,workbook,worksheet,file)

# sg.theme_previewer()
def inp_awal():
    layout = [
        [sg.Text('\n\n\nFile Name : ')],
        [sg.InputText()],
        [sg.Text('\nData Total : ',)],
        [sg.InputText()],
        [sg.Text("\n"*20+"\t"),sg.Submit(size=(8,2)), sg.Cancel('Exit',size=(8,2))]
    ]
    window = sg.Window('BMI to Excel',layout, size = (350,400))
    while True:
        event, values = window.read()
        window.close()
        if event==sg.WIN_CLOSED or event=="Cancel":
            break;
        elif event=="Submit":
            file=values[0]
            n = int(values[1])
            workbook = exc.Workbook(file+".xlsx")
            worksheet = workbook.add_worksheet()
            return  inp(n,workbook,worksheet,file)

inp_awal()