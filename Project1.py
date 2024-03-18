import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog as fd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import datetime as dt
from PIL import Image,ImageGrab
import tkinter.font as tf

global fields

def read_excel():
    global data
    global df 
    global fields
    global g_data
    global total_sales
    File=txt.get("0.0","end-1c")
    try:
        data=pd.read_excel(File)                         
        messagebox.showinfo("Load","Data Load Successfully")
        fields=list(data.columns)        
        print(fields)
        
       
    except Exception as e:
        messagebox.showinfo("fail","!! failed to load")


    display_values=list(fields)
    combobox1.config(values=display_values)
    combobox2.config(values=display_values)
        
window=tk.Tk()
window.title("Data Visualization Dashboard")
window.state('zoomed')
window.config(bg="#C0C0C0")

tf_font=tf.Font(family="Helvetica",size="15")

side_frame=tk.Frame(window)
side_frame.pack(side="left",fill="y")

charts_frame=tk.Frame(window)
charts_frame.pack() 

upper_frame=tk.Frame(charts_frame)
upper_frame.pack(fill="both",expand=True)

lower_frame=tk.Frame(charts_frame)
lower_frame.pack(fill="both",expand=True)


canvas=tk.Canvas(side_frame,width=200,height=1000)
canvas.pack()

start_color="#C33764"
end_color="#1D2671"

def gradient(canvas,x1,y1,x2,y2,color1,color2):
    canvas.create_rectangle(x1,y1,x2,y2,fill=color1,outline=color1)
    for i in range(1,100):
        r=int((i/100)*int(color2[1:3],16)+(1-i/100)*int(color1[1:3],16))
        g=int((i/100)*int(color2[3:5],16)+(1-i/100)*int(color1[3:5],16))
        b=int((i/100)*int(color2[5:7],16)+(1-i/100)*int(color1[5:7],16))
        color=f"#{r:02X}{g:02X}{b:02X}"
        canvas.create_rectangle(x1,y1 +i*(y2-y1)/100,x2,y1 +(i+1)*(y2-y1)/100,fill=color,outline=color)

gradient(canvas,0,0,200,1000,start_color,end_color)

def openfile():
    filepath=fd.askopenfilename(initialdir="D:\\",filetypes=[("all files","*.*"),("Excel Files","*.xlsx")])
    txt.insert(0.0,filepath)    

lblx=tk.Label(side_frame,text="X-axis",fg="black",font=8,bg="lightblue" )
lblx.place(x=10,y=200)

lblx=tk.Label(side_frame,text="Y-axis",fg="black",font=8,bg="lightblue")
lblx.place(x=10,y=235)

combobox1=ttk.Combobox(side_frame,width=15)
combobox1.place(x=80,y=207)

combobox2=ttk.Combobox(side_frame,width=15)
combobox2.place(x=80,y=238)
def option_selected(event):
    global fields
    global selected_option1
    selected_option1= combobox1.get()
    print("You selected:", selected_option1)
combobox1.bind("<<ComboboxSelected>>", option_selected)
def option_selected(event):
    global fields
    global selected_option2
    selected_option2= combobox2.get()
    print("You selected:", selected_option2)
    print(type(selected_option2))
combobox2.bind("<<ComboboxSelected>>", option_selected)

lbl=tk.Label(side_frame,text="Dashboard",font=40,fg="Black")
lbl.place(x=42,y=15)
btn_insert=tk.Button(side_frame,text="Browse...",fg="black",font=2,width=10,command=openfile)
btn_insert.place(x=15,y=100)

txt=tk.Text(side_frame,height=0.5,width=18)
txt.place(x=24,y=76)

btn_load=tk.Button(side_frame,text="Load",fg="black",font=5,width=10,command=read_excel)
btn_load.place(x=15,y=148)

fig,axes1=plt.subplots(2,3,figsize=(10,6))
fig.subplots_adjust(wspace=1.5, hspace=1.5)
fig.tight_layout(rect=[0, 0.09, 1, 0.99], w_pad=0.8, h_pad=1.2)
canvas1=FigureCanvasTkAgg(fig,master=window)
canvas1.get_tk_widget().pack(side="right",fill="both",expand=True)

def bar():
    global data
    global selected_option1
    global selected_option2
    axes1[0][0].clear()
    x_axis=data[selected_option1]
    y_axis=data[selected_option2]
    print(x_axis)
    print(y_axis)
    print(type(x_axis))
    print(type(y_axis))
    g_data=data.groupby(selected_option1)[selected_option2].sum()
    axes1[0][0].bar(g_data.index,g_data.values)
    for bar in axes1[0][0].containers[0]:
        value_in_lakhs = round(bar.get_height() / 1000, 2)
        axes1[0][0].text(bar.get_x() + bar.get_width() / 2, bar.get_height(), f'{value_in_lakhs:.2f}k', ha='center')
        
    axes1[0][0].set_xlabel(selected_option1)
    axes1[0][0].set_ylabel(selected_option2)
    tit=selected_option1 +" by "+ selected_option2
    axes1[0][0].set_title(tit)
       
    canvas1.draw()
    
    
btn_bar=tk.Button(side_frame,text="Bar Chart",fg="black",font=tf_font,width=10,command=bar)
btn_bar.place(x=15,y=300)

def area():
    global data
    global selected_option1
    global selected_option2
    axes1[0][1].clear()
    x_axis=data[selected_option1]
    y_axis=data[selected_option2]
    g_data=data.groupby(selected_option1)[selected_option2].sum() 
    axes1[0][1].fill_between(g_data.index,g_data.values)
    axes1[0][1].set_xlabel(selected_option1)
    axes1[0][1].set_ylabel(selected_option2)
    tit = selected_option1 + " by " + selected_option2
    axes1[0][1].set_title(tit)
    canvas1.draw()
    
btn_area=tk.Button(side_frame,text="Area Chart",fg="Black",font=tf_font,width=10,command=area)
btn_area.place(x=15,y=350)

def pie():
    global sumdata   
    axes1[0][2].clear() 
    values=data[selected_option2]
    print(values)
    g_data=data.groupby(selected_option1)[selected_option2].sum()
    x=data.groupby(selected_option1).count()
    totals = data.groupby(selected_option1)[selected_option2].sum().tolist()
    total=[]
    for i in totals:
        total=total+[i/1000]
    Labels=list(x.index)
    print(total)
    new_labels = [f'{total:.2f}k' for total in total]

    axes1[0][2].pie(g_data,labels=new_labels,autopct='%1.1f%%',startangle=0)
    axes1[0][2].axis('equal')

    axes1[0][2].legend(title="legend",labels=Labels,loc="upper left")
    tit=selected_option1 + " by " + selected_option2
    axes1[0][2].set_title(tit)
    canvas1.draw()
    
btn_pie=tk.Button(side_frame,text="Pie chart",fg="Black",font=tf_font,width=10,command=pie)
btn_pie.place(x=15,y=400) 

def line():
    global data
    global selected_option1
    global selected_option2
    global fig1
    axes1[1][1].clear()
    x_axis=data[selected_option1]
    y_axis=data[selected_option2]
    print(x_axis)
    print(y_axis)
    print(type(x_axis))
    print(type(y_axis))
    g_data=data.groupby(selected_option1)[selected_option2].sum()

    axes1[1][1].plot(g_data,marker='o')
    axes1[1][1].set_xlabel(selected_option1)
    axes1[1][1].set_ylabel(selected_option2)
    tit=selected_option1  +" by "+  selected_option2
    axes1[1][1].set_title(tit)
    
    canvas1.draw()
btn_line=tk.Button(side_frame,text="Line Chart",fg="black",font=tf_font,width=10,command=line)
btn_line.place(x=15,y=450)

def scatter():
    global selected_option2
    global selected_option1
    g_data=data.groupby(selected_option1)[selected_option2].sum()
    axes1[1][2].clear()
    axes1[1][2].scatter(g_data.index,g_data.values)
    axes1[1][2].set_xlabel(selected_option1)
    axes1[1][2].set_ylabel(selected_option2)
    tit=selected_option1  +" by "+  selected_option2
    axes1[1][2].set_title(tit)
    canvas1.draw()
btn_scatter=tk.Button(side_frame,text="Scatter plot",fg="black",font=tf_font,width=10,command=scatter)
btn_scatter.place(x=15,y=500)

def barh():
    global data
    global selected_option1
    global selected_option2
    axes1[1][0].clear()
    x_axis=data[selected_option1]
    y_axis=data[selected_option2]
    print(x_axis)
    print(y_axis)
    print(type(x_axis))
    print(type(y_axis))
    g_data=data.groupby(selected_option1)[selected_option2].sum()
    axes1[1][0].barh(g_data.index,g_data.values)    
    axes1[1][0].set_xlabel(selected_option2)
    axes1[1][0].set_ylabel(selected_option1)
    tit=selected_option1 +" by "+ selected_option2
    axes1[1][0].set_title(tit)
    
       
    canvas1.draw()
    
btn_bar=tk.Button(side_frame,text="Bar-H-chart",fg="black",font=tf_font,width=10,command=barh)
btn_bar.place(x=15,y=550)

def clear():
    global axes1
    global canvas1
    axes1[0][0].clear()
    axes1[0][1].clear()
    axes1[1][0].clear()
    axes1[1][1].clear()
    axes1[0][2].clear()
    axes1[1][2].clear()
    canvas1.draw()
    
btn_clea=tk.Button(side_frame,text="Clear",fg="black",width=10,font=tf_font,command=clear)
btn_clea.place(x=15,y=850)

def save():
    global figure
    file=fd.asksaveasfilename(defaultextension=".png",filetypes=[("PNG files","*.png"),("PDF files","*.pdf")])
    if file:
        plt.savefig(file+".png")
        plt.savefig(file+".pdf")

    messagebox.showinfo("Download","Download Successful")
btn_save=tk.Button(side_frame,text="Download",fg="black",font=5,width=10,command=save)
btn_save.place(x=15,y=900)

window.mainloop()





