import xlwt
from xlrd import open_workbook
from xlutils.copy import copy
from ttkthemes import themed_tk as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import IntVar
from tkinter import CENTER
import instaloader
import time

#--------TKINTER CONFIGS--------#
#----ROOT WINDOW SETTINGS----#
root = tk.ThemedTk()
root.get_themes()
root.set_theme("equilux")
root.title("Instagram Follow Extractor")
root.resizable(False, False)
root.geometry("325x190")
root.iconbitmap("icon.ico")

#----MAIN FRAME SETTINGS----#
main_frame = ttk.Frame()
main_frame.config(width="325", height="190")
main_frame.pack()

#----WIDGETS SETTINGS----#
#LABELS
username_label = ttk.Label(main_frame, text="Instagram Username")
target_label = ttk.Label(main_frame, text="Target Username")
password_label = ttk.Label(main_frame, text="Instagram Password")
filename_label = ttk.Label(main_frame, text="Excel Filename")
limit_label = ttk.Label(main_frame, text="Limit")

#ENTRIES
username_entry = ttk.Entry(main_frame, justify='center')
target_entry = ttk.Entry(main_frame, justify='center')
password_entry = ttk.Entry(main_frame, show="*", justify='center')
filename_entry = ttk.Entry(main_frame, justify='center')
limit_entry = ttk.Entry(main_frame, justify='center')

#CHECKBUTTONS
want_followees = False
want_followees = False
followers_check = IntVar()
followees_check = IntVar()

followers_checkbutton = ttk.Checkbutton(main_frame, text="Followers/Seguidores", variable=followers_check) 
followees_checkbutton = ttk.Checkbutton(main_frame, text="Followees/Siguiendo", variable=followees_check) 

#BUTTON
extract_button = ttk.Button(main_frame, text="EXTRACT NOW", command=lambda:LoggingIn(username_entry.get(), password_entry.get(), target_entry.get(), filename_entry.get(), limit_entry.get()))
#--------TKINTER CONFIGS--------#

#--------EXCEL CONFIGS--------#
#----BASIC SETTINGS----#
workbook = xlwt.Workbook()
sheet = workbook.add_sheet("Hoja 1", cell_overwrite_ok=True)

#----CUSTOM COLOR SETTINGS----#
#HEADER CELL COLOUR
xlwt.add_palette_colour("darker grey", 0x20)
workbook.set_colour_RGB(0x20, 38, 38, 38)
#ODD CELL COLOUR
xlwt.add_palette_colour("grey", 0x21)
workbook.set_colour_RGB(0x21, 67, 67, 67)
#EVEN CELL COLOUR
xlwt.add_palette_colour("lighter grey", 0x22)
workbook.set_colour_RGB(0x22, 121, 121, 121)

#----CELL SETTINGS----#
header_cell = xlwt.easyxf('pattern: pattern solid, fore_colour darker grey;''font: colour white, bold True;''align: horiz center;')
odd_cell = xlwt.easyxf('pattern: pattern solid, fore_colour grey;''font: colour white;')
even_cell = xlwt.easyxf('pattern: pattern solid, fore_colour lighter grey;''font: colour white;')
#--------EXCEL CONFIGS--------#

def WriteData(data, column, filename, limit):
    row = 1

    for user in data:
        if row < (int(limit)+1):
            if row%2 == 0:
                sheet.write(row, column, user.username, even_cell)
                workbook.save(filename + ".xls")
            else:
                sheet.write(row, column, user.username, odd_cell)
            row += 1
            main_frame.after(1)
            time.sleep(1.0)
        else:
            break

def ExtractData(target_profile, filename, limit):
    if(followees_check.get() == 1):
        want_followees = True
    else:
        want_followees = False

    if(followers_check.get() == 1):
        want_followers = True
    else:
        want_followers = False

    if want_followees:
        followees = target_profile.get_followees()
        sheet.col(0).width = 11063
        sheet.write(0, 0, "Followees/Siguiendo", header_cell)
        WriteData(followees, 0, filename, limit) 

        time.sleep(5)

    if want_followers:
        followers = target_profile.get_followers()
        sheet.col(1).width = 11063
        sheet.write(0, 1, "Followers/Seguidores", header_cell)
        WriteData(followers, 1, filename, limit)

def LoggingIn(username, password, target, filename, limit):
    start_time = time.time()
    #Default filename in case filename was not entered
    if filename == "":
        filename = "Instagram Follow Data"

    loginInsta = instaloader.Instaloader()
    try:
        loginInsta.login(username, password)
    except instaloader.BadCredentialsException:
        messagebox.showerror("Bad Credentials Error", "Contraseña incorrecta, intenta nuevamente")
        return
    except instaloader.InvalidArgumentException:
        messagebox.showerror("Invalid Argument Error", "El usuario solicitado no existe, intenta nuevamente")
        return

    try:
        target_profile = instaloader.Profile.from_username(loginInsta.context, str(target))
    except instaloader.exceptions.ProfileNotExistsException:
        messagebox.showerror("Profile Not Exists Error", "El usuario objetivo solicitado no existe, intenta nuevamente")
        return

    ExtractData(target_profile, filename, limit)
    workbook.save(filename + ".xls")
    messagebox.showinfo("Instagram Follow Extractor", "La descarga del listado de seguidores y seguidos ha finalizado\n" + 
                        "Tiempo de ejecucion: " + str(round(time.time() - start_time, 2)) + " segundos")

#----WIDGETS PLACING----#
#LABELS
username_label.place(x=33, y=10)
target_label.place(x=193, y=10)
password_label.place(x=33, y=55)
filename_label.place(x=198, y=55)
limit_label.place(x=220, y=135)

#ENTRIES
username_entry.place(x=20, y=30, width=135)
target_entry.place(x=170, y=30, width=135)
password_entry.place(x=20, y=75, width=135)
filename_entry.place(x=170, y=75, width=135)
limit_entry.place(x=170, y=152, width=135)

#CHECKBUTTONS
followers_checkbutton.place(x=0, y=130)
followees_checkbutton.place(x=0, y=155)

#BUTTON
extract_button.place(x=162, y=116, anchor=CENTER)

root.mainloop()