from tkinter import *
from tkinter import ttk
import functions
from main import Tabs
from constants import *

# =====Create Window=====
window = Tk()
window.title("RadioWaste Manager")
window.config(bg=BG)
window.geometry("1050x650")
functions.center_window(window, 1050, 650)

# =====Create Main Frame=====
main_frame = Frame(window, bg=BG, highlightthickness=0)
main_frame.pack(fill="both", expand=True)
main_frame.columnconfigure(0, weight=0)
main_frame.columnconfigure(1, weight=1)
main_frame.rowconfigure(0, weight=1)

#=====Create Button Frame inside Main Frame=====
button_frame = Frame(main_frame, bg=BG, highlightthickness=0)
button_frame.grid(column=0, row=0, sticky="nsw", padx=(15,5), pady=20)
button_frame.grid_rowconfigure(4, weight=1)

# =====Create Customized Tabs Frame inside Main Frame=====
style = ttk.Style()
style.theme_use("default")
style.configure("TNotebook.Tab", background=C3, foreground="white", font=(FONT_NAME,8, "normal"), padding=[8, 8])
style.map("TNotebook.Tab", background=[("selected", C4)], foreground=[("selected", "white")])
style.configure("TNotebook", background=BG, borderwidth=0)
style.configure("TFrame", background=BG, borderwidth=0)

tabs_frame = ttk.Notebook(main_frame)
tabs_frame.grid(column=1, row=0, sticky="nsew", padx=10, pady=10)

# =====Create Main Tab with image=====
main_tab = Frame(tabs_frame, bg=BG)
tabs_frame.add(main_tab, text="Main Menu")

canvas = Canvas(main_tab, width=780, height=600, bg=BG, highlightthickness=0)
test_img = PhotoImage(file="white_bg_radioactive.png")
canvas.create_image(350,450, image=test_img)
main_text = canvas.create_text(400,30,text="MAIN MENU", font=(FONT_NAME,32,"bold"), fill="white")
canvas.pack()

func = Tabs(window, tabs_frame, main_tab)

# =====Create Main Buttons=====
insert_new_button = Button(button_frame, text="VIALS", **BUTTON_STYLE, command=lambda:func.create_new_tab("Vials"))
insert_new_button.grid(column=0, row=0, padx=5, pady=5, sticky="w")
gen_button = Button(button_frame, text="GENERATORS", **BUTTON_STYLE, command=lambda:func.create_new_tab("Generators"))
gen_button.grid(column=0, row=1, padx=5, pady=5, sticky="w")
i131_button = Button(button_frame, text="I-131", **BUTTON_STYLE, command=lambda:func.create_new_tab("I131"))
i131_button.grid(column=0, row=2, padx=5, pady=5, sticky="w")
disposal_button = Button(button_frame, text="DISPOSAL", **BUTTON_STYLE, command=lambda:func.create_new_tab("Disposal"))
disposal_button.grid(column=0, row=3, padx=5, pady=5, sticky="w")
exit_button = Button(button_frame, text="Exit", **{k:v for k,v in BUTTON_STYLE.items() if k not in ['width','height']}, width=12, height=1, command=quit)
exit_button.grid(column=0, row=5, padx=0, pady=0, sticky="s")


window.mainloop()
