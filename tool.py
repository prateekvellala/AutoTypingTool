import tkinter as tk
import tkinter.scrolledtext as scrolledtext
from tkinter import messagebox
import webbrowser
import time
import multiprocessing
import sys
from tkinter import filedialog
from docx import Document

def typing(delay, interval, data):
	delay = int(delay)
	time.sleep(delay)
	import pyautogui
	pyautogui.FAILSAFE = False
	pyautogui.write(data, interval=interval)


def start_typing():
	global t1
	t1 =  multiprocessing.Process(target=typing, args=(ent_delay.get(), ent_interval.get(), txt_box.get("1.0", tk.END)[:-1]))
	t1.start()
	messagebox.showinfo("Message", "Please select a target window for typing.")


def stop_typing():
	t1.terminate() 
	t1.join()
	sys.stdout.flush()
	global i
        
	
def load_text():
    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt"), ("Word Files", "*.docx;*.doc")])
    if file_path:
        if file_path.endswith(".txt"):
            with open(file_path, "r") as file:
                text = file.read()
        elif file_path.endswith((".docx", ".doc")):
            doc = Document(file_path)
            paragraphs = [p.text for p in doc.paragraphs]
            text = "\n".join(paragraphs)
        else:
            messagebox.showerror("Error", "Invalid file format.")
            return

        txt_box.delete("1.0", tk.END)
        txt_box.insert(tk.END, text)



def clear_text():
    txt_box.delete("1.0", tk.END)


def exit_program():
	sys.exit()


def select_all(event):
    txt_box.tag_add(tk.SEL, "1.0", tk.END)
    txt_box.mark_set(tk.INSERT, "1.0")
    txt_box.see(tk.INSERT)
    return 'break'


def callback(url):
    webbrowser.open_new(url)


def configure_weight():
	
	frm_params.columnconfigure(0, weight=1)
	frm_params.columnconfigure(1, weight=1)
	frm_params.rowconfigure(0, weight=1)
	frm_params.rowconfigure(1, weight=1)
	
	frm_buttons.columnconfigure(0, weight=1)
	frm_buttons.columnconfigure(1, weight=1)
	frm_buttons.columnconfigure(2, weight=1)
	frm_buttons.rowconfigure(0, weight=1)
	
	window.columnconfigure(0, weight=1)
	window.rowconfigure(0, weight=1)
	window.rowconfigure(1, weight=1)
	window.rowconfigure(2, weight=1)
	window.rowconfigure(3, weight=1)
	window.rowconfigure(4, weight=1)	


def create_main_window():
    window.title("Auto Typing Tool")
    window.configure(bg="black")

    global frm_params
    frm_params = tk.Frame(window, bg="black")
    frm_params.grid(row=0, column=0)

    lbl_delay = tk.Label(text="Initial Delay (In Seconds)", master=frm_params, bg="black", fg="white")
    lbl_delay.grid(row=0, column=0, padx=50, pady=5)

    global ent_delay
    ent_delay = tk.Entry(justify='center', master=frm_params)
    ent_delay.insert(0, "3")
    ent_delay.config(bg="black", fg="white", insertbackground="white")
    ent_delay.grid(row=1, column=0, padx=50)

    lbl_interval = tk.Label(text="Interval (In Seconds)", master=frm_params, bg="black", fg="white")
    lbl_interval.grid(row=0, column=1, padx=50, pady=5)

    global ent_interval
    ent_interval = tk.Entry(justify='center', master=frm_params)
    ent_interval.insert(0, "0.01")
    ent_interval.config(bg="black", fg="white", insertbackground="white")
    ent_interval.grid(row=1, column=1, padx=50)

    ent_delay.config(bg="black", fg="white", insertbackground="white")
    ent_interval.config(bg="black", fg="white", insertbackground="white")


    lbl_data = tk.Label(text="Enter Your Text Here", font='Helvetica 18 bold', bg="black", fg="white")
    lbl_data.grid(row=3, column=0, pady=(10, 2))

    global txt_box
    txt_box = scrolledtext.ScrolledText(window, undo=True)
    txt_box.configure(font='Helvetica 12', fg="white", bg="black", insertbackground="white", width=80, height=30, wrap="word")
    txt_box.grid(row=4, column=0)

    txt_box.bind("<Control-Key-a>", select_all)
    txt_box.bind("<Control-Key-A>", select_all)

    global frm_buttons
    frm_buttons = tk.Frame(window, bg="black")
    frm_buttons.grid(row=5, column=0)

    start = tk.Button(text="Start", master=frm_buttons, command=start_typing, bg="black", fg="white")
    start.grid(row=0, column=0, padx=10, pady=10)

    stop = tk.Button(text="Stop", master=frm_buttons, command=stop_typing, bg="black", fg="white")
    stop.grid(row=0, column=1, padx=10, pady=10)

    exit_btn = tk.Button(text="Exit", master=frm_buttons, command=exit_program, bg="black", fg="white")
    exit_btn.grid(row=0, column=2, padx=10, pady=10)

    clear_btn = tk.Button(text="Clear Text", master=frm_buttons, command=clear_text, bg="black", fg="white")
    clear_btn.grid(row=0, column=3, padx=10, pady=10)

    upload_btn = tk.Button(text="Upload Text", master=frm_buttons, command=load_text, bg="black", fg="white")
    upload_btn.grid(row=0, column=4, padx=10, pady=10)

    configure_weight()
    window.mainloop()



if __name__ == '__main__':
	multiprocessing.freeze_support()
	window = tk.Tk()
	create_main_window()