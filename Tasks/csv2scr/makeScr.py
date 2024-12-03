import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox



def check_name_input(func):
    def wrapper(*args, **kwargs):
        if (len(file_name.get()) == 0):
            messagebox.showwarning(title= "Wrong Input", message="Please choose a script file name.")
            return

        if (len(file_path_box.get()) == 0):
            messagebox.showwarning(title= "Wrong Input", message="Please choose a valid csv file.")
        return func(*args, **kwargs)
    return wrapper

@check_name_input
def main():
    filename = file_name.get()
    file = "".join([filename, ".scr"])

    with open(file, "w"):
        pass
    
    with open(file_path_var.get(), "r") as excel:
        number_of_points = 0
        controlPoints = list()

        for line in excel:
            number_of_points += 1
            try:
                items = line.split(",")
                pointID, x, y, z, layer = (items + [None] * 5)[:5]  # Fill missing values with None
                coordinate = ",".join([x,y,z])
            except Exception as e:
                messagebox.showwarning(f"Error processing line: {line} - {e}")
                on_close()

            if layer != None:
                layer = layer.strip()
            else:
                controlPoints.append((pointID, coordinate))
                layer = "ST"

            with open(file, "a") as f:
                f.write(f"-LAYER S TX_Point_Name_{layer.upper()} \n")
                f.write(f"POINT {coordinate}\n")
                f.write(f"TEXT {coordinate} 0.4 100 {pointID}\n")

    messagebox.showinfo(title="Csv Data Info", message=
                        f"Number of points: {number_of_points}"
                        f"Probable control points and it's coordinates: {controlPoints}"
                        )
    return

def browse_file():
    filename = filedialog.askopenfilename(
        title="Select a csv file",
        filetypes=(("CSV and TXT Files", "*.csv *.txt"), ("All Files", "*.*"))
    )
    if filename and os.path.splitext(filename)[1].lower() in ['.csv', '.txt']:
        print(f"Selected File: {filename}")
        file_path_var.set(filename)
        return
    else:
        messagebox.showwarning(title= "Wrong Input", message="Please choose a valid csv file.")


def on_close():
    root.destroy()

root = tk.Tk()
root.title("Create Script File")

title_label = tk.Label(root, text="Title of scr file:")
title_label.grid(row=0, column=0, padx=10, pady=5)

file_name = tk.Entry(root, width=50)
file_name.grid(row=0, column=1, padx=10, pady=5)

file_path_var = tk.StringVar()
file_path_label = tk.Label(root, text="Select .csv file:")
file_path_label.grid(row=1, column=0, padx=10, pady=5)

file_path_box = tk.Entry(root, textvariable=file_path_var, width=50)
file_path_box.grid(row=1, column=1, padx=10, pady=5)

browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=1, column=2, padx=5, pady=5)

run = tk.Button(root, text="Create Script", command=main)
run.grid(row=2, column=1, pady=10)

cancel_button = tk.Button(root, text="Cancel", command=on_close)
cancel_button.grid(row=2, column=2, pady=10)

root.protocol("WM_DELETE_WINDOW", on_close)

root.mainloop()


# if __name__ == "__main__":
#     main()
