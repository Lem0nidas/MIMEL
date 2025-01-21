import os
import sys
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES



def check_name_input(func):
    def wrapper(*args, **kwargs):
        if (len(file_name.get()) == 0):
            messagebox.showwarning(title= "Wrong Input", message="Please choose a script file name.")
            return

        if (len(file_path_box.get('1.0', 'end').strip()) == 0):
            messagebox.showwarning(title= "Wrong Input", message="Please choose a valid csv file.")
        return func(*args, **kwargs)
    return wrapper

@check_name_input
def main():
    coords_path = load_control_points()
    clean_csv_path = file_path_box.get('1.0', 'end').strip().replace('{', '').replace('}', '')

    directory = os.path.dirname(clean_csv_path)
    filename = file_name.get()
    file = os.path.join(directory, f"{filename}.scr")

    with open(file, "w") as f:
        f.write("OSMODE 0 3DOSMODE 0\n")
        pass

    with open(clean_csv_path, "r") as excel:
        number_of_points = 0
        controlPoints = list()

        for line in excel:
            number_of_points += 1
            try:
                items = line.split(",")
                pointID, x, y, z, layer = (items + [None] * 5)[:5]
                coordinate = ",".join([x,y,z])

                pointID = pointID.strip('" ')
                layer = layer.strip('" \n').replace(" ", "-")
            except Exception as e:
                messagebox.showwarning(f"Error processing line: {line} - {e}")
                on_close()

            if len(layer) > 0:
                layer_command = check_layer(layer)
                coord_check = "No control point found in csv"
            else:
                idCoords = control_point_coords(coords_path, pointID.upper())
                if idCoords:
                    
                    controlPoints.append((pointID, coordinate))
                    rounded_coords = [str(round(float(i), 3)) for i in idCoords]
                    
                    if (",".join(rounded_coords) == coordinate):
                        coord_check = "The control point is valid (It's coordinates is included in FINAL_COORDS_TR)"
                    else:
                        coord_check = "The control point is NOT valid (Coordinates DO NOT match in FINAL_COORDS_TR)"

                    layer_command = "-LAYER M TX_Point_Name_ST \n"
                else:
                    coord_check = "The control point was not found in FINAL_COORDS_TR"
                    layer_command = "-LAYER M TX_Point_Name_STK \n"

            with open(file, "a") as f:
                f.write(layer_command)
                f.write(f"POINT {coordinate}\n")
                f.write(f"TEXT {coordinate} 0.4 100 {pointID}\n")

    report = f"Number of points: {number_of_points}\nControl points and it's coordinates: \n{controlPoints}\n" + coord_check

    messagebox.showinfo(title="Csv Data Info", message=report)
    return

def load_control_points():
    if getattr(sys, "frozen", False):
        # If the script is frozen (converted to an .exe)
        bundle_dir = sys._MEIPASS
    else:
        bundle_dir = os.path.abspath(os.path.dirname(__file__))

    coords_path = os.path.join(bundle_dir, 'Data', 'ALL_COORDS.txt')

    return coords_path

def control_point_coords(file_path, search_id):
    with open(file_path, 'r') as file:
        for line in file:
            values = line.split(',')
            if values[0] == search_id:
                return (values[1], values[2], values[3])
    return None


def check_layer(description):
    if (description.upper() in ["PD", "FR", "TX"]):
        return f"-LAYER S TX_Point_Name_{description.upper()} \n"
    return f"-LAYER M TX_Point_Name_{description.upper()} \n"


def browse_file():
    filename = filedialog.askopenfilename(
        title="Select a csv file",
        filetypes=(("CSV and TXT Files", "*.csv *.txt"), ("All Files", "*.*"))
    )
    if filename and os.path.splitext(filename)[1].lower() in ['.csv', '.txt']:
        file_path_box.delete('1.0', 'end')
        file_path_box.insert('1.0', filename)
        return
    else:
        messagebox.showwarning(title= "Wrong Input", message="Please choose a valid csv file.")

def on_drop(event):
    file_path_box.delete('1.0', 'end')
    file_path_box.insert('1.0', event.data)
    return

def on_close():
    root.destroy()

root = TkinterDnD.Tk()
root.title("Create Script File")

title_label = tk.Label(root, text="Title of scr file:")
title_label.grid(row=0, column=0, padx=10, pady=5)

file_name = tk.Entry(root, width=50)
file_name.grid(row=0, column=1, padx=10, pady=5)

file_path_label = tk.Label(root, text="Select .csv file:")
file_path_label.grid(row=1, column=0, padx=10, pady=5)

file_path_box = tk.Text(root, width=50, height=5, bg="lightgray", wrap='word')
file_path_box.insert('1.0', "Drag and drop files here...")
file_path_box.grid(row=1, column=1, padx=10, pady=5)

file_path_box.drop_target_register(DND_FILES)
file_path_box.dnd_bind('<<Drop>>', on_drop)

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
