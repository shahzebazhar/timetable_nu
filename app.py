import tkinter as tk
from tkinter import ttk
import json

class App:

    def __init__(self, name : str):
        self.root = tk.Tk(name)

    

    def run(self):
        self.root.mainloop()

def changed(course_menu : tk.StringVar, section_menu : tk.StringVar, section_drop : tk.OptionMenu, sections : dict):
    
    section_menu.set("-")
    section_drop["menu"].delete(0, "end")

    try:
        for section in sections[course_menu.get().lower()].keys():
            section_drop["menu"].add_command(label=section.upper(), command=tk._setit(section_menu, section.upper()))
    except KeyError:
        pass


def check_input(event, combo_box, lst):
    value = event.widget.get()

    if value == '':
        combo_box['values'] = lst
    else:
        data = []
        for item in lst:
            if value.lower() in item.lower():
                data.append(item)

        combo_box['values'] = data

def main():
    root = tk.Tk(screenName="TimeTable Generator")
    root.geometry("800x600")
    
    frame = tk.Frame(root)
    frame.pack()

    course_menu = tk.StringVar(master=root)
    course_menu.set("Select Course")

    timetable = json.load(open("timetable.json"))

    courses = [s.title() for s in timetable["timetable"].keys()]

    course_drop = ttk.Combobox(master=frame, height=20, textvariable=course_menu, values=courses)
    course_drop.config(width=100)
    course_drop.grid(column=0, row=0)

    section_menu = tk.StringVar(master=root)
    section_menu.set("-")
    
    section_drop = tk.OptionMenu(frame, section_menu, "-")

    sections = timetable["timetable"]

    course_drop.bind("<KeyRelease>", lambda e : check_input(e, course_drop, courses))

    course_menu.trace("w", lambda *_ : changed(course_menu, section_menu, section_drop, sections))

    course_drop.grid(column=0, row=0)
    section_drop.grid(column=1, row=0)

    root.mainloop()


if __name__ == "__main__":
    main()