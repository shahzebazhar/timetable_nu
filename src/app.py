import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tt_parser import TimeTableParser
import pandas as pd
import json
import webbrowser

class App:

    def __init__(self, name : str, window_size : str):
        self.root = tk.Tk(name)
        self.widgets = dict()
        self.pady = 15
        self.wRow = 0

        try:
            self.timetable = json.load(open("timetable.json"))
            self.courses = self.get_courses(self.timetable)
            self.sections = self.timetable["timetable"]
        except FileNotFoundError:
            self.timetable = {}
            self.courses = []
            self.sections = []
        
        self.total_courses = 7

        self.root.geometry(window_size)
        self.rootframe = tk.Frame(self.root)
        self.rootframe.pack()

    def widget_changed(self, keys, sections): 
        rowid = keys[0]
        section_menu = self.widgets[rowid]["Section Menu"]
        section_drop = self.widgets[rowid]["Section Dropbox"]
        course_menu = self.widgets[rowid]["Course Menu"]
        
        section_menu.set("-")
        section_drop["menu"].delete(0, "end")

        try:
            for section in sections[course_menu.get().lower()].keys():
                section_drop["menu"].add_command(label=section.upper(), command=tk._setit(section_menu, section.upper()))
        except KeyError:
            pass

    def check_input(self, event, lst):
        
        combo_box = event.widget
        value = event.widget.get()

        if value == '':
            combo_box['values'] = lst
        else:
            data = []
            for item in lst:
                if value.lower() in item.lower():
                    data.append(item)

            combo_box['values'] = data

    def courses_submitted(self):
        data = {
            "Courses" : [],
            "Sections" : [],
            "Day" : [],
            "Venue" : [],
            "Start Time" : [],
            "End Time" : []
        }
        for key in self.widgets:
            row = self.widgets[key]
            course = row['Course Menu'].get()
            section = row['Section Menu'].get()
        
            try:
                for el in self.sections[course.lower()][section.lower()]:
                    data["Courses"].append(course.lower())
                    data["Sections"].append(section.lower())
                    data["Day"].append(el[0].lower())
                    data["Venue"].append(el[1].lower())
                    data["Start Time"].append(el[2].lower())
                    data["End Time"].append(el[3].lower())
                for el in self.sections[f"{course.lower()} lab"][section.lower()]:
                    data["Courses"].append(f"{course.lower()} lab")
                    data["Sections"].append(section.lower())
                    data["Day"].append(el[0].lower())
                    data["Venue"].append(el[1].lower())
                    data["Start Time"].append(el[2].lower())
                    data["End Time"].append(el[3].lower())
            except KeyError as e:
                print(f"No course named {course} or no lab for {course}")

        ttdf = self.generate_timetable(data)

        ttdf.to_html("Timetable.html")

        print("Timetable generated to Timetable.html")

        webbrowser.open_new_tab("Timetable.html")

    def generate_timetable(self, data : dict) -> pd.DataFrame:
        timings = self.timetable["timings"]
        days = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday"]
        
        # Exclude the timings from the end if they don't exist.
        i = len(timings) - 1
        while i >= 0:
            if timings[i] in data["End Time"] or timings[i] in data["Start Time"]:
                break
            else:
                i -= 1
            
        timings = timings[:i + 1]
        print(timings)

        # Excluding all the days after which there are no classes
        i = len(days) - 1
        while i >= 0:
            if days[i] in data["Day"]:
                break
            else:
                i -= 1

        days = days[:i + 1]
        print(days)

        # Generating a structure for timetable.

        tt_struct = dict()
        for timing in timings:
            tt_struct[timing] = [""] * len(days)
    
        for i in range(len(data["Courses"])):
            day_idx = days.index(data["Day"][i])
            content = f"{data['Courses'][i].title()} ({data['Sections'][i].upper()}) [VENUE: {data['Venue'][i].upper()}]"

            stidx = timings.index(data["Start Time"][i])
            enidx = timings.index(data["End Time"][i])

            while stidx < enidx:
                if tt_struct[timings[stidx]][day_idx] != "" and tt_struct[timings[stidx]][day_idx] != content:
                    print(f"Clash found for the course {tt_struct[timings[stidx]][day_idx]} and {content}, Adding it anyways by appending")
                    tt_struct[timings[stidx]][day_idx] += " | " + content
                else:
                    tt_struct[timings[stidx]][day_idx] = content

                stidx += 1
        
        return pd.DataFrame(tt_struct, index=days)



    def parse_timetable(self):
        file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

        parser = TimeTableParser(file)
        
        self.timetable = parser.parse()
        json.dump(self.timetable, open("timetable.json", "w"), indent=4)
        self.courses = self.get_courses(self.timetable)
        self.sections = self.timetable["timetable"]

        messagebox.showinfo("Timetable Updated", f"The timetable has been updated in accordance with {file}.")

    @staticmethod
    def get_courses(timetable):
        return [s.title() for s in timetable["timetable"].keys()]
    

    def run(self):
        tt_button = tk.Button(self.rootframe, text="Parse Timetable", command = self.parse_timetable)
        tt_button.grid(column=0, row=self.wRow, pady=self.pady)
        self.wRow += 1

        for i in range(self.total_courses):
            course_menu = tk.StringVar(master=self.root)
            row_id = str(course_menu)
            self.widgets[row_id] = dict()
            self.widgets[row_id]["Course Menu"] = course_menu
            self.widgets[row_id]["Course Menu"].set("Select Course")

            self.widgets[row_id]["Course Dropbox"] = ttk.Combobox(master=self.rootframe, height=20, 
                                                        textvariable=self.widgets[row_id]["Course Menu"], values=self.courses)
            self.widgets[row_id]["Course Dropbox"].config(width=100)
            self.widgets[row_id]["Course Dropbox"].grid(column=0, row=self.wRow)

            self.widgets[row_id]["Section Menu"] = tk.StringVar(master=self.root)
            self.widgets[row_id]["Section Menu"].set("-")

            self.widgets[row_id]["Section Dropbox"] = tk.OptionMenu(self.rootframe, self.widgets[row_id]["Section Menu"], "-")

            self.widgets[row_id]["Course Dropbox"].bind("<KeyRelease>", lambda e : self.check_input(e, self.courses))

            course_menu.trace("w", lambda *_ : self.widget_changed(_, self.sections))

            self.widgets[row_id]["Course Dropbox"].grid(column=0, row=self.wRow, pady=self.pady)
            self.widgets[row_id]["Section Dropbox"].grid(column=1, row=self.wRow, pady=self.pady)
            self.wRow += 1
            
        submit_btn = tk.Button(self.rootframe, text="Submit", command = lambda : self.courses_submitted())
        submit_btn.grid(column=0, row=self.wRow, pady=self.pady)
        self.wRow += 1

        self.root.mainloop()


def main():
    app = App("TimeTable Generator", "800x600")

    app.run()


if __name__ == "__main__":
    main()