
import os
import sys
import json
import os.path
import tkinter
import gspread
import datetime
from oauth2client.service_account import ServiceAccountCredentials

HELP_AREAS = ['Printer', 'Password Reset', 'App Support']
SCOPES = ['https://spreadsheets.google.com/feeds']


class FilteringCombobox:
    def __init__(self, root, values, height=100):
        self.root = root
        self.values = values
        self.height = height

        self.entry = tkinter.Entry(self.root)
        self.entry.bind('<KeyRelease>', self.check_key)
        self.entry.bind('<Button-1>', lambda e: self.expand())
        self.entry.bind('<Escape>', lambda e: self.collapse())
        self.entry_config = []

        self.listbox = tkinter.Listbox(self.root)
        self.listbox.bind('<Double-1>', lambda e: self.double_click())
        self.listbox.bind('<Escape>', lambda e: self.collapse())
        self.update(self.values)

    def expand(self):
        x, y = (self.entry.winfo_x(), self.entry.winfo_y() + self.entry.winfo_height())
        width = self.entry.winfo_width()
        self.listbox.place(x=x, y=y, width=width, height=self.height)
        self.listbox.lift()

    def collapse(self):
        self.listbox.place_forget()

    def check_key(self, event):
        value = event.widget.get()
        if value == '':
            data = self.values
        else:
            data = []
            for item in self.values:
                if value.lower() in item.lower():
                    data.append(item)
        self.update(data)

    def update(self, data):
        self.listbox.delete(0, 'end')
        for item in data:
            self.listbox.insert('end', item)

    def double_click(self):
        current_selection = self.listbox.curselection()
        self.entry.delete(0, 'end')
        self.entry.insert(tkinter.INSERT, self.listbox.get(current_selection))

    def pack(self, **kwargs):
        self.entry.pack(kwargs)
        self.entry_config = ['pack', kwargs]

    def grid(self, **kwargs):
        self.entry.grid(kwargs)
        self.entry_config = ['grid', kwargs]

    def place(self, **kwargs):
        self.entry.place(kwargs)
        self.entry_config = ['place', kwargs]

    def configure_entry(self, **kwargs):
        if self.entry_config[0] == 'pack':
            self.entry.pack_forget()
        elif self.entry_config[0] == 'grid':
            self.entry.grid_forget()
        elif self.entry_config[0] == 'place':
            self.entry.place_forget()

        self.entry = tkinter.Entry(self.root, kwargs)

        if self.entry_config[0] == 'pack':
            self.entry.pack(self.entry_config[1])
        elif self.entry_config[0] == 'grid':
            self.entry.grid(self.entry_config[1])
        elif self.entry_config[0] == 'place':
            self.entry.place(self.entry_config[1])

        self.entry.bind('<KeyRelease>', self.check_key)
        self.entry.bind('<Button-1>', lambda e: self.expand())

    def configure_listbox(self, **kwargs):
        self.listbox.place_forget()

        self.listbox = tkinter.Listbox(self.root, kwargs)
        self.listbox.bind('<Double-1>', lambda e: self.double_click())
        self.update(self.values)

    def get(self):
        return self.entry.get()

    def set(self, value):
        self.entry.delete(0, 'end')
        self.entry.insert(0, value)


class Spinbox(tkinter.Spinbox):
    def __init__(self, *args, **kwargs):
        tkinter.Spinbox.__init__(self, *args, **kwargs)
        self.bind('<MouseWheel>', self.mouse_wheel)
        self.bind('<Button-4>', self.mouse_wheel)
        self.bind('<Button-5>', self.mouse_wheel)

    def mouse_wheel(self, event):
        if event.num == 5 or event.delta == -120:
            self.invoke('buttondown')
        elif event.num == 4 or event.delta == 120:
            self.invoke('buttonup')


class CreateToolTip(object):
    def __init__(self, widget, text='widget info'):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", lambda e: self.enter())
        self.widget.bind("<Leave>", lambda e: self.close())
        self.tw = None

    def enter(self):
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        # creates a top level window
        self.tw = tkinter.Toplevel(self.widget)
        # Leaves only the label and removes the app window
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = tkinter.Label(self.tw, text=self.text, justify='left', background='yellow', relief='solid', borderwidth=1, font=("times", "8", "normal"))
        label.pack(ipadx=1)

    def close(self):
        if self.tw:
            self.tw.destroy()


class Gui:
    def __init__(self):
        with open(os.path.join(os.path.dirname(sys.executable), 'document_ids.json'), 'r') as json_file:
            json_data = json.load(json_file)
            self.user_id = json_data['usersID']
            self.log_id = json_data['logID']

        self.users_dict = {}
        self.client = None
        self.load_users()

        self.root = tkinter.Tk()
        self.root.geometry('520x400')
        self.root.resizable(width=False, height=False)

        tkinter.Label(self.root, text='Service Report', font='Helvetica 16 bold').place(x=10, y=10, width=150, height=30)
        self.frame = tkinter.Frame(self.root, borderwidth=2, relief='groove')
        self.frame.place(x=10, y=40, width=500, height=350)

        # Name
        tkinter.Label(self.frame, text='Name:', anchor=tkinter.W).place(x=10, y=10, width=60, height=25)
        self.name = FilteringCombobox(self.frame, self.users_dict.keys())
        self.name.place(x=90, y=10, width=140, height=25)

        # Department
        tkinter.Label(self.frame, text='Department:', anchor=tkinter.W).place(x=10, y=40, width=70, height=25)
        self.department = FilteringCombobox(self.frame, list(set(self.users_dict.values())))
        self.department.place(x=90, y=40, width=140, height=25)
        self.name.listbox.bind('<Double-1>', lambda e: self.auto_fill_department())

        # Category
        tkinter.Label(self.frame, text='Help Area:', anchor=tkinter.W).place(x=250, y=10, width=60, height=25)
        self.category = FilteringCombobox(self.frame, HELP_AREAS)
        self.category.place(x=320, y=10, width=150, height=25)

        # Fixed By
        tkinter.Label(self.frame, text='Fixed By:', anchor=tkinter.W).place(x=250, y=40, width=60, height=25)
        self.technician = FilteringCombobox(self.frame, self.users_dict.keys())
        self.technician.set(' '.join([s.capitalize() for s in os.getlogin().split('_')]))
        self.technician.place(x=320, y=40, width=150, height=25)

        # Solution
        tkinter.Label(self.frame, text='Solution:', anchor=tkinter.W).place(x=10, y=70, width=60, height=25)
        self.solution = tkinter.Text(self.frame)
        self.solution.place(x=90, y=70, width=380, height=200)

        # Duration
        tkinter.Label(self.frame, text='Job Length:', anchor=tkinter.W).place(x=10, y=280, width=60, height=25)
        self.duration_hours = Spinbox(self.frame, from_=0, to=23, increment=1)
        self.duration_hours.place(x=90, y=280, width=75, height=25)
        CreateToolTip(self.duration_hours, 'Hours')
        self.duration_minutes = Spinbox(self.frame, from_=0, to=59, increment=5)
        self.duration_minutes.place(x=165, y=280, width=75, height=25)
        CreateToolTip(self.duration_minutes, 'Minutes')

        # Submit
        self.submit = tkinter.Button(self.frame, text='Submit', borderwidth=2, relief='groove', command=lambda: self.save())
        self.submit.place(x=340, y=280, width=130, height=25)

        self.root.bind('<Configure>', lambda e: self.defocus_widgets())
        self.root.bind('<ButtonRelease-1>', lambda e: self.defocus_widgets())

        self.root.mainloop()

    def load_users(self):
        credentials = ServiceAccountCredentials.from_json_keyfile_name(os.path.join(os.path.dirname(sys.executable), 'credentials.json'), SCOPES)
        self.client = gspread.authorize(credentials)
        sheet = self.client.open_by_key(self.user_id)
        worksheet = sheet.sheet1

        users = [' '.join([name.capitalize() for name in user.split('_')]) for user in worksheet.col_values(1)[2:]]
        departments = worksheet.col_values(2)
        self.users_dict = {users[i]: departments[i] for i in range(len(users))}

    def defocus_widgets(self):
        x, y = self.root.winfo_pointerxy()
        active_widget = self.root.winfo_containing(x, y)
        if self.name.entry != active_widget and self.name.listbox != active_widget:
            self.name.collapse()
        if self.department.entry != active_widget and self.department.listbox != active_widget:
            self.department.collapse()
        if self.category.entry != active_widget and self.category.listbox != active_widget:
            self.category.collapse()
        if self.technician.entry != active_widget and self.technician.listbox != active_widget:
            self.technician.collapse()

    def auto_fill_department(self):
        user = self.name.listbox.get(self.name.listbox.curselection())
        department = self.users_dict[user]
        if department == 'None':
            department = ''
        self.department.entry.delete(0, 'end')
        self.department.entry.insert(tkinter.INSERT, department)
        self.name.double_click()

    def save(self):
        sheet = self.client.open_by_key(self.log_id)
        worksheet = sheet.sheet1

        row = str(len(list(filter(None, worksheet.col_values(1))))+1)
        worksheet.update_acell(f'A{row}', self.name.get())
        worksheet.update_acell(f'B{row}', self.department.get())
        worksheet.update_acell(f'C{row}', self.category.get())
        worksheet.update_acell(f'D{row}', self.technician.get())
        worksheet.update_acell(f'E{row}', self.solution.get('1.0', 'end-1c'))
        worksheet.update_acell(f'F{row}', f'{self.duration_hours.get()}h{self.duration_minutes.get()}m')
        worksheet.update_acell(f'G{row}', datetime.datetime.now().strftime('%m/%d/%Y %H:%M:%S'))


if __name__ == '__main__':
    Gui()
