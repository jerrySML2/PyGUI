import pandas as pd
from tkinter import Tk, Label, Button, Checkbutton, Radiobutton, IntVar, BooleanVar, Toplevel
from tkinter.messagebox import showinfo
import win32com.client

# Load the Excel file
file_path = 'C:\\CalendarPA\\CPA_Schedule.xlsx'
df = pd.read_excel(file_path)

# Sort and create unique roles
df = df.sort_values(by="Role")
unique_roles = df["Role"].unique()

# Create the main dialog
def open_dialog():
    def submit():
        selected_roles = [role for role, var in role_vars.items() if var.get()]
        action = action_var.get()

        if not selected_roles:
            showinfo("Error", "Please select at least one role.")
            return

        if action not in (1, 2):
            showinfo("Error", "Please select an action (Create or Delete).")
            return

        dialog.destroy()
        process_events(selected_roles, action)

    dialog = Toplevel(root)
    dialog.title("Select Roles and Action")

    Label(dialog, text="Select Roles:").pack(anchor="w")

    role_vars = {role: BooleanVar() for role in unique_roles}
    for role, var in role_vars.items():
        Checkbutton(dialog, text=role, variable=var).pack(anchor="w")

    Label(dialog, text="Select Action:").pack(anchor="w")

    action_var = IntVar()
    Radiobutton(dialog, text="Create", variable=action_var, value=1).pack(anchor="w")
    Radiobutton(dialog, text="Delete", variable=action_var, value=2).pack(anchor="w")

    Button(dialog, text="Submit", command=submit).pack()

# Process events based on user selection
def process_events(selected_roles, action):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    calendar = namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar

    for role in selected_roles:
        role_events = df[df["Role"] == role]
        for _, event in role_events.iterrows():
            subject = event["Subject"]
            start = event["Start"]
            end = event["End"]

            if action == 1:  # Create events
                outlook_event = calendar.Items.Add()
                outlook_event.Subject = subject
                outlook_event.Start = start
                outlook_event.End = end
                outlook_event.Save()
                print(f"Created event: {subject}")

            elif action == 2:  # Delete events
                items = calendar.Items
                items.IncludeRecurrences = True
                filtered_items = items.Restrict(f"[Subject] = '{subject}' AND [Start] >= '{start}' AND [End] <= '{end}'")
                for item in filtered_items:
                    item.Delete()
                    print(f"Deleted event: {subject}")

# Set up the main Tkinter window
root = Tk()
root.withdraw()  # Hide the main window
open_dialog()
root.mainloop()
