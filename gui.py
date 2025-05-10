import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import json
import os
from main import EmailAutomation
import threading
import logging
from datetime import datetime

class EmailAutomationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Automation System")
        self.root.geometry("800x600")
        
        self.automation = EmailAutomation()
        self.is_running = False
        self.automation_thread = None
        
        self.create_gui()
        self.load_config()
        
    def create_gui(self):
        # Create notebook for tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(expand=True, fill='both', padx=10, pady=5)
        
        # Email Settings Tab
        email_frame = ttk.Frame(notebook)
        notebook.add(email_frame, text='Email Settings')
        self.create_email_settings(email_frame)
        
        # Recipients Tab
        recipients_frame = ttk.Frame(notebook)
        notebook.add(recipients_frame, text='Recipients')
        self.create_recipients_section(recipients_frame)
        
        # Schedule Tab
        schedule_frame = ttk.Frame(notebook)
        notebook.add(schedule_frame, text='Schedule')
        self.create_schedule_section(schedule_frame)
        
        # Logs Tab
        logs_frame = ttk.Frame(notebook)
        notebook.add(logs_frame, text='Logs')
        self.create_logs_section(logs_frame)
        
        # Control Buttons
        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill='x', padx=10, pady=5)
        
        self.start_button = ttk.Button(control_frame, text="Start Automation", command=self.start_automation)
        self.start_button.pack(side='left', padx=5)
        
        self.stop_button = ttk.Button(control_frame, text="Stop Automation", command=self.stop_automation, state='disabled')
        self.stop_button.pack(side='left', padx=5)
        
        self.save_button = ttk.Button(control_frame, text="Save Settings", command=self.save_config)
        self.save_button.pack(side='right', padx=5)

    def create_email_settings(self, parent):
        # Email Settings
        settings_frame = ttk.LabelFrame(parent, text="Email Configuration")
        settings_frame.pack(fill='x', padx=10, pady=5)
        
        # SMTP Server
        ttk.Label(settings_frame, text="SMTP Server:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.smtp_server = ttk.Entry(settings_frame, width=40)
        self.smtp_server.grid(row=0, column=1, padx=5, pady=5)
        
        # SMTP Port
        ttk.Label(settings_frame, text="SMTP Port:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.smtp_port = ttk.Entry(settings_frame, width=40)
        self.smtp_port.grid(row=1, column=1, padx=5, pady=5)
        
        # Sender Email
        ttk.Label(settings_frame, text="Sender Email:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        self.sender_email = ttk.Entry(settings_frame, width=40)
        self.sender_email.grid(row=2, column=1, padx=5, pady=5)
        
        # App Password
        ttk.Label(settings_frame, text="App Password:").grid(row=3, column=0, padx=5, pady=5, sticky='w')
        self.app_password = ttk.Entry(settings_frame, width=40, show="*")
        self.app_password.grid(row=3, column=1, padx=5, pady=5)
        
        # Email Template
        template_frame = ttk.LabelFrame(parent, text="Email Template")
        template_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(template_frame, text="Subject:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.email_subject = ttk.Entry(template_frame, width=40)
        self.email_subject.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(template_frame, text="Body:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.email_body = scrolledtext.ScrolledText(template_frame, width=40, height=5)
        self.email_body.grid(row=1, column=1, padx=5, pady=5)

    def create_recipients_section(self, parent):
        # Recipients List
        recipients_frame = ttk.LabelFrame(parent, text="Email Recipients")
        recipients_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Recipients Listbox
        self.recipients_listbox = tk.Listbox(recipients_frame, width=50, height=10)
        self.recipients_listbox.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(recipients_frame, orient="vertical", command=self.recipients_listbox.yview)
        scrollbar.pack(side='right', fill='y')
        self.recipients_listbox.configure(yscrollcommand=scrollbar.set)
        
        # Add/Remove Recipients
        control_frame = ttk.Frame(parent)
        control_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(control_frame, text="New Recipient:").pack(side='left', padx=5)
        self.new_recipient = ttk.Entry(control_frame, width=40)
        self.new_recipient.pack(side='left', padx=5)
        
        ttk.Button(control_frame, text="Add", command=self.add_recipient).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Remove Selected", command=self.remove_recipient).pack(side='left', padx=5)

    def create_schedule_section(self, parent):
        # Schedule Settings
        schedule_frame = ttk.LabelFrame(parent, text="Schedule Configuration")
        schedule_frame.pack(fill='x', padx=10, pady=5)
        
        # Time Entry
        ttk.Label(schedule_frame, text="Time (HH:MM):").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.schedule_time = ttk.Entry(schedule_frame, width=10)
        self.schedule_time.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(schedule_frame, text="Add Time", command=self.add_schedule_time).grid(row=0, column=2, padx=5, pady=5)
        
        # Schedule Listbox
        self.schedule_listbox = tk.Listbox(schedule_frame, width=20, height=5)
        self.schedule_listbox.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky='w')
        ttk.Button(schedule_frame, text="Remove Selected", command=self.remove_schedule_time).grid(row=1, column=2, padx=5, pady=5)
        
        # Days Selection
        days_frame = ttk.LabelFrame(parent, text="Days of Week")
        days_frame.pack(fill='x', padx=10, pady=5)
        
        self.days_vars = {}
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        for i, day in enumerate(days):
            var = tk.BooleanVar()
            self.days_vars[day.lower()] = var
            ttk.Checkbutton(days_frame, text=day, variable=var).grid(row=i//4, column=i%4, padx=5, pady=5, sticky='w')

    def create_logs_section(self, parent):
        # Logs Display
        self.logs_text = scrolledtext.ScrolledText(parent, width=80, height=20)
        self.logs_text.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Custom handler for logging to GUI
        class GUILogHandler(logging.Handler):
            def __init__(self, text_widget):
                super().__init__()
                self.text_widget = text_widget
                
            def emit(self, record):
                msg = self.format(record)
                self.text_widget.insert('end', msg + '\n')
                self.text_widget.see('end')
        
        # Add GUI handler to logger
        gui_handler = GUILogHandler(self.logs_text)
        gui_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logging.getLogger().addHandler(gui_handler)

    def load_config(self):
        config = self.automation.config
        
        # Load email settings
        self.smtp_server.insert(0, config["email"]["smtp_server"])
        self.smtp_port.insert(0, str(config["email"]["smtp_port"]))
        self.sender_email.insert(0, config["email"]["sender_email"])
        self.app_password.insert(0, config["email"]["app_password"])
        
        # Load email template
        self.email_subject.insert(0, config["email_template"]["subject"])
        self.email_body.insert('1.0', config["email_template"]["body"])
        
        # Load recipients
        for recipient in config["recipients"]:
            self.recipients_listbox.insert('end', recipient)
        
        # Load schedule times
        for time in config["schedule"]["times"]:
            self.schedule_listbox.insert('end', time)
        
        # Load days
        for day in config["schedule"]["days"]:
            if day in self.days_vars:
                self.days_vars[day].set(True)

    def save_config(self):
        try:
            config = {
                "email": {
                    "smtp_server": self.smtp_server.get(),
                    "smtp_port": int(self.smtp_port.get()),
                    "sender_email": self.sender_email.get(),
                    "app_password": self.app_password.get()
                },
                "email_template": {
                    "subject": self.email_subject.get(),
                    "body": self.email_body.get('1.0', 'end-1c')
                },
                "recipients": list(self.recipients_listbox.get(0, 'end')),
                "schedule": {
                    "times": list(self.schedule_listbox.get(0, 'end')),
                    "days": [day for day, var in self.days_vars.items() if var.get()]
                }
            }
            
            with open(self.automation.config_file, 'w') as f:
                json.dump(config, f, indent=4)
            
            self.automation.config = config
            messagebox.showinfo("Success", "Settings saved successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save settings: {str(e)}")

    def add_recipient(self):
        email = self.new_recipient.get().strip()
        if email and self.automation.validate_email(email):
            self.recipients_listbox.insert('end', email)
            self.new_recipient.delete(0, 'end')
        else:
            messagebox.showerror("Error", "Please enter a valid email address")

    def remove_recipient(self):
        selection = self.recipients_listbox.curselection()
        if selection:
            self.recipients_listbox.delete(selection)

    def add_schedule_time(self):
        time = self.schedule_time.get().strip()
        try:
            datetime.strptime(time, "%H:%M")
            self.schedule_listbox.insert('end', time)
            self.schedule_time.delete(0, 'end')
        except ValueError:
            messagebox.showerror("Error", "Please enter time in HH:MM format")

    def remove_schedule_time(self):
        selection = self.schedule_listbox.curselection()
        if selection:
            self.schedule_listbox.delete(selection)

    def start_automation(self):
        if not self.is_running:
            self.save_config()  # Save current settings before starting
            
            if not self.automation.config["email"]["sender_email"] or not self.automation.config["email"]["app_password"]:
                messagebox.showerror("Error", "Please configure email credentials first")
                return
                
            if not self.automation.config["recipients"]:
                messagebox.showerror("Error", "Please add at least one recipient")
                return
                
            if not self.automation.config["schedule"]["times"]:
                messagebox.showerror("Error", "Please add at least one schedule time")
                return
                
            if not any(self.automation.config["schedule"]["days"]):
                messagebox.showerror("Error", "Please select at least one day")
                return
            
            self.is_running = True
            self.automation.is_running = True
            self.automation_thread = threading.Thread(target=self.automation.start)
            self.automation_thread.daemon = True
            self.automation_thread.start()
            
            self.start_button.configure(state='disabled')
            self.stop_button.configure(state='normal')
            logging.info("Email automation started")

    def stop_automation(self):
        if self.is_running:
            self.is_running = False
            self.automation.stop()
            self.start_button.configure(state='normal')
            self.stop_button.configure(state='disabled')
            logging.info("Email automation stopped")

def main():
    root = tk.Tk()
    app = EmailAutomationGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main() 