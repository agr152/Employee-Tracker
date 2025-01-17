#1 Import necessary modules
import customtkinter as CTk
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from tkcalendar import DateEntry
import pymysql
import csv
import sys
import customtkinter as CTk
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import pymysql
import sys
import smtplib
import os
from dotenv import load_dotenv
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import random
import string
import hashlib
import re
from datetime import datetime
import json
import time
from pynput import mouse
from tkinter import TclError

#https://stackoverflow.com/questions/31836104/pyinstaller-and-onefile-how-to-include-an-image-in-the-exe-file
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def validate_login():
    global user_session
    global username_entry, password_entry
    username = username_entry.get()
    password = password_entry.get()
    try:
        connection = pymysql.connect(
        host='Enter your Database host name',
        port=3306,#<-----this is the default port for DB connections you can replace if you have diffrent port number
        username='Enter your username for the database',
        password='Enter your password for the database',
        database='Enter your database name'
        )

        cursor = connection.cursor()
        
        # Check if the user exists and retrieve their status
        status_query = "SELECT status FROM users WHERE username = %s"
        cursor.execute(status_query, (username,))
        status_result = cursor.fetchone()

        if status_result:
            status = status_result[0]
            if status == "inactive":
                messagebox.showerror("Login Failed", "Your account is inactive. Please contact the administrator.")
                return
        else:
            messagebox.showerror("Login Failed", "Invalid username or password. Please try again.")
            return

        # Verify username and password if status is active
        query = "SELECT id, name, role FROM users WHERE username = %s AND password = %s"
        cursor.execute(query, (username, generate_hash(password)))
        result = cursor.fetchone()

        if result:
            user_id, user_name, role = result
            user_session = {"id": user_id, "name": user_name, "username": username, "role": role}
            login_window.destroy()

            if role == "admin":
                show_timer_admin(username)
            elif role == "user":
                show_timer(username)
            else:
                messagebox.showerror("Role Error", f"Unexpected role: {role}")
        else:
            messagebox.showerror("Login Failed", "Invalid username or password. Please try again.")

        cursor.close()
        connection.close()
    except pymysql.Error as err:
        messagebox.showerror("Database Error", f"Error: {err}")


def is_valid_email(to_email):
    """Validate email address for @gineesoft.com domain."""
    email_regex = r'^[a-zA-Z0-9._%+-]+@gineesoft\.com$'
    return re.match(email_regex, to_email) is not None

def is_valid_contact_number(number):
    """Validate contact number to ensure it contains exactly 10 digits."""
    contact_regex = r'^\d{10}$'
    return re.match(contact_regex, number) is not None

#4 Function to send email to the New user account
def register_employee(full_name, username, to_email, contact, status, role):
    if not full_name or not username or not to_email or not contact:
        messagebox.showerror("Registration Error", "Please fill in all fields!")
        return
    elif not is_valid_contact_number(contact):
        print("Invalid contact number.")
        messagebox.showerror("Contact Error", "The provided contact number is invalid.")
        return False

    elif not is_valid_email(to_email):
        print("Invalid email address.")
        messagebox.showerror("Email Error", "The provided email address is invalid.")
        return False

    conn = pymysql.connect(
        host='Enter your Database host name',
        port=3306,
        user='Enter your database username',
        password='Enter your database password',
        database='enter your database name'
    )
    if not conn:
        return

    cursor = conn.cursor()

    try:
        cursor.execute("SELECT username FROM users WHERE username = %s", (username,))
        if cursor.fetchone():
            messagebox.showerror("Registration Error", "User already exists!")
            return

        password = generate_random_password()

        # Insert user data into the database
        insert_query = "INSERT INTO users (name, username, password, email, contact, status, role) VALUES (%s, %s, %s, %s, %s, %s, %s)"
        cursor.execute(insert_query, (capitalize_words(full_name), username, generate_hash(password), to_email, contact, status, role))
        
        # Attempt to send the email
        try:
            send_email(to_email, username, password)
        except Exception as e:
            # Rollback the transaction if email sending fails
            conn.rollback()
            messagebox.showerror("Email Error", f"Failed to send email: {e}")
            return

        # Commit the transaction if email sending is successful
        conn.commit()
        messagebox.showinfo("Registration Success", "User registered successfully!")

        # Clear the entry fields
        reg_full_name_entry.delete(0, CTk.END)
        reg_username_entry.delete(0, CTk.END)
        reg_email_entry.delete(0, CTk.END)
        reg_contact_entry.delete(0, CTk.END)
    except pymysql.MySQLError as err:
        conn.rollback()  # Rollback in case of any database error
        messagebox.showerror("Database Error", f"Error during registration: {err}")
    finally:
        cursor.close()
        conn.close()


def send_email(to_email, username, password):
    load_dotenv()
    # Email configuration
    sender_email = os.getenv("EMAIL")  # Replace with your email in .env file in root directory
    sender_password = os.getenv("PASSWORD")  # Replace with your email password in .env file in root directory
    smtp_server = ""  # <----Enter the SMTP server (e.g., for Gmail)
    smtp_port = ""  # <---- Enter the SMTP port for TLS

    try:
        # Create the email
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = to_email
        message["Subject"] = "Welcome to GineeSoft - Your Credentials"
        body = f"""
        Dear User,

        Your account has been successfully created. Below are your login credentials:

        Username: {username}
        Password: {password}

        Please keep this information secure and do not share it with anyone.

        Regards,
        Admin Team
        """
        message.attach(MIMEText(body, "plain"))

        # Connect to the SMTP server and send the email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Upgrade to a secure connection
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, to_email, message.as_string())

        print("Email sent successfully.")
    except Exception as e:
        print(f"Failed to send email: {e}")
        raise e  # Raise the exception so the calling function can handle it




#6 download attendance by date range
def download_attendance_by_date_range(from_date, to_date, selected_option):
    conn = pymysql.connect(
        host='Enter your Database host name',
        port=3306,
        user='Enter your database username',
        password='Enter your database password',
        database='enter your database name'
)
    if not conn:
        return

    cursor = conn.cursor()
    try:
        if selected_option == "Active":
            # Query for active users
            query = """
                SELECT users.name, attendance.date, attendance.punch_in_time, 
                       attendance.breaktime, attendance.punch_out_time, attendance.total_time
                FROM users
                INNER JOIN attendance ON users.username = attendance.username
                WHERE attendance.date BETWEEN %s AND %s
                  AND users.status = %s
                ORDER BY users.name, attendance.date
            """
            cursor.execute(query, (from_date, to_date, selected_option))
        else:
            # Query for inactive users with additional columns (email and contact)
            query = """
                SELECT users.name, users.email, users.contact, attendance.date, attendance.punch_in_time, 
                       attendance.breaktime, attendance.punch_out_time, attendance.total_time
                FROM users
                INNER JOIN attendance ON users.username = attendance.username
                WHERE attendance.date BETWEEN %s AND %s
                  AND users.status = %s
                ORDER BY users.name, attendance.date
            """
            cursor.execute(query, (from_date, to_date, selected_option))

        records = cursor.fetchall()

        if not records:
            messagebox.showinfo("No Data", "No attendance records found for the selected date range.")
            return

        # Open a file dialog to save the file
        file_path = filedialog.asksaveasfilename(
            title="Save Attendance Record",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )

        if file_path:
            with open(file_path, mode="w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)

                # Write header row
                if selected_option == "Active":
                    writer.writerow(["Name", "Date", "Punch In Time", "Break Time", "Punch Out Time", "Total Time"])
                else:
                    writer.writerow(["Name", "Email", "Contact", "Date", "Punch In Time", "Break Time", "Punch Out Time", "Total Time"])

                # Write data rows
                current_name = None

                for record in records:
                    if selected_option == "Active":
                        name, date, punch_in_time, break_time, punch_out_time, total_time = record
                        if name != current_name:
                            writer.writerow([name])
                            current_name = name
                        writer.writerow(["", date, punch_in_time, break_time, punch_out_time, total_time])
                    else:
                        name, email, contact, date, punch_in_time, break_time, punch_out_time, total_time = record
                        if name != current_name:
                            writer.writerow([name, email, contact])
                            current_name = name
                        writer.writerow(["", "", "", date, punch_in_time, break_time, punch_out_time, total_time])

            messagebox.showinfo("Download Success", "Attendance records have been saved successfully!")

    except pymysql.MySQLError as err:
        messagebox.showerror("Database Error", f"Error: {err}")
    finally:
        cursor.close()
        conn.close()

def on_update_button_click(update_button):
    # Disable the button immediately
    update_button.configure(state="disabled")  # Disable the button to prevent multiple clicks

    # Get the input values
    full_name = reg_full_name_entry.get()
    username = reg_username_entry.get()
    role = dropdown.get()
    email = reg_email_entry.get()
    contact = reg_contact_entry.get()

    # Simulate the update process (replace with your actual function)
    result = update_employee(full_name, username, email, contact, role)

    # Re-enable the button after the operation is finished
    update_button.configure(state="normal")  # Re-enable the button


def update_employee(full_name, username, email, contact, role):
    if not full_name or not username or not email or not contact:
        messagebox.showerror("Update Error", "Please fill in all fields!")
        return "Failed: Missing required fields."
    elif not is_valid_contact_number(contact):
        messagebox.showerror("Contact Error", "The provided contact number is invalid.")
        return "Failed: Invalid contact number."
    elif not is_valid_email(email):
        messagebox.showerror("Email Error", "The provided email address is invalid.")
        return "Failed: Invalid email address."

    conn = pymysql.connect(
        host='Enter your Database host name',
        port=3306,
        user='Enter your database username',
        password='Enter your database password',
        database='enter your database name'
    )
    if not conn:
        return "Failed: Unable to connect to the database."

    cursor = conn.cursor()

    try:
        # Check if the user exists
        cursor.execute("SELECT username FROM users WHERE username = %s", (username,))
        if not cursor.fetchone():
            messagebox.showerror("Update Error", "User does not exist!")
            return "Failed: User not found."

        # Update the user's information
        update_query = """
            UPDATE users
            SET name = %s, email = %s, contact = %s, role = %s
            WHERE username = %s
        """
        cursor.execute(update_query, (capitalize_words(full_name), email, contact, role, username))
        conn.commit()

        messagebox.showinfo("Update Success", "User information updated successfully!")
        
        # Clear the entry fields
        reg_full_name_entry.delete(0, CTk.END)
        reg_username_entry.delete(0, CTk.END)
        reg_email_entry.delete(0, CTk.END)
        reg_contact_entry.delete(0, CTk.END)
        return "Success: User information updated."
    except pymysql.MySQLError as err:
        messagebox.showerror("Database Error", f"Error during update: {err}")
        return f"Failed: {err}"
    finally:
        cursor.close()
        conn.close()

def show_update_employee_window(admin_window, selected_employee):
    # Check if the update window is already open
    if hasattr(show_update_employee_window, "update_window") and show_update_employee_window.update_window.winfo_exists():
        # If the window is already open, bring it to the front and focus on it
        show_update_employee_window.update_window.lift()
        show_update_employee_window.update_window.focus_force()
        return

    # Create the update window if it's not already open
    update_window = CTk.CTkToplevel(admin_window)
    update_window.title("Update Employee Information")
    update_window.geometry("400x300")  # Adjust window size
    update_window.resizable(False, False)  # Prevent resizing

    # Extract details from the selected employee
    full_name = selected_employee[0]  # Assuming Full Name is the first column
    username = selected_employee[1]   # Assuming Username is the second column
    role = selected_employee[5]       # Assuming Role is the third column
    email = selected_employee[2]      # Assuming Email is the fourth column
    contact = selected_employee[3]    # Assuming Contact is the fifth column

    # Full Name Entry
    CTk.CTkLabel(update_window, text="Full Name:").place(x=50, y=20)
    global reg_full_name_entry
    reg_full_name_entry = CTk.CTkEntry(update_window,width=180)
    reg_full_name_entry.place(x=180, y=20)
    reg_full_name_entry.insert(0, full_name)  # Pre-fill the full name

    # Employee Username Entry
    CTk.CTkLabel(update_window, text="Username:").place(x=50, y=60)
    global reg_username_entry
    reg_username_entry = CTk.CTkEntry(update_window,width=180)
    reg_username_entry.place(x=180, y=60)
    reg_username_entry.insert(0, username)  # Pre-fill the username
    reg_username_entry.configure(state="disabled")  # Make username read-only (optional)

    # Role Entry
    options = ["user", "admin"]
    CTk.CTkLabel(update_window, text="Role:").place(x=50, y=100)
    global dropdown
    dropdown = CTk.CTkComboBox(update_window, values=options,width=180)
    dropdown.set(role)  # Pre-fill the role
    dropdown.place(x=180, y=100)

    # Email Entry
    CTk.CTkLabel(update_window, text="Email:").place(x=50, y=140)
    global reg_email_entry
    reg_email_entry = CTk.CTkEntry(update_window,width=180)
    reg_email_entry.place(x=180, y=140)
    reg_email_entry.insert(0, email)  # Pre-fill the email

    # Contact Entry
    CTk.CTkLabel(update_window, text="Contact:").place(x=50, y=180)
    global reg_contact_entry
    reg_contact_entry = CTk.CTkEntry(update_window,width=180)
    reg_contact_entry.place(x=180, y=180)
    reg_contact_entry.insert(0, contact)  # Pre-fill the contact

    # Update Button (to update the employee information)
    update_button = CTk.CTkButton(
        update_window,
        text="Update",
        command=lambda: on_update_button_click(update_button)
    )
    update_button.place(x=125, y=240)

    # Store reference to this window for future checks
    show_update_employee_window.update_window = update_window

    # Bring the window to the front and focus on it
    update_window.lift()
    update_window.focus_force()
    update_window.attributes("-topmost", True)




def capitalize_words(*strings):
    capitalized_strings = []
    for string in strings:
        # Remove any special characters and split the string into words
        cleaned_string = re.sub(r'[^a-zA-Z\s]', '', string)
        capitalized_string = " ".join([word.capitalize() for word in cleaned_string.split()])
        capitalized_strings.append(capitalized_string)
    return capitalized_strings

def register_employee(full_name, username, email,contact,status,role):
    # Simulating a registration process (replace with actual registration logic)
    time.sleep(2)  # Simulate a delay for registration (2 seconds)
    return "Registration Successful!"  # Success message

def on_add_button_click(add_button, loading_label):
    # Disable the button immediately
    add_button.configure(state="disabled")  # Use configure() instead of config()

    # Show loading indicator
    loading_label.place(x=125, y=210)  # Position the loading label

    # Get the input values
    full_name = reg_full_name_entry.get()
    username = reg_username_entry.get()
    status = "active"
    role = dropdown.get()
    email = reg_email_entry.get()
    contact = reg_contact_entry.get()

    # Simulate the registration process (replace with your actual function)
    result = register_employee(full_name, username, email, contact,status,role)

    # Hide the loading label and show success message
    loading_label.place_forget()
    

    # Re-enable the button after the operation is finished
    add_button.configure(state="normal")  # Use configure() instead of config()

def show_register_employee_window(admin_window):
    # Check if the register window is already open
    if hasattr(show_register_employee_window, "register_window") and show_register_employee_window.register_window.winfo_exists():
        # If the window is already open, bring it to the front and focus on it
        show_register_employee_window.register_window.lift()  # Bring the window to the front
        show_register_employee_window.register_window.focus_force()  # Focus on the window
        return

    # Create the register window if it's not already open
    register_window = CTk.CTkToplevel(admin_window)
    register_window.title("Register New Employee")
    register_window.geometry("400x300")  # Adjust window size
    register_window.resizable(False, False)  # Prevent resizing

    # Full Name Entry
    CTk.CTkLabel(register_window, text="Full Name:").place(x=50, y=20)
    global reg_full_name_entry
    reg_full_name_entry = CTk.CTkEntry(register_window,width=180)
    reg_full_name_entry.place(x=180, y=20)

    # Employee Username Entry
    CTk.CTkLabel(register_window, text="Username:").place(x=50, y=60)
    global reg_username_entry
    reg_username_entry = CTk.CTkEntry(register_window,width=180)
    reg_username_entry.place(x=180, y=60)

    #role entry
    options = ["user", "admin"]
    CTk.CTkLabel(register_window, text="Role:").place(x=50, y=100)
    global dropdown
    dropdown = CTk.CTkComboBox(register_window, values=options,width=180)
    dropdown.set(options[0])
    dropdown.place(x=180, y=100)

    # Email Entry
    CTk.CTkLabel(register_window, text="Email:").place(x=50, y=140)
    global reg_email_entry
    reg_email_entry = CTk.CTkEntry(register_window,width=180)
    reg_email_entry.place(x=180, y=140)

    # contact Entry
    CTk.CTkLabel(register_window, text="Contact:").place(x=50, y=180)
    global reg_contact_entry
    reg_contact_entry = CTk.CTkEntry(register_window,width=180)
    reg_contact_entry.place(x=180, y=180)

    # Add Button (to register the employee)
    add_button = CTk.CTkButton(
        register_window,
        text="Add",
        command=lambda: on_add_button_click(add_button, loading_label)
    )
    add_button.place(x=125, y=240)

    # Loading label (initially hidden)
    loading_label = CTk.CTkLabel(register_window, text="Loading...", text_color="gray")
    loading_label.place_forget()  # Hide initially

    # Success message label (initially hidden)
    
    # Store reference to this window for future checks
    show_register_employee_window.register_window = register_window

    # Bring the window to the front and focus on it
    register_window.lift()  # Make sure it's on top of other windows
    register_window.focus_force()  # Focus on the window to ensure it's active
    register_window.attributes("-topmost", True)  # Ensure the window is on top of all other windows




#10 Function to register a new employee
def register_employee(full_name, username, to_email, contact, status, role):
    if not full_name or not username or not to_email or not contact:
        messagebox.showerror("Registration Error", "Please fill in all fields!")
        return
    elif not is_valid_contact_number(contact):
        print("Invalid contact number.")
        messagebox.showerror("Contact Error", "The provided contact number is invalid.")
        return False

    elif not is_valid_email(to_email):
        print("Invalid email address.")
        messagebox.showerror("Email Error", "The provided email address is invalid.")
        return False

    conn = pymysql.connect(
        host='Enter your Database host name',
        port=3306,
        user='Enter your database username',
        password='Enter your database password',
        database='enter your database name'
    )
    if not conn:
        return

    cursor = conn.cursor()

    try:
        cursor.execute("SELECT username FROM users WHERE username = %s", (username,))
        if cursor.fetchone():
            messagebox.showerror("Registration Error", "User already exists!")
            return

        password = generate_random_password()

        # Insert user data into the database
        insert_query = "INSERT INTO users (name, username, password, email, contact, status, role) VALUES (%s, %s, %s, %s, %s, %s, %s)"
        cursor.execute(insert_query, (capitalize_words(full_name), username, generate_hash(password), to_email, contact, status, role))
        
        # Attempt to send the email
        try:
            send_email(to_email, username, password)
        except Exception as e:
            # Rollback the transaction if email sending fails
            conn.rollback()
            messagebox.showerror("Email Error", f"Failed to send email: {e}")
            return

        # Commit the transaction if email sending is successful
        conn.commit()
        messagebox.showinfo("Registration Success", "User registered successfully!")

        # Clear the entry fields
        reg_full_name_entry.delete(0, CTk.END)
        reg_username_entry.delete(0, CTk.END)
        reg_email_entry.delete(0, CTk.END)
        reg_contact_entry.delete(0, CTk.END)
    except pymysql.MySQLError as err:
        conn.rollback()  # Rollback in case of any database error
        messagebox.showerror("Database Error", f"Error during registration: {err}")
    finally:
        cursor.close()
        conn.close()

#11 fetching details
def fetch_employee_names(tree):
    conn = pymysql.connect(
        host='Enter your Database host name',
        port=3306,
        user='Enter your database username',
        password='Enter your database password',
        database='enter your database name'
    )
    if not conn:
        return

    cursor = conn.cursor()
    try:
        # Fetch both name and email of employees
        query = "SELECT name,username, email,contact,status,role FROM users WHERE status = 'active'"
        cursor.execute(query)
        records = cursor.fetchall()

        # Clear the Treeview before inserting new data
        for row in tree.get_children():
            tree.delete(row)

        # Insert fetched data into the Treeview
        for record in records:
            tree.insert("", "end", values=(record[0], record[1], record[2], record[3],record[4],record[5]))  
    except pymysql.MySQLError as err:
        messagebox.showerror("Database Error", f"Error: {err}")
    finally:
        cursor.close()
        conn.close()

#12 function to show attendance detials
def show_attendance_details(selected_name):
    conn = pymysql.connect(
        host='Enter your Database host name',
        port=3306,
        user='Enter your database username',
        password='Enter your database password',
        database='enter your database name'
)
    if not conn:
        return

    cursor = conn.cursor()
    try:
        # Fetch attendance details for the selected user
        query = """
        SELECT attendance.date, attendance.punch_in_time, attendance.breaktime, attendance.punch_out_time, attendance.total_time
        FROM users
        INNER JOIN attendance ON users.username = attendance.username
        WHERE users.name = %s
        """
        cursor.execute(query, (selected_name,))
        records = cursor.fetchall()

        if not records:
            messagebox.showinfo("Attendance Details", f"No attendance records found for {selected_name}.")
            return

        # Display attendance details in a new window
        attendance_window = CTk.CTkToplevel()
        attendance_window.title(f"Attendance Details - {selected_name}")
        attendance_window.geometry("600x400")

        # Define style for Attendance Treeview
        style = ttk.Style()
        style.configure(
            "Treeview",
            rowheight=25,  # Adjust row height
            borderwidth=1,
            relief="solid",
        )
        style.configure(
            "Treeview.Heading",
            font=("Arial", 12, "bold"),
            anchor="center"  # Center-align header text
        )
        style.map("Treeview", background=[("selected", "#347083")])  # Highlight row color on selection

        # Treeview for attendance details
        columns = ("Date", "Push In Time", "Break Time", "Push Out Time","Total Time")
        tree = ttk.Treeview(attendance_window, columns=columns, show="headings", height=10)
        for col in columns:
            tree.heading(col, text=col)  # Set column headings
            tree.column(col, anchor="center", width=150)  # Center-align content and set width
            tree.pack(pady=10, fill="both", expand=True)

        # Insert attendance records
        for record in records:
            tree.insert("", "end", values=record)

        # Download button
        download_button = CTk.CTkButton(
            attendance_window,
            text="Download Attendance",
            command=lambda: download_attendance(selected_name, records)
        )
        download_button.pack(pady=10)

    except pymysql.MySQLError as err:
        messagebox.showerror("Database Error", f"Error: {err}")
    finally:
        cursor.close()
        conn.close()


#13 downloading attendance for particular employee
def download_attendance(selected_name, records):
    # Open a file dialog to save the file
    file_path = filedialog.asksaveasfilename(
        title="Save Attendance Record",
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )

    if file_path:
        try:
            # Write records to CSV
            with open(file_path, mode="w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                # Write header
                writer.writerow(["Date", "Push In Time", "Break Time", "Push Out Time", "Total Time"])
                # Write attendance records
                writer.writerows(records)

            messagebox.showinfo("Download Success", f"Attendance record for {selected_name} has been saved successfully!")
        except Exception as e:
            messagebox.showerror("Download Error", f"An error occurred while saving the file: {e}")

#admin dashboard
def show_admin_dashboard():
    admin_window = CTk.CTkToplevel(anilbran_att_ad_user)
    admin_window.title("Admin Dashboard")
    admin_window.geometry("800x600")

    #delete employee functionality
    def inactivate_employee():
        connection = pymysql.connect(
        host='Enter your Database host name',
        port=3306,
        user='Enter your database username',
        password='Enter your database password',
        database='enter your database name'
)
        if not connection:
            return
        selected_item = tree.selection()  # Get the selected item
        if not selected_item:
            messagebox.showinfo("Error", "Please select an employee to inactivate.")
            return

        selected_employee = tree.item(selected_item, "values")  # Get the selected employee's details
        if not selected_employee:
            messagebox.showinfo("Error", "No employee selected.")
            return

        username = selected_employee[1]  # Assuming "Full Name" is the first column
        confirmation = messagebox.askyesno(
            "Confirm inactivation",
            f"Are you sure you want to inactivate the {username}?"
        )
        if not confirmation:
            return

        try:
            # Delete the employee from the users table
            with connection.cursor() as cursor:
                cursor.execute("UPDATE users SET status = %s WHERE username = %s", ('inactive', username))
                connection.commit()

            messagebox.showinfo("Success", f"Details of {username} have been inactivated.")
            refresh_treeview()  # Refresh the Treeview to reflect the changes
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while inactivating the employee: {e}")

        # Open the update employee window with the selected employee's details

    def redirect_to_update_window():
    # Get the selected item from the Treeview
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showinfo("Error", "Please select an employee to update.")
            return

        selected_employee = tree.item(selected_item, "values")  # Get the selected employee's details
        if not selected_employee:
            messagebox.showinfo("Error", "No employee selected.")
            return
        show_update_employee_window(admin_window, selected_employee)

    main_frame = CTk.CTkFrame(admin_window)
    main_frame.pack(side=CTk.LEFT, fill=CTk.BOTH, expand=True, padx=5, pady=5)

    # Left side frame for Treeview
    left_frame = CTk.CTkFrame(main_frame)
    left_frame.pack(side=CTk.LEFT, fill=CTk.BOTH, expand=True, padx=10, pady=10)

    # Right side frame for buttons and date picker - Increase width by adjusting width param
    right_frame = CTk.CTkFrame(main_frame)
    right_frame.pack(side=CTk.RIGHT, fill=CTk.Y, padx=10, pady=10, ipadx=70)  # Increased ipadx for more internal padding

    # Admin Dashboard title
    CTk.CTkLabel(left_frame, text="Admin Dashboard", font=("Arial", 14)).pack(pady=10)

    # Treeview Section
    CTk.CTkLabel(left_frame, text="Employee List", font=("Arial", 14)).pack(pady=10)

    # Define style for Treeview
    style = ttk.Style()
    style.configure(
        "Treeview",
        rowheight=20,  # Reduced row height
        borderwidth=1,
        relief="solid"
    )
    style.configure(
        "Treeview.Heading",
        font=("Arial", 12, "bold"),
        anchor="center"  
    )
    style.map("Treeview", background=[("selected", "#347083")])  # Highlight row color on selection

    # Define columns and create the Treeview with a smaller height and adjusted column widths
    tree = ttk.Treeview(left_frame, columns=("Full Name","User Name", "Email","Contact","Status","Role"), show="headings", height=6)  # Reduced height (6 rows)
    tree.heading("Full Name", text="Full Name")
    tree.heading("User Name", text="User Name")
    tree.heading("Email", text="Email")
    tree.heading("Contact", text="Contact")
    tree.heading("Status", text="Status")
    tree.heading("Role", text="Role")


    # Adjust column width for a more compact view
    tree.column("Full Name", anchor="center", width=150)  # Reduced width
    tree.column("User Name", anchor="center", width=150)  # Reduced width
    tree.column("Email", anchor="center", width=200)  # Reduced width
    tree.column("Contact", anchor="center", width=250)  # Reduced width
    tree.column("Status", anchor="center", width=250)  # Reduced width
    tree.column("Role", anchor="center", width=150)  # Reduced width
    tree.pack(pady=10, fill="both", expand=True)

    # Bind click event to show attendance details
    tree.bind("<Double-1>", lambda event: show_attendance_details(tree.item(tree.focus())["values"][0]))

    # Fetch and display employee list immediately when the page is loaded
    fetch_employee_names(tree)

    # Right side: Button to register a new employee
    

    # Date range selection
    CTk.CTkLabel(right_frame, text="From Date:", font=("Arial", 12)).pack(pady=10)
    
    # Enhanced DateEntry for "From Date"
    from_date_entry = DateEntry(right_frame, date_pattern='dd-mm-yyyy', width=15, font=('Arial', 12), background='lightblue', foreground='black')
    from_date_entry.pack(pady=5)

    CTk.CTkLabel(right_frame, text="To Date:", font=("Arial", 12)).pack(pady=10)
    
    # Enhanced DateEntry for "To Date"
    to_date_entry = DateEntry(right_frame, date_pattern='dd-mm-yyyy', width=15, font=('Arial', 12), background='lightblue', foreground='black')
    to_date_entry.pack(pady=5)

    selected_option = CTk.StringVar(value="Active")
    radio_active = CTk.CTkRadioButton(right_frame, text="Active", variable=selected_option, value="Active")
    radio_active.place(x=40,y=190)
    radio_inactive = CTk.CTkRadioButton(right_frame, text="Inactive", variable=selected_option, value="Inactive",fg_color="#D50F10")
    radio_inactive.place(x=140,y=190)

    

    # Download attendance button
    download_button = CTk.CTkButton(
        right_frame,
        text="Download Attendance",
        command=lambda: download_attendance_by_date_range(from_date_entry.get(), to_date_entry.get() ,selected_option.get())
    )
    download_button.place(x=50,y=250)

    register_button = CTk.CTkButton(
        right_frame, text="Register New Employee", command=lambda: show_register_employee_window(admin_window)
    )
    register_button.place(x=47,y=350)

    # Update Employee button
    update_button = CTk.CTkButton(
        right_frame,
        text="Update Employee",fg_color="#65C73E",
        command=redirect_to_update_window  # Reference the function here
    )
    update_button.place(x=50,y=410)

    inactive_button = CTk.CTkButton(
        right_frame,
        text="Inactivate Employee",fg_color="#D50F10",
        command=inactivate_employee  # Reference the function here
    )
    inactive_button.place(x=50,y=470)

    # Refresh Button
    def refresh_treeview():
        # Clear the existing entries in the Treeview
        for item in tree.get_children():
            tree.delete(item)

        # Fetch and display the updated employee list
        fetch_employee_names(tree)

    # Add Refresh button below the Treeview
    refresh_button = CTk.CTkButton(
        left_frame,
        text="Refresh Employee List",
        command=refresh_treeview
    )
    refresh_button.pack(pady=10)







def start_timer():
    global start_time, timer_running, last_activity_time, break_mode, punch_in_time, break_time
    if not timer_running:
        timer_running = True
        break_mode = False
        break_time = "No Break"  # Default value 0 if no break
        start_time = time.time()
        last_activity_time = time.time()
        punch_in_time = time.strftime("%I:%M:%S %p")
        update_timer()
        monitor_inactivity()
        root.iconify()
        
def calculate_hours_between(time1: str, time2: str) -> str:
    """
    Calculate the total hours between two time values and return as a string.
    
    Args:
        time1 (str): The first time in string format (e.g., "10:51:20 AM").
        time2 (str): The second time in string format (e.g., "03:15:45 PM").
    
    Returns:
        str: The total hours between the two times, formatted as a string.
    """
    # Define the fixed time format
    time_format = "%I:%M:%S %p"
    
    # Parse the time strings into datetime objects
    t1 = datetime.strptime(time1, time_format)
    t2 = datetime.strptime(time2, time_format)
    
    # Calculate the difference in hours
    time_difference = abs((t2 - t1).total_seconds() / 3600)
    
    # Format and return the result as a string
    return f"{time_difference:.2f} hours"
def stop_timer():
    global timer_running, punch_out_time, date, break_time
    if timer_running:
        timer_running = False
        punch_out_time = time.strftime("%I:%M:%S %p")
        date = time.strftime("%d-%m-%Y")
        
        try:
            connection = pymysql.connect(
                host='Enter your Database host name',
                port=3306,
                user='Enter your database username',
                password='Enter your database password',
                database='enter your database name'
            )
            cursor = connection.cursor()
            query = """
                INSERT INTO attendance (username, date, punch_in_time, breaktime, punch_out_time, total_time) 
                VALUES (%s, %s, %s, %s, %s, %s)
            """
            cursor.execute(query, ( 
                user_session["username"], 
                date, 
                punch_in_time,
                break_time, 
                punch_out_time,
                calculate_hours_between(punch_in_time, punch_out_time)
            ))
            connection.commit()

            messagebox.showinfo("Punch Out", f"Punch In Time: {punch_in_time}\nPunch Out Time: {punch_out_time}\nRecord saved to database.")
            cursor.close()
            connection.close()

        except pymysql.Error as err:
            messagebox.showerror("Database Error", f"Error: {err}")

        finally:
            if root and root.winfo_exists():
                root.destroy()
                sys.exit()

def monitor_inactivity():
    global last_activity_time, break_mode
    if timer_running and not break_mode:
        try:
            if root and root.winfo_exists():  # Ensure root is valid and exists
                if time.time() - last_activity_time > 900:  # 15 Minutes of inactivity
                    ask_continue_or_exit()
                else:
                    root.after(1000, monitor_inactivity)
        except Exception as e:
            print(f"Error in monitor_inactivity: {e}")

def update_timer():
    if timer_running:
        try:
            if root and root.winfo_exists():  # Ensure root is valid and exists
                elapsed_time = time.time() - start_time
                hours, remainder = divmod(int(elapsed_time), 3600)
                minutes, seconds = divmod(remainder, 60)
                timer_label.configure(text=f"{hours:02}:{minutes:02}:{seconds:02}")
                root.after(1000, update_timer)
        except Exception as e:
            print(f"Error in update_timer: {e}")


def ask_continue_or_exit():
    def continue_action():
        global last_activity_time
        last_activity_time = time.time()
        popup.destroy()
        monitor_inactivity()

    def exit_action():
        stop_timer()
        popup.destroy()
        root.destroy()

    def auto_exit():
        if popup.winfo_exists():
            stop_timer()
            popup.destroy()
            root.destory()

    popup = CTk.CTkToplevel(root)
    popup.title("Inactivity Detected")
    popup.geometry("300x100")
    popup.resizable(False,False)
    CTk.CTkLabel(popup, text="You have been inactive for too long").pack(pady=10)

    button_frame = CTk.CTkFrame(popup)
    button_frame.pack(pady=5)

    continue_button = CTk.CTkButton(button_frame, text="Continue", command=continue_action, fg_color="green", width=10)
    continue_button.grid(row=0, column=0, padx=5)

    exit_button = CTk.CTkButton(button_frame, text="Exit", command=exit_action, fg_color="red", width=10)
    exit_button.grid(row=0, column=1, padx=5)

    popup.after(15*60*1000, auto_exit)

def on_mouse_move(x, y):
    global last_activity_time
    last_activity_time = time.time()

def add_to_break():
    global break_mode, break_time
    break_mode = True
    messagebox.showinfo("Break Mode", "Break mode enabled for 45 minutes.")
    break_time = time.strftime("%I:%M:%S %p")
    root.after(45 * 60 * 1000, disable_break)  # 45 minutes in milliseconds

def disable_break():
    global break_mode
    break_mode = False
    messagebox.showinfo("Break Mode", "Break mode disabled. Inactivity detection is now active.")


#7 generate random password
def generate_random_password():
    # Define character sets
    letters = string.ascii_letters  # Uppercase and lowercase letters
    digits = string.digits  # Numbers 0-9
    special_chars = "!@$%&*"

    # Ensure password contains at least one letter, one digit, and one special character
    password = [
        random.choice(letters),
        random.choice(digits),
        random.choice(special_chars),
    ]

    # Fill the remaining length with random characters from all sets
    all_chars = letters + digits + special_chars
    password += random.choices(all_chars, k=3)  # Remaining 3 characters

    # Shuffle the list to ensure randomness and return as a string
    random.shuffle(password)
    return ''.join(password)

def generate_hash(input_string):
        # Ensure the input is a string
    if not isinstance(input_string, str):
        raise ValueError("Input must be a string.")

    # Create a SHA-256 hash object
    hash_object = hashlib.sha256()

    # Update the hash object with the input string encoded in UTF-8
    hash_object.update(input_string.encode('utf-8'))

    # Return the hexadecimal digest of the hash
    return hash_object.hexdigest()




def change_password():
    if hasattr(change_password, "change_window") and change_password.change_window.winfo_exists():
        # If the window is already open, bring it to the front and focus on it
        change_password.change_window.lift()  # Bring the window to the front
        change_password.change_window.focus_force()  # Focus on the window
        return
    def validate_and_update_password():
        username = username_entry.get()
        old_password = old_password_entry.get()
        new_password = new_password_entry.get()
        confirm_password = confirm_password_entry.get()
        if not username or not old_password or not new_password or not confirm_password:
            messagebox.showerror("Error", "Please fill out all fields.")
            return
        if new_password != confirm_password:
            messagebox.showerror("Error", "New passwords do not match.")
            return
        try:
            connection = pymysql.connect(
                host='Enter your Database host name',
                port=3306,
                user='Enter your database username',
                password='Enter your database password',
                database='enter your database name'
            )
            cursor = connection.cursor()

            # Verify the old password
            query = "SELECT password FROM users WHERE username = %s"
            cursor.execute(query, (username,))
            result = cursor.fetchone()

            if result and result[0] == generate_hash(old_password):
                # Update the password
                hashed_password = generate_hash(new_password)
                update_query = "UPDATE users SET password = %s WHERE username = %s"
                cursor.execute(update_query, (hashed_password, username))
                connection.commit()

                messagebox.showinfo("Success", "Password changed successfully. Please log in with your new password.")
                change_window.destroy()
            else:
                messagebox.showerror("Error", "Invalid username or old password.")

            cursor.close()
            connection.close()
        except pymysql.Error as err:
            messagebox.showerror("Database Error", f"Error: {err}")


    change_window = CTk.CTkToplevel()
    change_window.title("Change Password")
    change_window.geometry("380x270")
    change_window.resizable(False, False)

    CTk.CTkLabel(change_window, text="Change Password").place(x=150, y=20)
    CTk.CTkLabel(change_window, text="Enter your username:").place(x=50,y=60)
    username_entry = CTk.CTkEntry(change_window)
    username_entry.place(x=190,y=60)

    CTk.CTkLabel(change_window, text="Old Password:").place(x=50,y=100)
    old_password_entry = CTk.CTkEntry(change_window, show="*")
    old_password_entry.place(x=190,y=100)

    CTk.CTkLabel(change_window, text="New Password:").place(x=50,y=140)
    new_password_entry = CTk.CTkEntry(change_window, show="*")
    new_password_entry.place(x=190,y=140)

    CTk.CTkLabel(change_window, text="Confirm Password:").place(x=50,y=180)
    confirm_password_entry = CTk.CTkEntry(change_window, show="*")
    confirm_password_entry.place(x=190,y=180)

    change_button = CTk.CTkButton(change_window, text="Change Password", command=validate_and_update_password)
    change_button.place(x=120,y=230)

    change_window.lift()  # Make sure it's on top of other windows
    change_window.focus_force()  # Focus on the window to ensure it's active
    change_window.attributes("-topmost", True)  # Ensure the window is on top of all other windows



def show_timer_admin(username):
    global root, timer_running, start_time, last_activity_time, timer_label, break_mode, punch_in_time, punch_out_time, listener

    root = CTk.CTk()
    root.title("Punch In Timer")
    root.geometry("250x280")
    root.resizable(False, False)
    
    timer_running = False
    start_time = 0
    last_activity_time = 0
    break_mode = False
    punch_in_time = ""
    punch_out_time = ""

    title_label = CTk.CTkLabel(root, text="Punch In Timer", font=("Helvetica", 16))
    title_label.pack(pady=10)

    user_label = CTk.CTkLabel(root, text=f"Hi, {user_session['name']}", font=("Helvetica", 12))
    user_label.pack(pady=5)

    timer_label = CTk.CTkLabel(root, text="00:00:00", font=("Helvetica", 24))
    timer_label.pack(pady=10)

    button_frame = CTk.CTkFrame(root)
    button_frame.pack(pady=10)

    punch_in_button = CTk.CTkButton(button_frame, text="Punch In ", command=start_timer, fg_color="#009900", width=10)
    punch_in_button.grid(row=0, column=0, padx=5)

    punch_out_button = CTk.CTkButton(button_frame, text="Punch Out", command=stop_timer, fg_color="#ed2939", width=10)
    punch_out_button.grid(row=0, column=1, padx=5)

    break_button = CTk.CTkButton(root, text="Add to Break", command=add_to_break, fg_color="blue", width=150)
    break_button.pack(pady=10)

    admin_button = CTk.CTkButton(root, text="Admin", command=show_admin_dashboard, fg_color="#00bfff",text_color="black", width=150)
    admin_button.pack(pady=10)

    

    listener = mouse.Listener(on_move=on_mouse_move)
    listener.start()

    def on_close():
        global root
        if listener:
            listener.stop()  # Ensure listener stops safely
        if timer_running:  # If the timer is still running, call `stop_timer` to save the data
            stop_timer()
        try:
            root.destroy()  # Properly close the window
        except TclError:
            pass  # Ignore if the window is already destroyed
        sys.exit()  # Exit the application

    root.protocol("WM_DELETE_WINDOW", on_close)

    root.mainloop()


def show_timer(username):
    global root, timer_running, start_time, last_activity_time, timer_label, break_mode, punch_in_time, punch_out_time

    root = CTk.CTk()
    root.title("Punch In Timer")
    root.geometry("250x250")
    root.resizable(False, False)
    
    timer_running = False
    start_time = 0
    last_activity_time = 0
    break_mode = False
    punch_in_time = ""
    punch_out_time = ""

    title_label = CTk.CTkLabel(root, text="Punch In Timer", font=("Helvetica", 16))
    title_label.pack(pady=10)

    user_label = CTk.CTkLabel(root, text=f"Hi, {user_session['name']}", font=("Helvetica", 12))
    user_label.pack(pady=5)

    timer_label = CTk.CTkLabel(root, text="00:00:00", font=("Helvetica", 30))
    timer_label.pack(pady=10)

    button_frame = CTk.CTkFrame(root)
    button_frame.pack(pady=10)

    punch_in_button = CTk.CTkButton(button_frame, text="Punch In ", command=start_timer, fg_color="#009900", width=10)
    punch_in_button.grid(row=0, column=0, padx=5)

    punch_out_button = CTk.CTkButton(button_frame, text="Punch Out", command=stop_timer, fg_color="#ed2939", width=10)
    punch_out_button.grid(row=0, column=1, padx=5)

    break_button = CTk.CTkButton(root, text="Add to Break", command=add_to_break, fg_color="blue", width=150)
    break_button.pack(pady=10)

    listener = mouse.Listener(on_move=on_mouse_move)
    listener.start()

    def on_close():
        global root
        if listener:
            listener.stop()  # Ensure listener stops safely
        if timer_running:  # If the timer is still running, call `stop_timer` to save the data
            stop_timer()
        try:
            root.destroy()  # Properly close the window
        except TclError:
            pass  # Ignore if the window is already destroyed
        sys.exit()  # Exit the application

        

    root.protocol("WM_DELETE_WINDOW", on_close)

    root.mainloop()
    
# Main application
# Function to handle application closing
def on_anilbran_att_ad_user_closing():
    if messagebox.askyesno("Exit", "Are you sure you want to exit the application?"):
        anilbran_att_ad_user.destroy()
        sys.exit()
# Main application window
anilbran_att_ad_user = CTk.CTk()
anilbran_att_ad_user.protocol("WM_DELETE_WINDOW", on_anilbran_att_ad_user_closing)
anilbran_att_ad_user.withdraw()  # Hide the main window initially
# Login window
login_window = CTk.CTkToplevel(anilbran_att_ad_user)
login_window.title("Login")
login_window.geometry("300x220")
login_window.resizable(False, False)
def on_login_closing():
    if messagebox.askyesno("Exit", "Are you sure you want to exit the application?"):
        
        anilbran_att_ad_user.destroy()
        sys.exit()
login_window.protocol("WM_DELETE_WINDOW", on_login_closing)
CTk.CTkLabel(login_window, text="Username:").place(x=40,y=20)
username_entry = CTk.CTkEntry(login_window)
username_entry.place(x=120,y=20)
CTk.CTkLabel(login_window, text="Password:").place(x=40,y=70)
password_entry = CTk.CTkEntry(login_window, show="*")
password_entry.place(x=120,y=70)
login_button = CTk.CTkButton(login_window, text="Login", command=validate_login)
login_button.place(x=80,y=130)
change_password_button = CTk.CTkButton(login_window, text="Change Password", command= change_password, fg_color="blue")
change_password_button.place(x=80,y=170)
anilbran_att_ad_user.mainloop()

#END OF THE CODE HERE !!!


