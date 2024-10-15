import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import scrolledtext
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import sys
import threading
from datetime import datetime, timedelta
import os

class HochschulsportBookingApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Hochschulsport Booking")
        self.master.geometry("800x800")
        self.master.resizable(False, False)

        self.driver = None
        self.create_widgets()
        self.update_clock()

    def create_widgets(self):
        main_frame = ttk.Frame(self.master, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        disclaimer = ("By pressing the Submit button, you acknowledge and agree to the universities participation terms and conditions. You accept full responsibility for running this code, and you understand that no personal data will be transmitted or stored anywhere other than as specified in the application.")
        disclaimer_label = ttk.Label(main_frame, text=disclaimer, wraplength=760, foreground="red")
        disclaimer_label.pack(pady=(0, 10))

        self.time_var = tk.StringVar()
        clock_label = ttk.Label(main_frame, textvariable=self.time_var, font=("Helvetica", 12))
        clock_label.pack(anchor='ne', pady=(0, 10))

        os_frame = ttk.Frame(main_frame)
        os_frame.pack(fill=tk.X, pady=5)

        os_label = ttk.Label(os_frame, text="Select Operating System:")
        os_label.pack(side=tk.LEFT, padx=(0, 10))

        self.os_var = tk.StringVar()
        os_dropdown = ttk.Combobox(os_frame, textvariable=self.os_var, state='readonly')
        os_dropdown['values'] = ("Windows", "mac_arm64", "mac_x64")
        os_dropdown.current(0)
        os_dropdown.pack(side=tk.LEFT)

        self.add_divider(main_frame)

        options_frame = ttk.Frame(main_frame)
        options_frame.pack(fill=tk.X, pady=10)

        self.form_option = tk.StringVar(value="option1")
        option1_radiobutton = ttk.Radiobutton(options_frame, text="Login with Email and Password",
                                             variable=self.form_option, value="option1",
                                             command=self.toggle_form_options)
        option1_radiobutton.pack(anchor='w', pady=2)

        option2_radiobutton = ttk.Radiobutton(options_frame, text="Full Registration Form",
                                             variable=self.form_option, value="option2",
                                             command=self.toggle_form_options)
        option2_radiobutton.pack(anchor='w', pady=2)

        self.option1_frame = ttk.Frame(main_frame)
        self.option1_frame.pack(fill=tk.X, pady=5)

        email_label = ttk.Label(self.option1_frame, text="Email:")
        email_label.grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.email_var = tk.StringVar()
        email_entry = ttk.Entry(self.option1_frame, textvariable=self.email_var, width=30)
        email_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w')

        password_label = ttk.Label(self.option1_frame, text="Password:")
        password_label.grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.password_var = tk.StringVar()
        password_entry = ttk.Entry(self.option1_frame, textvariable=self.password_var, show='*', width=30)
        password_entry.grid(row=1, column=1, padx=5, pady=5, sticky='w')

        self.option2_frame = ttk.Frame(main_frame)
        self.option2_frame.pack_forget()

        gender_label = ttk.Label(self.option2_frame, text="Gender:")
        gender_label.grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.gender_var = tk.StringVar()
        gender_dropdown = ttk.Combobox(self.option2_frame, textvariable=self.gender_var, state='readonly')
        gender_dropdown['values'] = ("Male", "Female", "Diverse", "No Info")
        gender_dropdown.current(0)
        gender_dropdown.grid(row=0, column=1, padx=5, pady=5, sticky='w')

        fname_label = ttk.Label(self.option2_frame, text="First Name:")
        fname_label.grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.fname_var = tk.StringVar()
        fname_entry = ttk.Entry(self.option2_frame, textvariable=self.fname_var, width=30)
        fname_entry.grid(row=1, column=1, padx=5, pady=5, sticky='w')

        lname_label = ttk.Label(self.option2_frame, text="Last Name:")
        lname_label.grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.lname_var = tk.StringVar()
        lname_entry = ttk.Entry(self.option2_frame, textvariable=self.lname_var, width=30)
        lname_entry.grid(row=2, column=1, padx=5, pady=5, sticky='w')

        street_label = ttk.Label(self.option2_frame, text="Street & No.:")
        street_label.grid(row=3, column=0, padx=5, pady=5, sticky='e')
        self.street_var = tk.StringVar()
        street_entry = ttk.Entry(self.option2_frame, textvariable=self.street_var, width=30)
        street_entry.grid(row=3, column=1, padx=5, pady=5, sticky='w')

        zipcode_label = ttk.Label(self.option2_frame, text="Zipcode & City:")
        zipcode_label.grid(row=4, column=0, padx=5, pady=5, sticky='e')
        self.zipcode_var = tk.StringVar()
        zipcode_entry = ttk.Entry(self.option2_frame, textvariable=self.zipcode_var, width=10)
        zipcode_entry.grid(row=4, column=1, padx=(5, 0), pady=5, sticky='w')
        self.city_var = tk.StringVar()
        city_entry = ttk.Entry(self.option2_frame, textvariable=self.city_var, width=20)
        city_entry.grid(row=4, column=1, padx=(100, 5), pady=5, sticky='w')

        matric_label = ttk.Label(self.option2_frame, text="Matriculation No.:")
        matric_label.grid(row=5, column=0, padx=5, pady=5, sticky='e')
        self.matric_var = tk.StringVar()
        matric_entry = ttk.Entry(self.option2_frame, textvariable=self.matric_var, width=30)
        matric_entry.grid(row=5, column=1, padx=5, pady=5, sticky='w')

        email2_label = ttk.Label(self.option2_frame, text="Email:")
        email2_label.grid(row=6, column=0, padx=5, pady=5, sticky='e')
        self.email2_var = tk.StringVar()
        email2_entry = ttk.Entry(self.option2_frame, textvariable=self.email2_var, width=30)
        email2_entry.grid(row=6, column=1, padx=5, pady=5, sticky='w')

        tel_label = ttk.Label(self.option2_frame, text="Telephone No. (Optional):")
        tel_label.grid(row=7, column=0, padx=5, pady=5, sticky='e')
        self.tel_var = tk.StringVar()
        tel_entry = ttk.Entry(self.option2_frame, textvariable=self.tel_var, width=30)
        tel_entry.grid(row=7, column=1, padx=5, pady=5, sticky='w')

        self.bottom_frame = ttk.Frame(main_frame)
        self.bottom_frame.pack(fill=tk.X, pady=10)

        self.add_divider(self.bottom_frame)

        request_frame = ttk.Frame(self.bottom_frame)
        request_frame.pack(fill=tk.X, pady=10)

        url_label = ttk.Label(request_frame, text="Hochschulsport URL:")
        url_label.grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.url_var = tk.StringVar(value="https://buchung.hsp.uni-tuebingen.de/angebote/aktueller_zeitraum/_Passwort_erstellen.html")
        url_entry = ttk.Entry(request_frame, textvariable=self.url_var, width=50)
        url_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w')

        # add course number field
        course_number_label = ttk.Label(request_frame, text="Course Number:")
        course_number_label.grid(row=0, column=2, padx=5, pady=5, sticky='e')
        self.course_number_var = tk.StringVar()
        course_number_entry = ttk.Entry(request_frame, textvariable=self.course_number_var, width=10)
        course_number_entry.grid(row=0, column=3, padx=5, pady=5, sticky='w')

        start_time_label = ttk.Label(request_frame, text="Start Time (HH:MM:SS):")
        start_time_label.grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.start_time_var = tk.StringVar()
        start_time_entry = ttk.Entry(request_frame, textvariable=self.start_time_var, width=10)
        start_time_entry.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        default_time = (datetime.now() + timedelta(minutes=5)).strftime("%H:%M:%S")
        start_time_entry.insert(0, default_time)

        submit_button = ttk.Button(self.bottom_frame, text="Submit", command=self.submit_form)
        submit_button.pack(pady=20)

    def add_divider(self, parent):
        divider = ttk.Separator(parent, orient='horizontal')
        divider.pack(fill=tk.X, pady=10)

    def toggle_form_options(self):
        if self.form_option.get() == "option1":
            self.option1_frame.pack(fill=tk.X, pady=5)
            self.option2_frame.pack_forget()
        else:
            self.option2_frame.pack(fill=tk.X, pady=5)
            self.option1_frame.pack_forget()
        self.bottom_frame.pack_forget()
        self.bottom_frame.pack(fill=tk.X, pady=10)

    def update_clock(self):
        now = time.strftime("%H:%M:%S")
        self.time_var.set(f"Current Time: {now}")
        self.master.after(1000, self.update_clock)

    def validate_zipcode(self, zipcode):
        return zipcode.isdigit() and len(zipcode) == 5

    def submit_form(self):
        os_selected = self.os_var.get()
        form_option = self.form_option.get()
        
        if form_option == "option1":
            email = self.email_var.get().strip()
            password = self.password_var.get().strip()
            if not email or not password:
                messagebox.showwarning("Input Required", "Please enter both Email and Password.")
                return
            user_data = {
                "option": "email_password",
                "email": email,
                "password": password
            }
        else:
            gender = self.gender_var.get()
            first_name = self.fname_var.get().strip()
            last_name = self.lname_var.get().strip()
            street = self.street_var.get().strip()
            zipcode = self.zipcode_var.get().strip()
            city = self.city_var.get().strip()
            matric_no = self.matric_var.get().strip()
            email = self.email2_var.get().strip()
            telephone = self.tel_var.get().strip()

            if not (first_name and last_name and street and zipcode and city and matric_no and email):
                messagebox.showwarning("Input Required", "Please fill in all mandatory fields.")
                return

            if not self.validate_zipcode(zipcode):
                messagebox.showwarning("Invalid Input", "Zipcode must be exactly 5 digits.")
                return

            user_data = {
                "option": "full_form",
                "gender": gender,
                "first_name": first_name,
                "last_name": last_name,
                "street": street,
                "zipcode": zipcode,
                "city": city,
                "matric_no": matric_no,
                "email": email,
                "telephone": telephone
            }

        url = self.url_var.get().strip()
        start_time_str = self.start_time_var.get().strip()
        course_number = self.course_number_var.get().strip()

        if not url:
            messagebox.showwarning("Input Required", "Please enter the Hochschulsport URL.")
            return

        try:
            start_time = datetime.strptime(start_time_str, "%H:%M:%S").time()
        except ValueError:
            messagebox.showwarning("Invalid Time", "Please enter start time in HH:MM:SS format.")
            return

        confirmation = messagebox.askyesno("Confirm Submission",
                                           "By submitting, you agree to the participation terms and data usage policies.\nDo you want to proceed?")
        if not confirmation:
            return

        collected_data = {
            "os": os_selected,
            "user_data": user_data,
            "url": url,
            "start_time": start_time_str,
            "course_number": course_number
        }

        # Start the booking process
        threading.Thread(target=self.start_booking_process, args=(collected_data,), daemon=True).start()

    def start_booking_process(self, collected_data):
        self.setup_webdriver(collected_data['os'])
        self.wait_until_booking_time(collected_data['start_time'])
        self.refresh_booking_page(collected_data['course_number'])
        self.perform_booking(collected_data)

    def setup_webdriver(self, os_type):
        options = Options()
        # actually show whats happening
        # options.add_argument('--headless')
        current_dir = os.path.dirname(os.path.realpath(__file__))
        
        if os_type == 'mac_x64':
            cService = webdriver.ChromeService(executable_path=os.path.join(current_dir, 'chromedriver_macx64'))
        if os_type == 'mac_arm64':
            cService = webdriver.ChromeService(executable_path=os.path.join(current_dir, 'chromedriver_macarm64'))
        else:  # macx64
            cService = webdriver.ChromeService(executable_path=os.path.join(current_dir, 'chromedriver_win32.exe'))
        self.driver = webdriver.Chrome(service=cService, options=options)
        self.driver.maximize_window()

    def wait_until_booking_time(self, start_time_str):
        start_time = datetime.strptime(start_time_str, "%H:%M:%S")
        start_time = datetime.now().replace(hour=start_time.hour, minute=start_time.minute, second=start_time.second)
        while True:
            now = datetime.now()
            if now >= start_time - timedelta(seconds=30):
                self.driver.get(self.url_var.get())
                if now >= start_time:
                    break
            time.sleep(1)


    def refresh_booking_page(self, course_number=''):
        max_attempts = 300  # 5 minutes of attempts
        attempt = 0

        self.handle_cookie_popup()

        while attempt < max_attempts:
            try:
                if course_number:
                    # Look for the specific course number
                    xpath = f"//td[@class='bs_sbuch']/a[@id='K{course_number}']/following-sibling::input[@type='submit']"
                else:
                    # Look for any submit button in the booking column
                    xpath = "//td[@class='bs_sbuch']//input[@type='submit']"
                
                submit_button = self.driver.find_element(By.XPATH, xpath)
                
                if submit_button.is_displayed() and submit_button.is_enabled():
                    print(f"Submit button found after {attempt} attempts.")
                    return submit_button
            except Exception as e:
                print(f"Attempt {attempt + 1}: Submit button not found. Refreshing...")
            
            time.sleep(1)  # Wait for 1 second before refreshing
            self.driver.refresh()
            attempt += 1
        
        raise Exception("Booking button did not appear within the allocated time.")

    def perform_booking(self, collected_data):
        try:
            self.click_booking_button()
            
            if collected_data['user_data']['option'] == 'email_password':
                self.fill_form_option_1(collected_data['user_data'])
            else:
                self.fill_form_option_2(collected_data['user_data'])
            
            self.submit_booking()
            messagebox.showinfo("Booking Complete", "The booking process has been completed successfully!")
        except Exception as e:
            messagebox.showerror("Booking Error", f"An error occurred during the booking process: {str(e)}")
        finally:
            pass

    def handle_cookie_popup(self):
        try:
            # Wait for the cookie popup to appear
            cookie_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@data-in2-modal-accept-button and contains(text(), 'Alle akzeptieren')]"))
            )
            cookie_button.click()
            time.sleep(1)
        except Exception as e:
            print(f"No cookie popup found or error occurred: {str(e)}")

    def click_booking_button(self):
        booking_button = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "bs_btn_buchen"))
        )
        booking_button.click()

    def fill_form_option_1(self, user_data):
        self.driver.switch_to.window(self.driver.window_handles[-1])
        
        pointer_div = self.driver.find_element(By.CLASS_NAME, "pointer")
        pointer_div.click()

        email_input = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.NAME, "pw_email"))
        )
        email_input.send_keys(user_data['email'])
        self.driver.find_element(By.XPATH, '//input[@type="password"]').send_keys(user_data['password'])

        self.driver.find_element(By.XPATH, '//input[@type="submit"]').click()

    def fill_form_option_2(self, user_data):
        self.driver.switch_to.window(self.driver.window_handles[-1])

        gender_map = {'Male': 'M', 'Female': 'W', 'Diverse': 'D', 'No Info': 'X'}
        gender_input = self.driver.find_element(By.CSS_SELECTOR, f"input[name='sex'][value='{gender_map[user_data['gender']]}']")
        gender_input.click()

        self.driver.find_element(By.NAME, "vorname").send_keys(user_data['first_name'])
        self.driver.find_element(By.NAME, "name").send_keys(user_data['last_name'])
        self.driver.find_element(By.NAME, "strasse").send_keys(user_data['street'])
        self.driver.find_element(By.NAME, "ort").send_keys(f"{user_data['zipcode']} {user_data['city']}")
        
        status_dropdown = self.driver.find_element(By.ID, "BS_F1600")
        status_dropdown.send_keys("S-UNIT")
        time.sleep(1)
        self.driver.find_element(By.NAME, "matnr").send_keys(user_data['matric_no'])

        self.driver.find_element(By.NAME, "email").send_keys(user_data['email'])
        
        if user_data['telephone']:
            self.driver.find_element(By.NAME, "telefon").send_keys(user_data['telephone'])

    def submit_booking(self):
        self.driver.find_element(By.NAME, "tnbed").click()
        self.driver.find_element(By.CSS_SELECTOR, 'input.sub[type="submit"]').click()

        WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'input.sub[type="submit"][value="verbindlich buchen"]'))
        ).click()

    def on_closing(self):
        if self.driver:
            self.driver.quit()
        self.master.destroy()

def main():
    root = tk.Tk()
    app = HochschulsportBookingApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()