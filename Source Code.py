import csv
import tkinter as tk
from tkinter import messagebox, ttk
from PIL import Image, ImageTk
import time
import winsound


class StartScreen:
    def __init__(self):
        self.start_screen = tk.Tk()
        self.start_screen.geometry("600x340")
        self.start_screen.title("Address Book Program")
        self.start_screen.configure(bg="#E4CCB4")

        image = Image.open("bgs.png")
        background_image = ImageTk.PhotoImage(image)
        background_label = tk.Label(self.start_screen, image=background_image)
        background_label.place(x=0, y=0, relwidth=1, relheight=1)
        background_label.image = background_image

        tk.Label(self.start_screen, text="A Final Project for Object-Oriented Programming",
                 font=("Gabriola", 13, "bold"), bg="#E4CCB4").pack(pady=10)
        tk.Label(self.start_screen, text="Address Book Program", font=("Monotype Corsiva", 25, "bold"), bg="#E4CCB4",
                 fg="#532C1F").pack(pady=0)
        tk.Button(self.start_screen, text="START", font=("Montserrat", 15, "bold",), width=8,
                  command=self.show_loading_screen, bg="#794F2E", bd=0, fg="white").pack(pady=15)

        tk.Label(self.start_screen, text="PROGRAMMED BY:\nAlmario, Candido James\nBenciang, Gienel Aubrey\n"
                                         "Binasa, Kenneth Carl\nDe Leon, Novelle",
                 font=("Montserrat", 11), bg="#E4CCB4", justify=tk.CENTER).pack(pady=20)

        self.start_screen.mainloop()

    def show_loading_screen(self):
        self.start_screen.withdraw()

        loading_screen = tk.Toplevel()
        loading_screen.geometry("300x100")
        loading_screen.title("Loading...")
        loading_screen.config(bg="#E8D9CD")

        style = ttk.Style()
        style.theme_use('alt')
        style.configure("Custom.Horizontal.TProgressbar", background='#794F2E')

        progressbar = ttk.Progressbar(loading_screen, mode='determinate', length=250, style="Custom.Horizontal.TProgressbar")
        progressbar.pack(pady=40)

        progressbar.start()
        winsound.PlaySound("StartingSFX", winsound.SND_FILENAME)
        for i in range(101):
            progressbar['value'] = i
            loading_screen.update()  # Update the loading screen
            # Simulate a delay
            time.sleep(0.008)
        progressbar.stop()
        loading_screen.destroy()
        self.show_main_gui()
        loading_screen.after(99)
        loading_screen.mainloop()

    def show_main_gui(self):
        self.start_screen.deiconify()
        self.start_screen.destroy()
        x = AddressBookGUI()


class AddressBookGUI:
    def __init__(self):
        try:
            self.Address_Book_GUI = tk.Tk()

            self.Address_Book_GUI.geometry('600x340')
            self.Address_Book_GUI.config(bg="#E8D9CD")
            self.Address_Book_GUI.resizable(0, 0)
            self.Address_Book_GUI.title("Address Book - Final Project")

            image = Image.open("bgs.png")
            background_image = ImageTk.PhotoImage(image)
            background_label = tk.Label(self.Address_Book_GUI, image=background_image)
            background_label.place(x=0, y=0, relwidth=1, relheight=1)
            background_label.image = background_image

            tk.Label(self.Address_Book_GUI, text="ADDRESS BOOK", font=("Gabriola", 24, "bold"), fg="#794F2E",
                     bg="#E4CCB4").pack(padx=3, pady=2.5)

            self.sortUser = [
                "First Name",
                "Last Name",
                "Address",
                "Contact"
            ]

            self.Address_Book_Format()
            self.readCSV_USER()
            self.Address_Book_Titles()
            self.userSearchDropdownFormat()
            self.userSearchBoxFormat()
            self.Address_Book_List()
            self.Display_User_Box()
            self.Address_bookButtons()

            self.Address_Book_GUI.mainloop()
        except FileNotFoundError:
            messagebox.showerror("Error", "Required file 'AddressBookData.csv' not found.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def Address_Book_Format(self):
        try:
            self.Address_Book_GUI.geometry('600x340')
            self.Address_Book_GUI.config(bg="#E8D9CD")
            self.Address_Book_GUI.resizable(0, 0)
            self.Address_Book_GUI.title("Address Book - Final Project")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def readCSV_USER(self):
        try:
            global header
            self.Address_Book_Main = []
            with open('AddressBookData.csv', newline='') as csvfile:
                csv_reader = csv.reader(csvfile, delimiter=',')
                header = next(csv_reader)
                for row in csv_reader:
                    self.Address_Book_Main.append(row)
        except FileNotFoundError:
            messagebox.showerror("Error", "Required file 'AddressBookData.csv' not found.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def writeCSV_USER(self, addList):
        try:
            with open('AddressBookData.csv', 'w', newline='') as csv_file:
                writeobj = csv.writer(csv_file, delimiter=',')
                writeobj.writerow(header)
                for row in addList:
                    writeobj.writerow(row)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def Address_Book_Titles(self):
        try:
            tk.Label(self.Address_Book_GUI, text="First Name:", font="Gabriola 16 bold", bg="#E4CCB4", fg="#794F2E").place(
                x=95, y=70)
            self.U_FirstName = tk.StringVar()
            tk.Entry(self.Address_Book_GUI, textvariable=self.U_FirstName, font="Gabriola 12 bold", bg="#D1A27D",
                     fg="white", width=25).place(x=190, y=80)

            tk.Label(self.Address_Book_GUI, text="Last Name:", font="Gabriola 16 bold", bg="#E4CCB4", fg="#794F2E").place(
                x=97, y=100)
            self.U_LastName = tk.StringVar()
            tk.Entry(self.Address_Book_GUI, textvariable=self.U_LastName, font="Gabriola 12 bold", bg="#D1A27D",
                     fg="white", width=25, selectborderwidth=10).place(x=190, y=110)

            tk.Label(self.Address_Book_GUI, text="Address:", font="Gabriola 16 bold", bg="#E4CCB4", fg="#794F2E").place(
                x=115, y=130)
            self.U_Address = tk.StringVar()
            tk.Entry(self.Address_Book_GUI, textvariable=self.U_Address, font="Gabriola 12 bold", bg="#D1A27D",
                     fg="white", width=25, selectborderwidth=10).place(x=190, y=140)

            tk.Label(self.Address_Book_GUI, text="Contact Number:", font="Gabriola 16 bold", bg="#E4CCB4",
                     fg="#794F2E").place(x=54, y=160)
            self.U_ContactNumber = tk.StringVar()
            tk.Entry(self.Address_Book_GUI, textvariable=self.U_ContactNumber, font="Gabriola 12 bold", bg="#D1A27D",
                     fg="white", width=25, selectborderwidth=10).place(x=190, y=170)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def userSearchDropdownFormat(self):
        try:
            self.searchUser = ttk.Combobox(self.Address_Book_GUI, value=self.sortUser, height=10, width=10,
                                           font=("Montserrat", 8))
            self.searchUser.current(0)
            self.searchUser.place(x=500, y=50)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def userSearchBoxFormat(self):
        try:
            self.userEntrySearch = tk.Entry(self.Address_Book_GUI, font="Montserrat 8", width=16)
            self.userEntrySearch.place(x=380, y=50)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def Address_Book_List(self):
        try:
            self.main_Store_USER = tk.Frame(self.Address_Book_GUI)
            self.main_Store_USER.place(x=380, y=80)
            scroll_Store = tk.Scrollbar(self.main_Store_USER, orient=tk.VERTICAL)
            self.main_box_Store_USER = tk.Listbox(self.main_Store_USER, yscrollcommand=scroll_Store.set, height=6,
                                               width=26, font="Gabriola 12 bold", bg="#794F2E", selectborderwidth=5, bd=0,
                                               foreground="white")
            scroll_Store.config(command=self.main_box_Store_USER.yview)
            scroll_Store.pack(side=tk.RIGHT, fill=tk.Y)
            self.main_box_Store_USER.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def main_Index_User(self):
        try:
            if len(self.main_box_Store_USER.curselection()) == 0:
                messagebox.showerror("Error", "Please Select the User first")
            else:
                return int(self.main_box_Store_USER.curselection()[0])
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def User_add_contact(self):
        try:
            if len(self.U_ContactNumber.get()) == 11:
                if self.validate_name(self.U_FirstName.get()) and self.validate_name(self.U_LastName.get()) \
                        and self.U_Address.get() and self.U_ContactNumber.get():
                    self.Address_Book_Main.append([self.U_FirstName.get(), self.U_LastName.get(), self.U_Address.get(),
                                                   self.U_ContactNumber.get()])
                    messagebox.showinfo("Successful", "The Contact has been Added")
                    self.writeCSV_USER(self.Address_Book_Main)
                    self.user_Refresh()
                else:
                    messagebox.showerror("Error", "Invalid Entry \nPlease provide valid names and complete all the Entry \nType (N/A) if none")
            else:
                messagebox.showerror("Error", "Contact number must be exactly 11 digits long.")
            self.Display_User_Box()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def User_edit_contact(self):
        try:
            selected_name = self.main_box_Store_USER.get(tk.ACTIVE)
            selected_index = next((i for i, data in enumerate(self.Address_Book_Main) if selected_name in data), None)
            if selected_index is not None:
                if self.validate_name(self.U_FirstName.get()) and self.validate_name(self.U_LastName.get()) \
                        and self.U_Address.get() and len(self.U_ContactNumber.get()) == 11:
                    self.Address_Book_Main[selected_index] = [
                        self.U_FirstName.get(), self.U_LastName.get(), self.U_Address.get(), self.U_ContactNumber.get()
                    ]
                    messagebox.showinfo("Successful", "The Contact has been Edited")
                    self.writeCSV_USER(self.Address_Book_Main)
                    self.user_Refresh()
                else:
                    messagebox.showerror("Error", "Please provide valid names, complete all the Entry, "
                                                 "and ensure that the contact number is exactly 11 digits long.")
                self.Display_User_Box()
                self.main_box_Store_USER.selection_clear(0, tk.END)
                self.main_box_Store_USER.selection_set(selected_index)
            else:
                messagebox.showerror("Error", "Please select a contact to edit.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def validate_name(self, name):
        return not any(char.isdigit() for char in name)

    def User_delete_contact(self):
        try:
            del self.Address_Book_Main[self.main_Index_User()]
            self.writeCSV_USER(self.Address_Book_Main)
            self.Display_User_Box()
            messagebox.showinfo("Successful", "The Contact has been Deleted")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def User_view_contact(self):
        try:
            selected_index = self.main_box_Store_USER.curselection()
            if not selected_index:
                messagebox.showerror("Error", "Please select a contact to view.")
            else:
                selected_contact = self.main_box_Store_USER.get(selected_index)
                searchSorter = self.searchUser.get()
                for data in self.Address_Book_Main:
                    if searchSorter == "First Name" and selected_contact == data[0]:
                        User_First_Name, User_Last_Name, User_Address, User_ContactNumber = data
                        self.U_FirstName.set(User_First_Name)
                        self.U_LastName.set(User_Last_Name)
                        self.U_Address.set(User_Address)
                        self.U_ContactNumber.set(User_ContactNumber)
                    elif searchSorter == "Last Name" and selected_contact == data[1]:
                        User_First_Name, User_Last_Name, User_Address, User_ContactNumber = data
                        self.U_FirstName.set(User_First_Name)
                        self.U_LastName.set(User_Last_Name)
                        self.U_Address.set(User_Address)
                        self.U_ContactNumber.set(User_ContactNumber)
                    elif searchSorter == "Address" and selected_contact == data[2]:
                        User_First_Name, User_Last_Name, User_Address, User_ContactNumber = data
                        self.U_FirstName.set(User_First_Name)
                        self.U_LastName.set(User_Last_Name)
                        self.U_Address.set(User_Address)
                        self.U_ContactNumber.set(User_ContactNumber)
                    elif searchSorter == "Contact" and selected_contact == data[3]:
                        User_First_Name, User_Last_Name, User_Address, User_ContactNumber = data
                        self.U_FirstName.set(User_First_Name)
                        self.U_LastName.set(User_Last_Name)
                        self.U_Address.set(User_Address)
                        self.U_ContactNumber.set(User_ContactNumber)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def user_Refresh(self):
        try:
            self.U_FirstName.set("")
            self.U_LastName.set("")
            self.U_Address.set("")
            self.U_ContactNumber.set("")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def User_exit_contact(self):
        try:
            self.Address_Book_GUI.destroy()
            winsound.PlaySound("EndSFX", winsound.SND_FILENAME)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def Display_User_Box(self):
        try:
            self.Address_Book_Main.sort()
            self.main_box_Store_USER.delete(0, tk.END)
            searchSorter = self.searchUser.get()
            for firstname, lastname, address, contactnumber in self.Address_Book_Main:
                if searchSorter == "First Name":
                    self.main_box_Store_USER.insert(tk.END, firstname)
                elif searchSorter == "Last Name":
                    self.main_box_Store_USER.insert(tk.END, lastname)
                elif searchSorter == "Address":
                        self.main_box_Store_USER.insert(tk.END, address)
                elif searchSorter == "Contact":
                    self.main_box_Store_USER.insert(tk.END, contactnumber)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def User_search_contact(self):
        try:
            def searchButtonPressed():
                userInput = self.userEntrySearch.get()
                searchSorter = self.searchUser.get()
                self.main_box_Store_USER.delete(0, tk.END)
                if not userInput:
                    # Display all contacts if the search box is empty
                    sorted_contacts = sorted(self.Address_Book_Main, key=lambda x: x[self.sortUser.index(searchSorter)])
                    for data in sorted_contacts:
                        self.main_box_Store_USER.insert(tk.END, data[self.sortUser.index(searchSorter)])
                elif searchSorter:
                    # Perform the search based on the search type
                    for data in self.Address_Book_Main:
                        if searchSorter == "First Name" and userInput.lower() in data[0].lower():
                            self.main_box_Store_USER.insert(tk.END, data[0])
                        elif searchSorter == "Last Name" and userInput.lower() in data[1].lower():
                            self.main_box_Store_USER.insert(tk.END, data[1])
                        elif searchSorter == "Address" and userInput.lower() in data[2].lower():
                            self.main_box_Store_USER.insert(tk.END, data[2])
                        elif searchSorter == "Contact" and userInput.lower() in data[3].lower():
                            self.main_box_Store_USER.insert(tk.END, data[3])
                else:
                    messagebox.showerror("Error", "Please enter a search query and select a search type.")

            tk.Button(self.Address_Book_GUI, text="SEARCH", font='Gabriola 12 bold', bg='#794F2E', foreground="white",
                      command=searchButtonPressed, bd=0, padx=10, pady=0).place(x=290, y=221)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def Address_bookButtons(self):
        try:
            tk.Button(self.Address_Book_GUI, text="ADD", font='Gabriola 12 bold', bg='#794F2E', foreground="white",
                   command=self.User_add_contact, bd=0, padx=10, pady=0).place(x=25, y=221)
            tk.Button(self.Address_Book_GUI, text="EDIT", font='Gabriola 12 bold', bg='#794F2E', foreground="white",
                   command=self.User_edit_contact, bd=0, padx=10, pady=0).place(x=85, y=221)
            tk.Button(self.Address_Book_GUI, text="DELETE", font='Gabriola 12 bold', bg='#794F2E', foreground="white",
                   command=self.User_delete_contact, bd=0, padx=10, pady=0).place(x=147, y=221)
            tk.Button(self.Address_Book_GUI, text="VIEW", font='Gabriola 12 bold', bg='#794F2E', foreground="white",
                   command=self.User_view_contact, bd=0, padx=10, pady=0).place(x=226, y=221)
            tk.Button(self.Address_Book_GUI, text="SEARCH", font='Gabriola 12 bold', bg='#794F2E', foreground="white",
                   command=self.User_search_contact, bd=0, padx=10, pady=0).place(x=290, y=221)
            tk.Button(self.Address_Book_GUI, text="CLEAR", font='Gabriola 12 bold', bg='#794F2E', foreground="white",
                   command=self.user_Refresh, bd=0, padx=10, pady=0).place(x=236, y=275)
            tk.Button(self.Address_Book_GUI, text="EXIT", font='Gabriola 12 bold', bg='#EFEFE9', foreground="black",
                   command=self.User_exit_contact, bd=0, padx=10, pady=0).place(x=310, y=275)
        except Exception as e:
            messagebox.showerror("Error", str(e))


try:
    proj = StartScreen()
except Exception as e:
    messagebox.showerror("Error", str(e))
