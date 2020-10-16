from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import shelve
import email_downloader
import threading




# Main email app(root)
class EmailApp(Tk):

    #################################################################################################################################

    def __init__(self):
        super().__init__()
        self.geometry("380x200")
        self.title("Email Downloader")


        # List box with the current downloads
        self.email_list = Listbox(self, height=7, width=50, border=0)
        self.email_list.grid(row=0, column=0, columnspan=3, rowspan=3, pady=20, padx=20, sticky=(N, E, S, W))
        self.email_list.bind('<<ListboxSelect>>', self.select_item)
        

        # buttons
        self.new_download_btn = ttk.Button(self, text='New', command=self.open_new_download)
        self.new_download_btn.grid(row=3, column=0)


        self.continue_btn = ttk.Button(self, text='Continue', command=self.continue_download)
        self.continue_btn.grid(row=3, column=1)

        self.pause_btn = ttk.Button(self, text='Exit', command=self.close_window)
        self.pause_btn.grid(row=3, column=2)

        # Populate the email list with unfinished emails
        self.populate_email_list()

    #################################################################################################################################


    def open_new_download(self):
        """Creates a new download instance"""
        new = NewDownload(self)
        new.grab_set()

    #################################################################################################################################


    def select_item(self, event):
        """Select an item in the list box"""
        index = self.email_list.curselection()[0]
        selected_item = self.email_list.get(index)
        self.continue_email = selected_item[0]

    #################################################################################################################################


    def continue_download(self):
        """Continue an email download"""
        try:
            shelfFile = shelve.open('save')
            email = shelfFile[self.continue_email]['email']
            password = shelfFile[self.continue_email]['password']
            uids = shelfFile[self.continue_email]['uids']
            last_download_index = shelfFile[self.continue_email]['last_downloaded']
            number_of_email = shelfFile[self.continue_email]['number_of_emails']
            shelfFile.close()

            # Connect to the server and resume the download
            conn = ConnectToServer(self)
            conn_obj = threading.Thread(target=conn.continue_download, kwargs={'email': email, 'password': password, 'number_of_email': number_of_email, 'uids':uids, 'last_download_index': last_download_index})
            conn_obj.start()

        except AttributeError:
            messagebox.showerror('No selection', "Please select a row in the list box!")
        
        except KeyError:
            messagebox.showerror('No selection', "Please select a row in the list box!")

    #################################################################################################################################


    def populate_email_list(self):
        """Populates the list of emails that have not yet finished downloading"""
        try:
            self.email_list.delete(0, END)
            shelfFile = shelve.open('save')
            for k, v in shelfFile.items():
                percentage_progress = (shelfFile[k]['last_downloaded'] / shelfFile[k]['number_of_emails']) * 100
                row = [shelfFile[k]['email'], str(shelfFile[k]['last_downloaded']) + ' out of '+  str(shelfFile[k]['number_of_emails']), str(percentage_progress)+'%']
                self.email_list.insert(END, row)
            shelfFile.close()
        except KeyError:
            shelfFile.close()

    #################################################################################################################################

    def close_window(self):
        """Closes main window, exits the main application"""
        self.destroy()





# Responsible for new download instances
class NewDownload(Toplevel):

    #################################################################################################################################

    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.geometry("330x200")
        self.title("Download")

        # Email download frame
        download_frame = ttk.Labelframe(self, text="Download")
        download_frame.pack(fill='both', expand='yes', padx=20, pady=10)

        # email entry
        self.email = StringVar()
        self.email_lbl = ttk.Label(download_frame, text='Email')
        self.email_entry = ttk.Entry(download_frame, textvariable=self.email)
        self.email_lbl.grid(row=0, column=0, padx=5, pady=3,)
        self.email_entry.grid(row=0, column=1, padx=5, pady=3,)

        # Password Entry
        self.password = StringVar()
        self.password_lbl = ttk.Label(download_frame, text="Password")
        self.password_entry = ttk.Entry(download_frame, textvariable=self.password, show="*")
        self.password_lbl.grid(row=1, column=0, padx=5, pady=3,)
        self.password_entry.grid(row=1, column=1, padx=5, pady=3,)

        # number entry
        self.number = StringVar()
        self.number.set('All')
        self.number_lbl = ttk.Label(download_frame, text="Number of emails")
        self.number_entry = ttk.Entry(download_frame, textvariable=self.number)
        self.number_lbl.grid(row=2, column=0, padx=5, pady=3,)
        self.number_entry.grid(row=2, column=1, padx=5, pady=3,)


        # Download button
        download_btn = ttk.Button(download_frame, text='Download', command=self.connect_to_server)
        download_btn.grid(row=3, column=0, padx=5, pady=3, sticky=(E,W))

        # Cancel button
        cancel_button = ttk.Button(download_frame, text="Cancel", command=self.destroy)
        cancel_button.grid(row=3, column=1, padx=5, pady=3, sticky=(E,W))

    #################################################################################################################################

    def connect_to_server(self):
        """Connect to the imap server as well as log in to gmail"""

        if self.email.get() != '' and self.password.get() != '' and self.number.get() != '':

            # get the parameters
            email = self.email.get()
            password = self.password.get()
            number_of_email = self.number.get()

            # Creating a connection loading icon
            conn = ConnectToServer(self.parent)
        
            conn_obj = threading.Thread(target=conn.connect_to_the_server, kwargs={'email': email, 'password': password, 'number_of_email': number_of_email})
            conn_obj.start()
            self.destroy() # Close the download window and open the connecting progress loader as a thread

        
        else:
            messagebox.showerror("Empty Fields", "Please completely fill in the form.", parent=self)

    #################################################################################################################################
    
        

    



class ConnectToServer(Toplevel):

    #################################################################################################################################

    def __init__(self, parent):
        super().__init__(parent)
        self.geometry("220x70")
        self.title("Connecting")
        self.parent = parent

        # Connection status label
        self.status = StringVar()
        status_lbl = ttk.Label(self, text='Connecting...', textvariable=self.status)
        status_lbl.grid(row=0, column=0, padx=5, pady=5, sticky=W)

        # Progress bar for the connection status
        self.status_bar = ttk.Progressbar(self, orient=HORIZONTAL, length=200, value=0)
        self.status_bar.grid(row=1, column=0, padx=5, pady=5)

    #################################################################################################################################

    def connect_to_the_server(self, email, password, number_of_email):
        """Connect to the server and login"""
        conn = email_downloader.ContinueDownload(email, password, number_of_email).login_to_imap_server(self.parent, self, self.status, self.status_bar)

    #################################################################################################################################


    def continue_download(self, email, password, number_of_email, uids, last_download_index):
        """Continue a download, connect to the server, login and download"""
        conn = email_downloader.ContinueDownload(email, password, number_of_email, uids, last_download_index).login_to_imap_server(self.parent, self, self.status, self.status_bar)
        
    #################################################################################################################################
    




if __name__ == '__main__':
    app = EmailApp()
    app.mainloop()