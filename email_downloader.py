#Email Downloader
#This program downloads email and saves the contents in a csv file or json


########################################################################
"""HOW TO RUN THE PROGRAM"""
"""
1.Activate the environment with imapclient, pyzmail and console_progressbar installed
"""
########################################################################



import imapclient
import pyzmail
import csv
import sys
import time
import json
import shelve
import errno

from tkinter import *
from tkinter import ttk
from tkinter import messagebox



#Get email, password and the number of emails


class Download(Toplevel):
    
    #################################################################################################################################
    
    def __init__(self, parent):
        super().__init__(parent)
        self.geometry("220x100")
        self.title("Downloading")

        # Progress Label
        self.progress = StringVar()
        self.progress_lbl = ttk.Label(self, text='Getting ready')
        self.progress_lbl.grid(row=0, column=0, padx=5, pady=5, sticky=W)

        # Counter for number of emails
        self.counter = 0

        # Download Progress bar
        self.progress_bar = ttk.Progressbar(self, orient=HORIZONTAL, length=200, value=1)
        self.progress_bar.grid(row=1, column=0, padx=5, pady=5, columnspan=2)


        # Pause and cancel downloads
        self.pause = ttk.Button(self, text='Pause', command=self.pause)
        self.pause.grid(row=2, column=0, padx=5, pady=5)

        self.cancel = ttk.Button(self, text="Cancel", command=self.cancel)
        self.cancel.grid(row=2, column=1, padx=5, pady=5)


        self.pausing = False
        self.canceling = False
        self.complete = False

    #################################################################################################################################


    def configure_progressbar(self, maximum):
        """Configure the progress bar, we supply the maximum value to it"""
        self.maximum = maximum
        self.progress_bar.configure(maximum=maximum)


    #################################################################################################################################


    def append_progress(self):
        """Appending the progress to the bar"""
        self.counter += 1
        self.progress_bar['value'] = self.counter
        self.progress_lbl.configure(text=f'{self.counter} of {self.maximum}')

        self.is_download_complete()

        if self.pausing == True:
            return True
        elif self.canceling == True:
            return True
        else:
            pass


    #################################################################################################################################
    

    def set_counter(self, value):
        """Set the value of the counter if we're continuing"""
        self.counter  = value

    #################################################################################################################################


    def pause(self):
        """Pause the download"""
        self.pausing = True

    #################################################################################################################################


    def cancel(self):
        """Cancel the download"""
        if messagebox.askyesno("Cancel download", "Are you sure you want to cancel this download?"):
            self.canceling = True
            self.destroy()

    #################################################################################################################################


    def download_complete(self, email, main_window):
        """Clears the email list and removes data from shelve file when download complete"""

        try:
            shelfFile = shelve.open('save')
            del shelfFile[email]
            shelfFile.close()

            main_window.populate_email_list()

        except KeyError:
            pass

    def is_download_complete(self):
        if self.progress_bar['value'] == self.maximum:
            self.complete = True


    #################################################################################################################################


    def save_progress(self, email, password, uids, number_of_emails, main_window):
        """This function saves the progress by getting the list of the uids and the index of the last one that was downloaded.
        It also collects the email and password to enable resuming"""

        shelfFile = shelve.open('save')
        shelfFile[str(email)] = {
            'email': email,
            'password': password,
            'uids': uids,
            'last_downloaded': self.counter,
            'number_of_emails': number_of_emails,
        }

        shelfFile.close()

        # supposed to populate the email list in the main window, still to be tested to see if it works
        main_window.populate_email_list()

        # exit the window after we are done saving
        self.destroy()
    
    #################################################################################################################################



class ContinueDownload:

    #################################################################################################################################

    def __init__(self, email, password, number_of_email, uids=None, last_download_index=None):
        """Yeah, initializing everything"""
        self.email = str(email)
        self.password = str(password)
        self.number_of_email = number_of_email
        self.uids = uids
        self.last_download_index = last_download_index


    #################################################################################################################################
        

    def login_to_imap_server(self, main_window, conn_window, status, statusbar):
        """Log in to the imap server"""

        self.main_window = main_window

        # connecting to the server
        status.set('Connecting to emap server')
        statusbar['value'] = 0
        print("Trying to connect to the server")
        
        try:
            imapObj = imapclient.IMAPClient('imap.gmail.com', ssl=True)
            print("Successfully connected to the IMAP server...")

            # Try logging into gmail
            print("Trying to log in to gmail...")
            status.set('Logging into gmail...')
            statusbar['value'] = 50

            try:
                imapObj.login(self.email, self.password)
                print("Logged in")
                status.set('Succesfully logged in!')
                statusbar['value'] = 100
                conn_window.destroy() # Destroys the connection progress window

                self.download = Download(main_window) # Open download window
                if self.last_download_index != None:
                    self.download.set_counter(self.last_download_index)

                self.select_email_uids(imapObj)
            except:
                print("Failed to log you in, make sure your password and email are correct \nand that your have enabled non-google apps in the google settings")
                messagebox.showerror("Login Failed", "Failed to log your in, make sure your password and email are correct \nand that your have enabled non-google apps in the google settings.")
                conn_window.destroy()
        
        except:
            print("Failed to connect, probably some network error")
            conn_window.destroy()
            messagebox.showerror("Connection Failed", "Failed to connect, please check your connection.")


    #################################################################################################################################


    def select_email_uids(self, imap_object):
        """Select uids for email data to be extracted"""
        imap_object.select_folder('INBOX', readonly=True)
        if self.uids == None:
            self.uids = imap_object.search('ALL')
            if str(self.number_of_email) != 'All':
                self.uids = [i for i in self.uids[:int(self.number_of_email)]]
            self.number_of_email = len(self.uids)


            # Configure the progress bar, get the maximum value
            self.download.configure_progressbar(self.number_of_email)



            self.get_email_content_from_uids(imap_object)

            self.logout_of_imap_server(imap_object)

        else:
            self.download.configure_progressbar(self.number_of_email)

            self.get_email_content_from_uids(imap_object)

            self.logout_of_imap_server(imap_object)


    #################################################################################################################################


    #TODO: Create a function that calls envelope with args self.uids
    def get_email_content_from_uids(self, imap_object):
        """Get email data from the respective email uid"""

        # if the last download index is not none, it means we're continuing a download
        

        error_dict = {}


        # If there is a last download index, that means we are continuing a download, otherwise, it's a new one
        if self.last_download_index is not None:
            download_index = self.last_download_index
        else:
            download_index = 0

        for msgid, data in imap_object.fetch(self.uids[download_index:], ['ENVELOPE']).items():
            try:
                envelope = data[b'ENVELOPE']
                rawMessage = imap_object.fetch([msgid], ['BODY[]', 'FLAGS'])
                message = pyzmail.PyzMessage.factory(rawMessage[msgid][b'BODY[]'])

                #function to save to csv
                self.save_data_in_json(envelope, message)


                # Append progress bar
                if self.download.append_progress() == True:
                    break

            except TypeError as e:
                error_dict[msgid] = str(e)

                # Append progress bar
                if self.download.append_progress() == True:
                    break

                continue
            
            except AttributeError as e:
                error_dict[msgid] = str(e)

                # Append progress bar
                if self.download.append_progress() == True:
                    break

                continue

            except UnicodeEncodeError as e:
                error_dict[msgid] = str(e)

                # Append progress bar
                if self.download.append_progress() == True:
                    break

                continue

            except errno.EHOSTDOWN:
                print("Host is down.")

            except Exception as e:
                error_dict[msgid] = str(e)

                # Let's detect a network error
                print(e)
                # if there is a network error, pause the download and exit
                if 'connection' and 'host' and 'closed' in e:
                    messagebox.showerror("Connection Terminated", "Can no longer connect to the host. Pausing download.")
                    self.download.save_progress(self.email, self.password, self.uids, self.number_of_email, self.main_window)
                    self.download.destroy()
                    break

                # Append progress bar
                if self.download.append_progress() == True:
                    break

                # continue
    


        # trying to pause the program
        if self.download.pausing == True:
            print("Program Paused!")
            self.download.save_progress(self.email, self.password, self.uids, self.number_of_email, self.main_window)
        elif self.download.complete == True:
             self.download.download_complete(self.email, self.main_window)
             messagebox.showinfo("Download Completed", "Your download is complete")

       

        
        if error_dict:
            """If there is anything in the error dictionary, it saves the stuff in there into a file for later reference"""
            print("Error Have Been Encountered...")
            print(error_dict)

            with open("errors.txt", 'w') as f:
                f.write(str(error_dict))
            print("Errors have been saved to error dict!")

            #the shelve saving file is going here, might finish the rest tomorrow, this shit might damage my eyes👀
            shelfFile = shelve.open("error_data")
            shelfFile['error_dict'] = error_dict
            shelfFile.close()

        else:
            """Send email or text message, 
            code to do that is below:
            """
            print("The program has finished running and no errors have been encountered!")
            self.download.destroy()

            


    #################################################################################################################################


    def save_data_in_csv(self, envelope, message):
        """Writing the information to a csv file"""

        data_output_file = open('email_data.csv', 'a', newline='')
        csv_writer = csv.writer(data_output_file)
        csv_writer.writerow([envelope.date, message.get_addresses('from'), message.get_addresses('to'), message.get_subject(), message.text_part.get_payload().decode(message.text_part.charset)])
        data_output_file.close()


    #################################################################################################################################


    def save_data_in_json(self,envelope, message):
        """Writing the information to a json file"""

        #1. Convert the data into a dictionary
        email_dict = {
            "date": str(envelope.date),
            "from": message.get_addresses('from'),
            "to": message.get_addresses('to'),
            "subject": str(message.get_subject()),
            "text": str(message.text_part.get_payload().decode(message.text_part.charset)),
        }

        #2. Convert the dictionary into json then dump the shit into a json file
        with open(str(self.email)+".json", 'a') as f:
            f.write(json.dumps(email_dict, sort_keys=True, indent=4))

    #################################################################################################################################


    def logout_of_imap_server(self, imap_object):
        """This function logs out of the imap server"""

        print("Logging Out...")
        imap_object.logout()
        print("Logged Out!!")

    #################################################################################################################################
    

############################################################################################################################################################       
# DELETE THE CODE BELOW, IT AIN'T DO SHIT!

class EmailDownload:
    
    """This Class Downloads emails in the inbox of a particular gmail account and create a csv file with the information"""

    #################################################################################################################################

    def __init__(self, email, password, number_of_email):
        """Yeah, initializing everything"""
        self.email = str(email)
        self.password = str(password)
        self.number_of_email = number_of_email
        

    #################################################################################################################################


    def login_to_imap_server(self, main_window, conn_window, status, statusbar):
        """Log in to the imap server"""

        self.main_window = main_window

        # connecting to the server
        status.set('Connecting to emap server')
        statusbar['value'] = 0
        print("Trying to connect to the server")
        
        try:
            imapObj = imapclient.IMAPClient('imap.gmail.com', ssl=True)
            print("Successfully connected to the IMAP server...")

            # Try logging into gmail
            print("Trying to log in to gmail...")
            status.set('Logging into gmail...')
            statusbar['value'] = 50

            try:
                imapObj.login(self.email, self.password)
                print("Logged in")
                status.set('Succesfully logged in!')
                statusbar['value'] = 100
                conn_window.destroy() # Destroys the connection progress window

                self.download = Download(main_window) # Open download window

                self.select_email_uids(imapObj)
            except:
                print("Failed to log you in, make sure your password and email are correct \nand that your have enabled non-google apps in the google settings")
        
        except:
            print("Failed to connect, probably some network error")

    #################################################################################################################################


    def select_email_uids(self, imap_object):
        """Select uids for email data to be extracted"""
        imap_object.select_folder('INBOX', readonly=True)
        self.uids = imap_object.search('ALL')
        if str(self.number_of_email) != 'All':
            self.uids = [i for i in self.uids[:int(self.number_of_email)]]
        self.number_of_email = len(self.uids)


        # Configure the progress bar, get the maximum value
        self.download.configure_progressbar(self.number_of_email)



        self.get_email_content_from_uids(imap_object)

        self.logout_of_imap_server(imap_object)

    #################################################################################################################################

    #TODO: Create a function that calls envelope with args self.uids
    def get_email_content_from_uids(self, imap_object):
        """Get email data from the respective email uid"""
        counter = 1
        error_dict = {}
        pb = ProgressBar(total=int(self.number_of_email), prefix='Start', suffix='End', decimals=3, length=100, fill='#', zfill='-')

        for msgid, data in imap_object.fetch(self.uids, ['ENVELOPE']).items():
            try:
                envelope = data[b'ENVELOPE']
                rawMessage = imap_object.fetch([msgid], ['BODY[]', 'FLAGS'])
                message = pyzmail.PyzMessage.factory(rawMessage[msgid][b'BODY[]'])

                #function to save to csv
                self.save_data_in_json(envelope, message)

                """Progress Bar Logic"""
                pb.print_progress_bar(counter)
                counter+=1

                # Append progress bar
                if self.download.append_progress() == True:
                    break

            except TypeError as e:
                counter+=1
                error_dict[msgid] = str(e)
                pb.print_progress_bar(counter)

                # Append progress bar
                if self.download.append_progress() == True:
                    break

                continue
            
            except AttributeError as e:
                counter+=1
                error_dict[msgid] = str(e)
                pb.print_progress_bar(counter)

                # Append progress bar
                if self.download.append_progress() == True:
                    break

                continue

            except UnicodeEncodeError as e:
                counter+=1
                error_dict[msgid] = str(e)
                pb.print_progress_bar(counter)

                # Append progress bar
                if self.download.append_progress() == True:
                    break

                continue

            except Exception as e:
                counter+=1
                error_dict[msgid] = str(e)
                pb.print_progress_bar(counter)

                # Append progress bar
                if self.download.append_progress() == True:
                    break

                continue


        # tring to pause the program
        if self.download.pausing == True:
            print("Program Paused!")
            self.download.save_progress(self.email, self.password, self.uids, self.number_of_email, self.main_window)

        
        if error_dict:
            """If there is anything in the error dictionary, it saves the stuff in there into a file for later reference"""
            print("Error Have Been Encountered...")
            print(error_dict)

            with open("errors.txt", 'w') as f:
                f.write(str(error_dict))
            print("Errors have been saved to error dict!")

            #the shelve saving file is going here, might finish the rest tomorrow, this shit might damage my eyes👀
            shelfFile = shelve.open("error_data")
            shelfFile['error_dict'] = error_dict
            shelfFile.close()

        else:
            """Send email or text message, 
            code to do that is below:
            """
            print("The program has finished running and no errors have been encountered!")
            self.download.destroy()
            messagebox.showinfo("Download Completed", "Your download is complete")

    #################################################################################################################################


    def save_data_in_csv(self, envelope, message):
        """Writing the information to a csv file"""

        data_output_file = open('email_data.csv', 'a', newline='')
        csv_writer = csv.writer(data_output_file)
        csv_writer.writerow([envelope.date, message.get_addresses('from'), message.get_addresses('to'), message.get_subject(), message.text_part.get_payload().decode(message.text_part.charset)])
        data_output_file.close()

    #################################################################################################################################


    def save_data_in_json(self,envelope, message):
        """Writing the information to a json file"""

        #1. Convert the data into a dictionary
        email_dict = {
            "date": str(envelope.date),
            "from": message.get_addresses('from'),
            "to": message.get_addresses('to'),
            "subject": str(message.get_subject()),
            "text": str(message.text_part.get_payload().decode(message.text_part.charset)),
        }

        #2. Convert the dictionary into json then dump the shit into a json file
        with open("email_data.json", 'a') as f:
            f.write(json.dumps(email_dict, sort_keys=True, indent=4))

    #################################################################################################################################


    def logout_of_imap_server(self, imap_object):
        """This function logs out of the imap server"""

        print("Logging Out...")
        imap_object.logout()
        print("Logged Out!!")
    
    #################################################################################################################################


