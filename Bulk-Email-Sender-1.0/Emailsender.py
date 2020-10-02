# ===========================================Importing Modules=======================================
# ===========================================Importing Modules=======================================
# ===========================================Importing Modules=======================================
# ===========================================Importing Modules=======================================


from tkinter import*
from tkinter import ttk
from tkinter import messagebox, filedialog
from PIL import ImageTk, Image
import os
import pandas as pd
import time
import smtplib
from email.message import EmailMessage
import imghdr



class bulk_email():
    def __init__(self, root):
        self.root = root
        self.root.title("Bulk Email Sender".center(320))
        self.root.geometry("1000x550+200+75")
        self.root.resizable(False, False)
        self.root.config(bg="black")

        # =================Icons==============================
        # =================Icons==============================
        # =================Icons==============================
        # =================Icons==============================

        self.email_icon = ImageTk.PhotoImage(file="Pics\Email.jpg")
        self.setting_icon = ImageTk.PhotoImage(file="Pics\Settings.jpg")
        self.show_icon = ImageTk.PhotoImage(file="Pics\show-password.png")
        self.file_icon = ImageTk.PhotoImage(file="Pics\images (1).png")
        self.image_icon = ImageTk.PhotoImage(file="Pics\images.png")

        # *******************Variable*******************************

        self.pass_l = StringVar()
        self.path_ = StringVar()
        self.path_.set("")
        self.path_1 = StringVar()
        self.path_1.set("")
        self.L1 = []

        # ====================Title Bar=============================
        # ====================Title Bar=============================
        # ====================Title Bar=============================
        # ====================Title Bar=============================

        title = Label(self.root, text="Bulk Email Sender ", bg="#152238", fg="white", image=self.email_icon,
                      compound=LEFT, font=("Goudy Old Style", 48, "bold"), anchor="w").place(x=0, y=0, relwidth=1)
        sub_title = Label(self.root, text="Use Excel File to Send the Bulk Email at once, with just one column. Just Ensure the Email column has email.",
                          bg="yellow", font=("Calibri (Body)", 14, "bold"), anchor="w").place(x=0, y=95, relwidth=1)
        text_button = Button(self.root, bd=5, image=self.setting_icon,
                              command=self.setting_win,height=80, cursor="hand2").place(x=900, y=0)

        # =====================Body=======================================
        # =====================Body=======================================
        # =====================Body=======================================
        # =====================Body=======================================

        self.var_choice = StringVar()

        single = Radiobutton(self.root, value="single", variable=self.var_choice, bg="black",
    
                             activebackground="black", command=self.check).place(x=50, y=167)
        single_lbl = Label(self.root, bg="black", fg="white", text="Single", font=(
            "times new roman", 30, "bold")).place(x=80, y=155)

        bulk = Radiobutton(self.root, value="bulk", variable=self.var_choice, bg="black",
                           activebackground="black", command=self.check).place(x=200, y=167)
        bulk_lbl = Label(self.root, bg="black", fg="white", text="Bulk", font=(
            "times new roman", 30, "bold")).place(x=230, y=155)

        self.var_choice.set("single")

        # ====================Entries=======================================
        # ====================Entries=======================================
        # ====================Entries=======================================
        # ====================Entries=======================================

        to_lbl = Label(self.root, text="To(Email Address)", font=(
            "time new roman", 18, "bold"), fg="white", bg="black").place(x=50, y=220)

        sub_lbl = Label(self.root, text="SUBJECT", font=(
            "time new roman", 18, "bold"), fg="white", bg="black").place(x=50, y=280)

        msg_lbl = Label(self.root, text="MESSAGE", font=(
            "time new roman", 18, "bold"), fg="white", bg="black").place(x=50, y=340)

        sender_lbl = Label(self.root, text="Sender Name", font=(
            "time new roman", 18, "bold"), fg="white", bg="black").place(x=500, y=125)

        self.to_entry = Entry(self.root, bd=5, width=20,
                              bg="lightgrey",  font=("times new roman", 18))
        self.to_entry.place(x=300, y=220)

        self.sub_entry = Entry(self.root, bd=5, width=30,
                               bg="lightgrey", font=("times new roman", 18))
        self.sub_entry.place(x=300, y=280)

        self.msg_entry = Text(self.root, bd=5, bg="lightgrey",
                              font=("times new roman", 12))
        self.msg_entry.place(x=300, y=340, width=670, height=120)

        self.msg_entry.insert("1.0","Hi!\nSir/Mamam,\nHello this mail is from Thomson Press India Ltd. to Mr./Ms ")

        self.sender_entry = Entry(self.root, bd=5, width=25,
                              font=("times new roman", 18))
        self.sender_entry.place(x=685, y=125)


        # ============================Status========================================
        # ============================Status========================================
        # ============================Status========================================
        # ============================Status========================================

        self.Total_lbl = Label(self.root, font=(
            "time new roman", 18, "bold"), fg="white", bg="black")
        self.Total_lbl.place(x=50, y=500)

        self.sent_lbl = Label(self.root,  font=(
            "time new roman", 18, "bold"), fg="Green", bg="black")
        self.sent_lbl.place(x=300, y=500)

        self.left_lbl = Label(self.root,  font=(
            "time new roman", 18, "bold"), fg="Yellow", bg="black")
        self.left_lbl.place(x=400, y=500)

        self.failed_lbl = Label(self.root, font=(
            "time new roman", 18, "bold"), fg="RED", bg="black")
        self.failed_lbl.place(x=500, y=500)

        # ===========================Buttons=======================================
        # ===========================Buttons=======================================
        # ===========================Buttons=======================================
        # ===========================Buttons=======================================

        self.browser_btn = Button(self.root, text="Browse", font=(
            "times new roman", 18, "bold"), activebackground="grey", activeforeground="black", bg="grey", cursor="hand2", state="disable", command=self.browse_file)
        self.browser_btn.place(x=570, y=220, width=120, height=40)

        clear_btn = Button(self.root, text="Clear", font=("times new roman", 18, "bold"), activebackground="#262626",
                           activeforeground="white", bg="#262626", fg="white", cursor="hand2", command=self.clear1).place(x=700, y=500, width=120, height=40)

        send_btn = Button(self.root, text="Send", font=("times new roman", 18, "bold"), activebackground="#00B0F0", activeforeground="white",
                          bg="#00B0F0", bd=0,fg="white", cursor="hand2", command=self.send_email).place(x=850, y=500, width=120, height=40)

        attach_file_button = Button(self.root,image=self.file_icon, relief = GROOVE, activebackground="black",
                          bg="black", fg="white", bd=3, cursor="hand2", command=self.attachments_).place(x=750, y=220)

        attach_img_button = Button(self.root,image=self.image_icon, relief = GROOVE, activebackground="black",
                          bg="black", fg="white", bd=3, cursor="hand2", command=self.attachments_1).place(x=825, y=220)
        
        attach_show_button = Button(self.root,text="Show Attachments", relief = GROOVE, activebackground="black",
                          bg="black", fg="white", bd=3, cursor="hand2", command=self.show_attachment).place(x=750, y=260)

        self.check_exist_file()

# ***********************************************************************************************************************************

        # =====================================Functions=================================
        # =====================================Functions=================================
        # =====================================Functions=================================
        # =====================================Functions=================================

# ***********************************************************************************************************************************

        # ===========================Child Window ==========================================
        # ===========================Child Window ==========================================
        # ===========================Child Window ==========================================

        # **************************************************************************************************************

    def attachments_(self):
        self.file_dialog= filedialog.askopenfilename(initialdir='/', title="Choose Document to attach", filetype=((".doc","*doc"),(".docx","*docx"), (".html","*html"), (".htm","*htm"), (".txt","*txt"),(".odt","*odt"),(".pdf","*pdf"),(".xls","*xls"),(".xlsx","*xlsx"),(".csv","*csv"),(".ods","*ods"),(".ppt","*ppt"),(".pptx","*ppt"),(".xml","*xml"),("RTF","*rtf"),("All Files","*.*")))
        self.path_1 = self.file_dialog
        self.L1.append(self.path_1)

    def attachments_1(self):
        self.file_dialog= filedialog.askopenfilename(initialdir='/', title="Choose Image to attach", filetype=((".jpg","*jpg"),(".jpeg","*jpeg"), (".png","*png"), (".tif","*tif"), (".tiff","*tiff"),(".bmp","*bmp"),(".gif","*gif"),(".eps","*eps"),(".raw","*raw"),(".cr2","*cr2"),(".nef","*nef"),(".orf","*orf"),(".sr2","*sr2"),("All Files","*.*")))
        self.path_ = self.file_dialog
        self.L1.append(self.path_)

    
    def show_attachment(self):
        self.root3 = Toplevel()  # Child Window "Tk() can Also be use here"
        self.root3.title("Attached Files")
        self.root3.geometry("700x400+350+150")
        self.root3.configure(bg="black")
        self.root3.focus_force()  # Fouce on Child Window
        self.root3.grab_set()
        #self.root3.resizable(False, False)

        self.show_entry = Text(self.root3,fg="white", bd=10, bg="black",
                              font=("times new roman", 15))
        self.show_entry.place(x=0, y=30, relwidth=1, height=320)

        scroll_x=Scrollbar(self.show_entry,cursor="hand2", orient=HORIZONTAL)
        scroll_y=Scrollbar(self.show_entry, cursor="hand2",orient=VERTICAL)
        scroll_x.pack(side=BOTTOM,fill=X)
        scroll_y.pack(side=RIGHT,fill=Y)
        scroll_x.config(command=self.show_entry.xview)
        scroll_y.config(command=self.show_entry.yview)
        self.show_entry.config( yscrollcommand = scroll_y)
        self.show_entry.config( xscrollcommand = scroll_x)

        Exit3_button = Button(self.root3,text="Exit", relief = GROOVE, font=("",12,"bold"),activebackground="red",
                          activeforeground="white",bg="red", fg="white", bd=3, cursor="hand2",width=8 ,command=self.exit_win3).place(x=607, y=350)


        for i in self.L1:
            self.show_entry.insert("1.0","File :"+str(i)+"\n")

    def exit_win3(self):
        self.root3.destroy()

    def setting_win(self):
        self.check_exist_file()
        self.root2 = Toplevel()  # Child Window "Tk() can Also be use here"
        self.root2.title("Setting")
        self.root2.geometry("700x320+350+150")
        self.root2.configure(bg="black")
        self.root2.focus_force()  # Fouce on Child Window
        self.root2.grab_set()  # Hold the Child window till closinf of the window
        self.root2.resizable(False, False)
        # =====================================Child Window Title=================================
        # =====================================Child Window Title=================================
        # =====================================Child Window Title=================================
        # =====================================Child Window Title=================================

        title_child = Label(self.root2, text="Credentials Settings", bg="#152238", fg="white", image=self.setting_icon,
                            compound=LEFT, font=("Goudy Old Style", 48, "bold"), anchor="w").place(x=0, y=0, relwidth=1)
        sub_title_child = Label(self.root2, text="Enter Email adress and password from which you want send the email",
                                bg="yellow", font=("Calibri (Body)", 14, "bold"), anchor="w").place(x=0, y=95, relwidth=1)

        # ====================Child Window Entries=======================================
        # ====================Child Window Entries=======================================
        # ====================Child Window Entries=======================================
        # ====================Child Window Entries=======================================

        from_lbl = Label(self.root2, text="Email Address", font=(
            "time new roman", 18, "bold"), fg="white", bg="black").place(x=30, y=150)

        pass_lbl = Label(self.root2, text="Password", font=(
            "time new roman", 18, "bold"), fg="white", bg="black").place(x=30, y=200)

        self.from_entry = Entry(
            self.root2, bd=5, width=30, bg="lightgrey", font=("times new roman", 18))
        self.from_entry.place(x=230, y=150)

        self.pass_entry = Entry(
            self.root2, bd=5, width=30, show="*", bg="lightgrey", font=("times new roman", 18))
        self.pass_entry.place(x=230, y=200)

        # ===========================Child Window Buttons===================================
        # ===========================Child Window Buttons===================================
        # ===========================Child Window Buttons===================================
        # ===========================Child Window Buttons===================================

        clear_btn = Button(self.root2, text="Clear", font=("times new roman", 18, "bold"), activebackground="#262626",
                           activeforeground="white", bg="#262626", fg="white", cursor="hand2", command=self.clear2).place(x=300, y=260, width=140, height=30)

        save_btn = Button(self.root2, text="Save", font=("times new roman", 18, "bold"), activebackground="#00B0F0",
                          activeforeground="white", bg="#00B0F0", fg="white", cursor="hand2", command=self.save).place(x=465, y=260, width=140, height=30)

        
        self.show_pass_btn = Button(self.root2, image=self.show_icon,
                                    activebackground="black", bg="black", cursor="hand2", command=self.show)
        self.show_pass_btn.place(x=610, y=200, height=40)

        self.pass_lbl = Entry(self.root2, font=(
            "time new roman", 10), textvariable=self.pass_l, fg="white", bg="black").place(x=0, y=125)
        self.from_entry.insert(0, self.from_)
        self.pass_entry.insert(0, self.pass_)
        self.pass_l.set("Password Mode: Hidden")

    
    
    def clear1(self):
        self.to_entry.config(state="normal")
        self.to_entry.delete(0, END)
        self.msg_entry.delete("1.0", END)
        self.sub_entry.delete(0, END)
        self.var_choice.set("single")
        # can be written as state="disable"
        self.browser_btn.config(state=DISABLED)
        self.Total_lbl.config(text="")
        self.sent_lbl.config(text="")
        self.failed_lbl.config(text="")
        self.left_lbl.config(text="")
        self.L1.clear()

    
    
    def clear2(self):
        self.from_entry.delete(0, END)
        self.pass_entry.delete(0, END)

    
    
    def check(self):  # Checking single mail or bulk mail
        if self.var_choice.get() == "single":
            messagebox.showinfo(
                "Successfull ", "Single Email ", parent=self.root)
            self.browser_btn.config(state="disable")
            self.to_entry.config(state=NORMAL)
            self.to_entry.delete(0, END)
            self.sub_entry.delete(0,END)
            self.msg_entry.delete("1.0",END)
            self.msg_entry.insert("1.0","Hi!\nSir/Mamam,\nHello this mail is from Thomson Press India Ltd. to Mr./Ms ")
            self.Total_lbl.config(text="")
            self.sent_lbl.config(text="")
            self.failed_lbl.config(text="")
            self.left_lbl.config(text="")

        if self.var_choice.get() == "bulk":
            messagebox.showinfo("Successfull ", "Bulk Email", parent=self.root)
            self.browser_btn.config(state="normal")
            self.to_entry.delete(0, END)
            self.to_entry.delete(0, END)
            self.sub_entry.delete(0,END)
            self.msg_entry.delete("1.0",END)
            self.msg_entry.insert("1.0","Hi!\nSir/Mamam,\nHello this mail is from Thomson Press India Ltd. to Mr./Ms ")
            self.to_entry.config(state="readonly")


    def check_exist_file(self):
        if os.path.exists("credentials.txt") == False:
            f = open("credentials.txt", "w")
            f.write(",")
            f.close()

        f2 = open("credentials.txt", "r")
        self.credentials = []
        for i in f2:
            
            self.credentials.append([i.split(",")[0], i.split(",")[1]])
        
        self.from_ = self.credentials[0][0]
        self.pass_ = self.credentials[0][1]

    def save(self):
        if self.from_entry.get() == "" or self.pass_entry.get() == "":
            return messagebox.showerror(
                "Error", "All fields are required", parent=self.root2)

        elif "@" not in self.from_entry.get():
            return messagebox.showerror("Error", "Email should have '@' character", parent=self.root2)

        else:
            f = open("credentials.txt", "w")
            f.write(str(self.from_entry.get())+","+str(self.pass_entry.get()))
            f.close()
            messagebox.showinfo("Info", "Email Saved", parent=self.root2)
            self.check_exist_file()

    def show(self):
        a = self.pass_entry.get()
        self.pass_entry = Entry(
            self.root2, bd=5, width=30, bg="lightgrey", font=("times new roman", 18))
        self.pass_entry.place(x=230, y=200)
        self.pass_entry.insert(0, a)
        if self.pass_l.get() == "Password Mode: Hidden":
            self.pass_l.set("Password Mode: Shown")
            self.pass_entry = Entry(
                self.root2, bd=5, width=30, bg="lightgrey", font=("times new roman", 18))
            self.pass_entry.place(x=230, y=200)
            self.pass_entry.insert(0, a)

        elif self.pass_l.get() == "Password Mode: Shown":
            self.pass_l.set("Password Mode: Hidden")
            self.pass_entry = Entry(
                self.root2, bd=5, width=30, bg="lightgrey", show="*", font=("times new roman", 18))
            self.pass_entry.place(x=230, y=200)
            self.pass_entry.insert(0, a)

    def browse_file_name(self):
        #dialog = filedialog.askopenfile(
        #    initialdir='/', title="Select Excel File", filetypes=(("All Files", "*.*"), ("Excel Files", "xlsx")))
        if self.dialog != None:
            
            try:
                data = pd.read_excel(self.dialog.name)    
            except Exception:
                return messagebox.showerror("Wrong Format Error","File shold be in .xlsx format")
            
            L1=[]
            if 'Name' in data.columns:

                self.names= list(data['Name'])
            
                for i in self.names:
                    if pd.isnull(i) == False:
                        # print(i)
                        L1.append(i)
                self.names = L1
                if len(self.names) > 0:
                    pass
                else:
                    messagebox.showerror(
                        "Error", "This file haven't any emails", parent=self.root)

            elif 'name' in data.columns:
                self.names= list(data['Name'])
                
                for i in self.names:
                    if pd.isnull(i) == False:
                        # print(i)
                        L1.append(i)
                self.names = L1
                if len(self.names) > 0:
                    pass
                else:
                    messagebox.showerror(
                        "Error", "This file haven't any emails", parent=self.root)              
            else:
                messagebox.showerror(
                    "Error", "Please Select File Which have 'Name' or 'Name' columns", parent=self.root)
    
    def browse_file(self):
        #self.browse_file_name()
        self.dialog = filedialog.askopenfile(
            initialdir='/', title="Select Excel File", filetypes=(("All Files", "*.*"), ("Excel Files", "xlsx")))
        if self.dialog != None:
            
            try:
                data = pd.read_excel(self.dialog.name)    
            except Exception:
                return messagebox.showerror("Wrong Format Error","File shold be in .xlsx format")
            
            
            if 'Email' in data.columns:

                self.names= list(data['Name'])
            
                self.emails = list(data['Email'])
                L = []
                for i in self.emails:
                    if pd.isnull(i) == False:
                        # print(i)
                        L.append(i)
                self.emails = L
                if len(self.emails) > 0:
                    self.to_entry.config(state=NORMAL)
                    self.to_entry.delete(0, END)
                    self.to_entry.insert(0, str(self.dialog.name.split("/")[-1]))
                    self.to_entry.config(state="readonly")
                    self.Total_lbl.config(text="Total: "+str(len(self.emails)))
                    self.sent_lbl.config(text="Sent: ")
                    self.failed_lbl.config(text="Failed: ")
                    self.left_lbl.config(text="Left: ")
                else:
                    messagebox.showerror(
                        "Error", "This file haven't any emails", parent=self.root)

            elif 'email' in data.columns:
                # print("Exist")
                self.emails = list(data['email'])

                L = []
                for i in self.emails:
                    if pd.isnull(i) == False:
                        # print(i)
                        L.append(i)
                self.emails = L
                if len(self.emails) > 0:
                    self.to_entry.config(state=NORMAL)
                    self.to_entry.delete(0, END)
                    self.to_entry.insert(0, str(dialog.name.split("/")[-1]))
        
        
        
                    self.to_entry.config(state="readonly")
                    self.Total_lbl.config(text="Total: "+str(len(self.emails)))
                    self.sent_lbl.config(text="Sent: ")
                    self.failed_lbl.config(text="Failed: ")
                    self.left_lbl.config(text="Left: ")
                else:
                    messagebox.showerror(
                        "Error", "This file haven't any emails", parent=self.root)

            else:
                messagebox.showerror(
                    "Error", "Please Select File Which have 'Email' or 'email' columns", parent=self.root)
    

    
    def send_email(self):
        path_=self.path_
        path_1 = self.path_1
        a1 = len(self.msg_entry.get("1.0", END))
        if self.to_entry.get() == "" and self.sub_entry.get() == "" and self.sender_entry.get()=="" and self.sub_entry.get() == "" :
            return messagebox.showerror(
                "Error", "All fields are required", parent=self.root)

        if self.to_entry.get() == "":
            return messagebox.showerror(
                "Error", "To field is empty please write valid email", parent=self.root)
        
        if   self.sub_entry.get() == "" :
            return messagebox.askyesnocancel(
                "Error", "Subject", parent=self.root)
        if  a1 == 1 :
            return messagebox.showerror(
                "Error", "Message Fielf Empty", parent=self.root)

        if self.sender_entry.get()=="":
            messagebox.showerror(
                "Error", "Sender Name Required", parent=self.root)
        
                
        else:
            if self.var_choice.get() == "single":
                self.Email_sent(self.to_entry.get(), self.sub_entry.get(
                ), self.msg_entry.get("1.0", END), self.from_, self.pass_, path_, path_1)
    
            if self.var_choice.get() == "bulk":
                self.failed = []
                self.s_count = 0
                self.f_count = 0
                for a,b in zip(self.emails,self.names):
                    self.Email_sent1(a,b, self.sub_entry.get(
                    ), self.msg_entry.get("1.0", END), self.from_, self.pass_,path_,path_1)
                    if self.valuegain == "S":
                        self.s_count += 1
                    if self.valuegain == "F":
                        self.f_count += 1
                    self.statusbar()
                    time.sleep(1)
                messagebox.showinfo(
                    "Successfull ", "Email Sent, Check Status", parent=self.root)

    def statusbar(self):
        self.Total_lbl.config(text="Status: "+str(len(self.emails))+"=>>")
        self.sent_lbl.config(text="Sent: "+str(self.s_count))
        self.failed_lbl.config(text="Failed: "+str(self.f_count))
        self.left_lbl.config(
            text="Left: "+str(len(self.emails)-((self.s_count)+(self.f_count))))
        self.Total_lbl.update()
        self.sent_lbl.update()
        self.failed_lbl.update()
        self.left_lbl.update()


       
 
    def Email_sent(self,to_, sub_, msg_, from_, pass_,path_,path_1):

        send = smtplib.SMTP("smtp.gmail.com", 587)  # Create session for Gmail
        send.starttls()  # transport layer

        #===================Plain Text=====================

        msg = EmailMessage()
        msg["Subject"] = sub_
        msg["From"] = from_
        msg["To"] = to_
        msg.set_content(msg_+" "+"\nThanks and Regards\n"+self.sender_entry.get())        
             
        try:
            send.login(from_, pass_)
        except smtplib.SMTPAuthenticationError:
            messagebox.showerror("Error","Username or Password is wrong")    

            #===================Attachments====================
        for i in self.L1:    
            print(i)    
            try:
                try:
                    with open(i,"rb") as f:
                        file_data = f.read()
                        file_type = imghdr.what(f.name)
                        file_name = f.name
                    msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)
                except TypeError:
                    msg.add_attachment(file_data, maintype="image", subtype=file_type, filename=file_name)

            except Exception:
                pass
        
                        
        try:
            try:
                try:
                    send.send_message(msg)
                    messagebox.showinfo("Mailed","Successfully Mailed Your Conent")
                except smtplib.SMTPRecipientsRefused:
                    messagebox.showerror("Mailed","Mail Not Sent")
            except smtplib.SMTPException:
                messagebox.showerror("Mailed","Mail Not Sent")
        except smtplib.SMTPConnectError:
            messagebox.showerror("Error","Connection Error")
         
    def Email_sent1(self,to_, names_,sub_, msg_, from_, pass_,path_,path_1):

        self.valuegain = ""                
        send = smtplib.SMTP("smtp.gmail.com", 587)  # Create session for Gmail
        send.starttls()  # transport layer

        #===================Plain Text=====================

        msg = EmailMessage()
        msg["Subject"] = sub_
        msg["From"] = from_
        msg["To"] = to_
        msg.set_content(msg_+" "+names_+"\nThanks and Regards\n"+self.sender_entry.get())        
             
        try:
            send.login(from_, pass_)
        except smtplib.SMTPAuthenticationError:
            messagebox.showerror("Error","Username or Password is wrong")    

            #===================Attachments====================

        for i in self.L1:    
            #print(i)    
            try:
                try:
                    with open(i,"rb") as f:
                        file_data = f.read()
                        file_type = imghdr.what(f.name)
                        file_name = f.name
                    msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)
                except TypeError:
                    msg.add_attachment(file_data, maintype="image", subtype=file_type, filename=file_name)

            except Exception:
                pass
        
        

        try:
            try:
                try:
                    send.send_message(msg)
                    self.valuegain = "S"
                except smtplib.SMTPRecipientsRefused:
                    self.valuegain = "F"                        
            except smtplib.SMTPException:
                self.valuegain = "F"
        except smtplib.SMTPConnectError:
            messagebox.showerror("Error","Connection Error")
    
       
              
 
root = Tk()
obj = bulk_email(root)
root.mainloop()
