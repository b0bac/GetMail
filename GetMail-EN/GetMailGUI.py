import platform
import tkinter as tk
from tkinter import messagebox
from exchangelib import Account, Credentials, Configuration, FileAttachment, DELEGATE


EmailFormatString = """
**********************************************************************************\n
Email ID: %s\n
Sender: %s(%s)\n
CC: %s\n
BCC: %s\n
Time: %s\n
Subject: %s\n
**********************************************************************************\n
Body: %s\n
"""


class MailToolGuiWindows:
    def __init__(self):
        self.EmailObject = None
        self.Windows = tk.Tk()
        self.Windows.geometry("1200x750")
        self.Windows.title("GetMail v1.1 - English")
        self.Username = None
        self.UsernameLaber = tk.Label(self.Windows, text="Account:")
        self.UsernameInput = tk.Entry(self.Windows, width=20)
        self.Credential = None
        self.CredentialLaber = tk.Label(self.Windows, text="Password / NTLM Hash:")
        self.CredentialInput = tk.Entry(self.Windows, width=30)
        self.MailAddress = None
        self.PageSelect = None
        self.Page = None
        self.PageSize = 20
        self.MailAddressLaber = tk.Label(self.Windows, text="Email Address:")
        self.MailAddressInput = tk.Entry(self.Windows, width=30)
        self.AdvanceLaber = tk.Label(self.Windows, text="Advanced Options:")
        self.AutoDiscoverFlag = 1
        self.AutoDiscoverVar = tk.IntVar()
        self.AutoDiscoverLaber = tk.Label(self.Windows, text="Autodiscovery:")
        self.AutoDiscoverEnableLabel = tk.Label(self.Windows, text="Enable:")
        self.AutoDiscoverDisableLabel = tk.Label(self.Windows, text="Disable:")
        self.AutoDiscoverEnable = tk.Radiobutton(self.Windows, value=1, variable=self.AutoDiscoverVar, command=self.SetAutoDiscoverFlag)
        self.AutoDiscoverDisable = tk.Radiobutton(self.Windows, value=2, variable=self.AutoDiscoverVar, command=self.SetAutoDiscoverFlag)
        self.MailServer = None
        self.MailServerLabel = tk.Label(self.Windows, text="Exchange Server URL:")
        self.MailServerInput = tk.Entry(self.Windows, width=30)
        self.Keyword = None
        self.KeywordSearchLabel = tk.Label(self.Windows, text="Keyword Search:")
        self.KeywordSearchInput = tk.Entry(self.Windows, width=20)
        self.LoginButton = tk.Button(self.Windows, width=15, text="Start Collection", bg="blue", command=self.Login)
        self.LogoutButton = tk.Button(self.Windows, width=15, text="End Collection", command=self.Logout)
        self.SearchButton = tk.Button(self.Windows, width=15, text="Start Search", command=self.Search)
        self.FolderListLabel = tk.Label(self.Windows, text="Folder:")
        self.Folder = tk.StringVar()
        self.FolderListBox = tk.Listbox(self.Windows, selectmode=tk.SINGLE, height=30, listvariable=self.Folder)
        self.FolderSelectButton = tk.Button(self.Windows, height=1, width=15, text="Select Folder", command=self.FolderSelected)
        self.FolderSelectNextPageButton = tk.Button(self.Windows, height=1, width=6, text="Next Page", command=self.FolderSelectedNext)
        self.MailListLabel = tk.Label(self.Windows, text="Mail:")
        self.Mail = tk.StringVar()
        self.MailListBox = tk.Listbox(self.Windows,selectmode=tk.SINGLE, height=30, listvariable=self.Mail)
        self.MailSelectButton = tk.Button(self.Windows, height=1, width=6, text="Select", command=self.MailSelected)
        self.MailLabel = tk.Label(self.Windows, text="Contents:")
        self.MailText = tk.Text(self.Windows, height=39, relief=tk.RAISED, width=80, bg="gray")
        self.MailText.config(state="disable")
        self.AttachmentListLabel = tk.Label(self.Windows, text="Attachments:")
        self.AttachmentListBox = tk.Listbox(self.Windows, selectmode=tk.SINGLE, height=30)
        self.AttachmentDownloadButton = tk.Button(self.Windows, height=1, width=15, text="Download Attachment", command=self.DownloadAttachment)
        self.MailList = {}
        self.AttachmentList = []
        self.Mail = None
        self.AuthorLabel = tk.Label(self.Windows, text="Author: crsecscu@gmail.com, English Translation: @vysecurity")
        self.LoginButtonLabel = tk.Label(self.Windows, text="Login", font = "Helvetica -12 bold")
        self.LogoutButtonLabel = tk.Label(self.Windows, text="Logout", font = "Helvetica -12 bold")
        self.SearchButtonLabel = tk.Label(self.Windows, text="Search", font = "Helvetica -12 bold")
        self.FolderButtonLabel = tk.Label(self.Windows, text="Select", font = "Helvetica -12 bold")
        self.MailButtonLabel = tk.Label(self.Windows, text="Read", font = "Helvetica -12 bold")
        self.AttachmentButtonLabel = tk.Label(self.Windows, text="Download", font = "Helvetica -12 bold")

    def Graph(self):
        self.UsernameLaber.place(x=40, y=40)
        self.UsernameInput.place(x=110, y=40)
        self.CredentialLaber.place(x=300, y=40)
        self.CredentialInput.place(x=370, y=40)
        self.MailAddressLaber.place(x=650, y=40)
        self.MailAddressInput.place(x=720, y=40)
        self.AdvanceLaber.place(x=40, y=80)
        self.AutoDiscoverLaber.place(x=40, y=120)
        self.AutoDiscoverEnableLabel.place(x=160, y=120)
        self.AutoDiscoverEnable.place(x=200, y=120)
        self.AutoDiscoverDisableLabel.place(x=240, y=120)
        self.AutoDiscoverDisable.place(x=280, y=120)
        self.MailServerLabel.place(x=320, y=120)
        self.MailServerInput.place(x=420, y=120)
        self.KeywordSearchLabel.place(x=700, y=120)
        self.KeywordSearchInput.place(x=800, y=120)
        self.LoginButton.config(bg="green")
        self.LoginButton.place(x=1000, y=40)
        self.LogoutButton.config(bg="green")
        self.LogoutButton.place(x=1000, y=80)
        self.SearchButton.config(bg="green")
        self.SearchButton.place(x=1000, y=120)
        self.FolderListLabel.place(x=40, y=160)
        self.FolderListBox.place(x=40, y=200)
        self.FolderSelectButton.place(x=50, y=715)
        self.MailListLabel.place(x=220, y=160)
        self.MailListBox.place(x=220, y=200)
        self.MailSelectButton.place(x=220, y=715)
        self.FolderSelectNextPageButton.place(x=300, y=715)
        self.AuthorLabel.place(x=400 ,y=715)
        self.AttachmentListLabel.place(x=980, y=160)
        self.AttachmentListBox.place(x=980, y=200)
        self.AttachmentDownloadButton.config(bg="green")
        self.AttachmentDownloadButton.place(x=990, y=715)
        self.MailLabel.place(x=395, y=160)
        self.MailText.place(x=395, y=196)
        if platform.system() == "Darwin":
            pass
            #self.LoginButtonLabel.place(x=1058, y=45)
            #self.LogoutButtonLabel.place(x=1058, y=85)
            #self.SearchButtonLabel.place(x=1058, y=125)
            #self.FolderButtonLabel.place(x=108, y=720)
            #self.MailButtonLabel.place(x=288, y=720)
            #self.AttachmentButtonLabel.place(x=1048, y=720)
        self.Windows.mainloop()

    def ShowMessage(self, title, message):
        messagebox.showinfo(title, message)

    def Login(self):
        try:
            self.Username = self.UsernameInput.get()
        except Exception as exception:
            print("[+]0 %s"%str(exception))
            self.ShowMessage("警告", exception)
        try:
            self.Credential = self.CredentialInput.get()
        except Exception as exception:
            print("[+]1 %s" % str(exception))
            self.ShowMessage("警告", exception)
        try:
            self.MailAddress = self.MailAddressInput.get()
        except Exception as exception:
            print("[+]2 %s" % str(exception))
            self.ShowMessage("警告", exception)
        try:
            self.MailServer = self.MailServerInput.get()
        except Exception as exception:
            print("[+]3 %s" % str(exception))
            print(self.MailServer)
            self.MailServer = None
        try:
            self.Keyword = self.KeywordSearchInput.get()
        except Exception as exception:
            print("[+]4 %s" % str(exception))
            self.Keyword = None
        self.SetAutoDiscoverFlag()
        if self.AutoDiscoverFlag == 1:
            try:
                self.EmailObject = Account(self.MailAddress, credentials=Credentials(self.Username, self.Credential), autodiscover=True)
                self.ShowMessage("Info", "Connection Success")
            except Exception as exception:
                try:
                    self.EmailObject = Account(self.MailAddress, credentials=Credentials(self.Username, "00000000000000000000000000000000:"+self.Credential), autodiscover=True)
                    self.ShowMessage("Info", "Connection Success")
                except Exception as exception:
                    self.ShowMessage("Error", "Connection Failed")
        elif self.AutoDiscoverFlag == 2:
            try:
                self.EmailObject = Account(self.MailAddress, config=Configuration(server=self.MailServer, credentials=Credentials(self.Username, self.Credential)), autodiscover=False, access_type=DELEGATE)
                self.ShowMessage("Info", "Connection Success")
            except Exception as exception:
                try:
                    self.EmailObject = Account(self.MailAddress, config=Configuration(server=self.MailServer, credentials=Credentials(self.Username, "00000000000000000000000000000000:"+self.Credential)), autodiscover=False, access_type=DELEGATE)
                    self.ShowMessage("Info", "Connection Success")
                except Exception as exception:
                    self.ShowMessage("Error", "Connection Failed")
        else:
            self.ShowMessage("Error", "Connection Failed")
        if self.EmailObject != None:
            string = self.EmailObject.root.tree()
            name = None
            _list = []
            lines = string.split("\n")
            for line in lines:
                if line.find("│   ├── ") == 0:
                    name = line.split("│   ├── ")[-1]
                elif line.find("│   └── ") == 0:
                    name = line.split("│   └── ")[-1]
                _list.append(name)
            _list = list(set(_list))
            self.Folder = {}
            index = 0
            for name in _list:
                if name == None:
                    continue
                self.Folder[str(index)] = name
                index += 1
                self.FolderListBox.insert(tk.END, name)

    def Logout(self):
        try:
            self.EmailObject.close()
        except Exception as exception:
            del self.EmailObject
        self.EmailObject = None
        self.FolderListBox.delete(0, tk.END)
        self.MailListBox.delete(0, tk.END)
        self.AttachmentListBox.delete(0, tk.END)
        self.ShowMessage("Info", "Connection Closed")

    def SetAutoDiscoverFlag(self):
        try:
            self.AutoDiscoverFlag = self.AutoDiscoverVar.get()
        except Exception as exception:
            self.AutoDiscoverFlag = 1

    def Search(self):
        count = 0
        try:
            if isinstance(self.MailList, dict):
                count = len(self.MailList)
            else:
                count = self.MailList.count()
        except Exception as exception:
            count = 0
        if self.MailListBox.size == 0:
            self.ShowMessage("Info", "No emails found")
        else:
            if count == 0:
                self.ShowMessage("Error", "Search Exception")
            elif count == self.MailListBox.size():
                self.Keyword = self.KeywordSearchInput.get()
                if self.Keyword in ["", " ", None]:
                    self.ShowMessage("Error", "Search Keyword is Invalid")
                else:
                    self.FolderSelected()
            else:
                self.ShowMessage("Error", "Search Exception")
        self.Keyword = None

    def DownloadAttachment(self):
        index = self.AttachmentListBox.curselection()[0]
        attachment = self.AttachmentList[index]
        filename = self.Mail.id.split("/")[-1] + attachment.name
        with open(filename, "wb") as fw:
            fw.write(attachment.content)
        self.ShowMessage("Info" ,"Attachment %s Downloaded Successfully, saved as:%s"%(str(attachment.name), filename))

    def FolderSelected(self):
        self.MailListBox.delete(0, tk.END)
        index = self.FolderListBox.curselection()[0]
        folder = str(self.Folder[str(index)])
        target = self.EmailObject.root.glob(folder+"*")
        if target.folders == []:
            target = self.EmailObject.root.glob("*/" + folder)
            if target.folders == []:
                target = self.EmailObject.root.glob("**/" + folder)
                if target.folders == []:
                    self.ShowMessage("Info", "Folder not Found")
                    return
        if self.Keyword in ["", " ", None]:
            self.PageSelect = target.all().order_by("-datetime_received")
            self.Page = self.PageSelect.iterator()
            for i in range(self.PageSize):
                email = next(self.Page)
                self.MailList[i] = email
                try:
                    banner = str(email.sender.name) + ":" + str(email.subject) + ":" + str(email.id)
                except Exception as exception:
                    continue
                try:
                    self.MailListBox.insert(tk.END, banner)
                except Exception as exception:
                    continue
        else:
            self.MailList = {}
            index = 0
            _list = target.all()
            alltext = _list.filter(text_body__contains=self.Keyword)
            self.PageSelect = alltext.all().order_by("-datetime_received")
            self.Page = self.PageSelect.iterator()
            for i in range(self.PageSize):
                email = next(self.Page)
                try:
                    banner = str(email.sender.name) + ":" + str(email.subject) + ":" + str(email.id)
                except Exception as exception:
                    continue
                try:
                    self.MailListBox.insert(tk.END, banner)
                except Exception as exception:
                    continue
                self.MailList[index] = email
                index += 1

    def FolderSelectedNext(self):
        self.MailListBox.delete(0, tk.END)
        self.MailList = {}
        index = 0
        if self.Page == None:
            return
        for i in range(self.PageSize):
            email = next(self.Page)
            try:
                banner = str(email.sender.name) + ":" + str(email.subject) + ":" + str(email.id)
            except Exception as exception:
                continue
            try:
                self.MailListBox.insert(tk.END, banner)
            except Exception as exception:
                continue
            self.MailList[index] = email
            index += 1

    def MailSelected(self):
        self.MailText.config(state="normal")
        self.MailText.delete(1.0, tk.END)
        self.MailText.config(state="disable")
        self.AttachmentListBox.delete(0, tk.END)
        index = self.MailListBox.curselection()[0]
        self.Mail = self.MailList[index]
        ccstring = ""
        bccstring = ""
        if self.Mail.cc_recipients != None:
            for person in self.Mail.cc_recipients:
                ccstring += " %s(%s);" % (str(person.name), str(person.email_address))
        if self.Mail.bcc_recipients != None:
            for person in self.Mail.bcc_recipients:
                bccstring += " %s(%s);" % (str(person.name), str(person.email_address))
        string = EmailFormatString%(str(self.Mail.id), str(self.Mail.sender.name), str(self.Mail.sender.email_address), ccstring, bccstring, str(self.Mail.datetime_received), str(self.Mail.subject), str(self.Mail.text_body))
        self.MailText.config(state="normal")
        self.MailText.insert(tk.END, string)
        self.MailText.config(state="disable")
        self.AttachmentList = self.Mail.attachments
        for attachment in self.AttachmentList:
            if isinstance(attachment, FileAttachment):
                fileindex = attachment.name
                self.AttachmentListBox.insert(tk.END, fileindex)



if __name__ == '__main__':
    mailcollection = MailToolGuiWindows()
    mailcollection.Graph()
