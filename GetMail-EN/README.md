# GetMail
GetMail was developed by b0bac to facilitate the need to utilize NTLM hash to perform pass-the-hash attack to obtain read e-mails from on-premise Exchange servers. The cracking of the NTLM hash is dependent on the complexity of the password and therefore GetMail provides a unique advantage weaponizing the pass-the-hash attack.

GetMail supports the enumeration of the mailbox directory structure, reading of mail, attachment downloads, keyword searches, and other functions. Two authentication methods are supported: password-based authentication, and NTLM hash.

## Help
+  -h, --help                     show this help message and exit
+  -u USER, --user=USER           Please Input Username!
+  -H HASH, --hash=HASH           Please Input Ntlmhash: xx:xx Or xxxx!
+  -p PSWD, --pswd=PSWD           Please Input Password!
+  -e EMAIL, --email=EMAIL        Please Input Email Address!
+  -c COUNT, --count=COUNT        Please Input How Many Emails You Want To Read!
+  -s SERVER, --server=SERVER     Please Input Email Server Address!
+  -k KEYWORD, --keyword=KEYWORD  Please Input Keyword To Search!
+  -L, --List                     List All Email Floders!
+  -D, --download                 Whether Download Attachment Files Or Not!
+  -d, --display                  Show All Email Floders!
+  -l, --list                     List All Email Floders!
+  -f FLODER, --folder=FLODER     Please Input Email Floder Name!

## Usage
### List Mailbox Directory Structure
```bash
python3 getmail.py -u [USERNAME] -p [PASSWORD] -e [EMAIL ADDRESS] -L
python3 getmail.py -u [USERNAME] -H [NTLMHASH] -e [EMAIL ADDRESS] -L
```
### Read Mail
```bash
python3 getmail.py -u [USERNAME] -H [NTLMHASH] -e [EMAIL ADDRESS] -f [文件夹，默认是Inbox] -c 阅读邮件数量（按照时间倒序，最近的在最前面）
python3 getmail.py -u [USERNAME] -H [NTLMHASH] -e [EMAIL ADDRESS] -f [文件夹，默认是Inbox] -c 阅读邮件数量（按照时间倒序，最近的在最前面） -k [Keyword]（展示包含关键字的邮件）
```
### Download Attachments
```bash
python3 getmail.py -u [USERNAME] -H [NTLMHASH] -e [EMAIL ADDRESS] -f [文件夹，默认是Inbox] -c 阅读邮件数量（按照时间倒序，最近的在最前面）——D
```
## Repaired Issues
+ 0x01 Problem with reading more than 20 e-mails in the default inbox [FIXED]
+ 0x02 Automatic download of attachments with same name causes an attachment overwrite issue [FIXED]

## Changelog
+ v1.1 20210421
  + Support displaying of CC and BCC information
  + Optimize the display
  + Fix additional minor bugs
+ v1.1 20210422 Addition of GUI
  + MacOS support successfully tested
![image](https://user-images.githubusercontent.com/11972644/115653616-0e7b0d80-a362-11eb-816d-04fc6068bf6f.png)

+ v1.1 20210521 Addition of GUI version paging to reading emails and optimized the experience (supports MacOS)
  
![image](https://user-images.githubusercontent.com/11972644/119072887-de855f80-ba1e-11eb-8334-bdea960ab785.png)
