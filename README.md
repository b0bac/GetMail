# GetMail
利用NTLM Hash读取Exchange邮件：在进行内网渗透时候，我们经常拿到的是账号的Hash凭据而不是明文口令。在这种情况下采用邮件客户端或者WEBMAIL的方式读取邮件就很麻烦，需要进行破解，NTLM的破解主要依靠字典强度，破解概率并不是很大。其实Exchange提供了利用NTLM Hash凭据进行验证的方法，从而可以进行任何操作。本程序支持邮箱目录结构的列举、邮件的阅读、附件的下载、关键字的搜索等功能，支持明文密码和NTLM Hash凭证两种认证方式。
## 参数介绍
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

## 使用
### 列举邮箱目录结构
```bash
python3 getmail.py -u [USERNAME] -p [PASSWORD] -e [EMAIL ADDRESS] -L
python3 getmail.py -u [USERNAME] -H [NTLMHASH] -e [EMAIL ADDRESS] -L
```
### 阅读邮件
```bash
python3 getmail.py -u [USERNAME] -H [NTLMHASH] -e [EMAIL ADDRESS] -f [文件夹，默认是Inbox] -c 阅读邮件数量（按照时间倒序，最近的在最前面）
python3 getmail.py -u [USERNAME] -H [NTLMHASH] -e [EMAIL ADDRESS] -f [文件夹，默认是Inbox] -c 阅读邮件数量（按照时间倒序，最近的在最前面） -k [Keyword]（展示包含关键字的邮件）
```
### 下载附件
```bash
python3 getmail.py -u [USERNAME] -H [NTLMHASH] -e [EMAIL ADDRESS] -f [文件夹，默认是Inbox] -c 阅读邮件数量（按照时间倒序，最近的在最前面）——D
```
## ISSUE修复记录
+ 0x01 默认Inbox收件箱邮件阅读超过20封时候有会问题【已修复】
+ 0x02 同名附件自动下载导致附件覆盖问题【已修复】

## Changelog
+ v1.1 20210421
  + 支持展示抄送、密送
  + 优化展示效果
  + 修复一些bug   
+ v1.1 20210422 增加GUI版本
  + Mac测试通过 
![image](https://user-images.githubusercontent.com/11972644/115653616-0e7b0d80-a362-11eb-816d-04fc6068bf6f.png)

+ v1.1 20210521 增加GUI版本分页获取邮件功能，优化相关体验（MacOS测试通过）  
  
![119072887-de855f80-ba1e-11eb-8334-bdea960ab785](https://github.com/b0bac/GetMail/assets/11972644/c73c94fb-49aa-4497-8f80-1496d2063934)

