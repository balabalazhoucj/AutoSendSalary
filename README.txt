1.工资表.xls是工资模板，可以增加相应的同事到后面。但是现在不能
增加相应的列，因为现在的邮件地址是固定到一个地方的。如果邮件地址的
 列被更改就不能发送程序。
 
2.smtp_config.ini文件是相应的配置文件。
[smtp]

smtp_server = mail.retechcorp.com
  
smtp_port   = 25

smtp_from   = 你的邮箱地址  
  
#配置smtp登陆信息，用户名和密码
 
[user]
  
user_name = 你的邮箱地址  
passwd    = 你的邮箱密码  
  

#mail_content定制邮件内容，sample.txt文件里面是发送邮件的正文，可以更改这个文件的内容来更改相应邮件内容  
#mail_subject 是发送邮件的主题
  
[Info]
  
mail_content = sample.txt
  
mail_subject = 工资条

 

3.如果有什么问题请联系