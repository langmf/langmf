Для корректной работы необходимо прописать в файл "httpd.conf" следующие записи:


RewriteEngine on 
RewriteCond %{HTTP:Authorization} ^(.*) 
RewriteRule ^(.*) - [E=HTTP_AUTHORIZATION:%1]


И если следующая запись была такого вида: 

#LoadModule rewrite_module modules/mod_rewrite.so

То изменить ее на такую: 

LoadModule rewrite_module modules/mod_rewrite.so

И в конце перезагрузить веб-сервер Apache.

