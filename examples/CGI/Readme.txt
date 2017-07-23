Работа с CGI скриптами на основе движка LangMF, тестировалась под Apache сервером (WIN платформа).

Для корректной работы необходимо:
 
 1) Установить Apache сервер:
 
	http://www.sai.msu.su/apache/httpd/binaries/win32/httpd-2.2.25-win32-x86-no_ssl.msi
    http://www.apachehaus.com/cgi-bin/download.plx
 
 2) прописать в файл "httpd.conf" следующую запись:
  	
    # ALL LangMF Settings for Apache
    
    ScriptAlias /LMF/ "C:/Program Files/LangMF/"
    AddType application/x-httpd-mf .mf
    AddHandler mf-script .mf
    Action mf-script /LMF/LangMF.exe
    Action application/x-httpd-mf "/LMF/LangMF.exe"
    
    # for Apache v.2.4.x
    <Directory "C:/Program Files/LangMF/">
        Require all granted
    </Directory>
    
    # for Apache v.2.2.x
    <Directory "C:/Program Files/LangMF/">
        AllowOverride None
        Options None
        Order allow,deny
        Allow from all
    </Directory>
 
 3) установить LangMF в директорию "C:/Program Files/LangMF/".

------------------------------------------------------------------------------
При работе движка в cgi режиме не осуществляется показ форм и диалоговых окон. Если хотите активировать 
такую возможность то зайдите в "Пуск -> Выполнить" затем наберите там "services.msc" и нажмите "OK" ,затем 
в открывшемся окне идем по списку до сервиса с названием "Apache2.2", нажав на ней правой кнопкой мыши 
нажимаем на "Свойства" затем кликаем на вкладке "Вход в систему" и там ставим галочку 
"Разрешить взаимодействие с рабочим столом". Все осталось только перезапустить Apache.
