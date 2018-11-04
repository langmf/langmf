Работа с CGI скриптами на основе движка LangMF.

------------------------------------------------------------------------------
Для корректной работы под сервером Apache необходимо:

 1) Установить Apache сервер:

    https://www.apachehaus.com/cgi-bin/download.plx

 2) прописать в файл "httpd.conf" следующую запись:

    # LangMF Settings

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


------------------------------------------------------------------------------
Для корректной работы под сервером lighttpd необходимо:

 1) Установить lighttpd сервер:

    http://lighttpd.dtech.hu

 2) прописать в файл "lighttpd.conf" следующую запись:

    # LangMF Settings
    server.modules = ("mod_access", "mod_accesslog", "mod_cgi")
    static-file.exclude-extensions = ( ".mf" )
    cgi.assign = (".mf" => "c:/Program Files/LangMF/LangMF.exe" )


------------------------------------------------------------------------------
При работе движка в cgi режиме не осуществляется показ форм и диалоговых окон. Если хотите активировать 
такую возможность то зайдите в "Пуск -> Выполнить" затем наберите там "services.msc" и нажмите "OK" ,затем 
в открывшемся окне идем по списку до сервиса с названием "Apache2.4" или "lighttpd", нажав на ней правой кнопкой мыши,
нажимаем на "Свойства" затем кликаем на вкладке "Вход в систему" и там ставим галочку 
"Разрешить взаимодействие с рабочим столом". Все осталось только перезапустить сервер.
