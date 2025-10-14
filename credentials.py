from urllib.parse import quote_plus
from sqlalchemy import create_engine


# The credentials for the remote DB connection
username = "root"
password = quote_plus("shans1")
host = "30.1.1.4"
port = 3306
database = "cam"

# Dictionary of the department replacements
department_map = {
    "Ічня" : "ВП № 2 (Ічня) Прилуцького РВП ГУНП в Чернігівській області",
    "Бахмач" : "ВП № 1 (Бахмач) Ніжинського РУП ГУНП в Чернігівській області",
    "Бобровиця" : "ВП № 2 (Бобровиця) Ніжинського РУП ГУНП в Чернігівській області",
    "Борзна" : "ВП № 3 (Борзна) Ніжинського РУП ГУНП в Чернігівській області",
    "ВКА" : "ВКА ГУНП",
    "Городня" : "ВП № 2 (Городня) Чернігівського РУП ГУНП в Чернігівській області",
    "ДБР" : "ДБР м. Київа",
    "Козелець" : "ВП № 1 (Козелець) Чернігівського РУП ГУНП в Чернігівській області",
    "Короп" : "ВП № 1 (Короп) Новгород-Сіверського РВП ГУНП в Чернігівській області",
    "Корюківка" : "Корюківський РВП ГУНП в Чернігівській області",
    "Мена" : "ВП № 1 (Мена) Корюківського РВП ГУНП в Чернігівській області",
    "МунВарта" : "Муніципальна варта",
    "Новгород-Сіверський" : "Новгород-Сіверський РВП ГУНП в Чернігівській області",
    "Носівка" : "ВП № 4 (Носівка) Ніжинського РУП ГУНП в Чернігівській області",
    "Ніжин" : "Ніжинське РУП ГУНП в Чернігівській області",
    "Прилуки" : "Прилуцький РВП ГУНП в Чернігівській області",
    "Ріпки" : "ВП № 3 (Ріпки) Чернігівсього РУП ГУНП в Чернігівській області",
    "СУ" : "СУ ГУНП",
    "Сновськ" : "ВП № 2 (Сновськ) Корюківського РВП ГУНП в Чернігівській області",
    "УІАП" : "УІАП ГУНП",
    "УБН" : "УБН ГУНП",
    "УКР" : "УКР ГУНП",
    "УГІ" : "УГІ ГУНП",
    "УОАЗОР" : "УОАЗОР ГУНП",
    "УОС" : "УОС ГУНП",
    "УОТЗ" : "УОТЗ ГУНП",
    "УПД" : "УПД ГУНП",
    "УПО" : "Поліція охорони",
    "УПП" : "Патрульна поліція",
    "DCH_UPP": "Патрульна поліція",
    "УСР_ДСР" : "УСР ДСР НПУ",
    "ЧРУП" : "Чернігівське районне управління поліції",
    "ВМП" : "ВМП ГУНП"
}

# Dictionary of the table main columns rename
columns_rename = {
    "department" : (
                "Найменування структурного підрозділу "
                "територіального (міжрегіонального) органу "
                "поліції та або інших органів виконавчої (місцевої) влади"
            ),
    "count" : "Кількість АРМ",
    "username" : "Назва облікового запису",
    "security_level" : "Рівень доступу (звичайний або розширений)",
    "name" : "ПІБ користувача, який отримав доступ"
}

# The request to the remote DB to take the complex info about users
request = (
    "SELECT "
        "cu.title, "
        "c.username, "
        "c.name "
    "FROM ( "
        "SELECT " 
            "user_id, "
            "MAX(group_id) AS group_id " 
        "FROM cam_user_usergroup_map "
        "WHERE (group_id BETWEEN 44 AND 67) OR (group_id BETWEEN 70 AND 109) "
        "GROUP BY user_id "
    ") AS cuum "
    "JOIN ( "
            "SELECT "
                "id, "
                "title,"
                "report_order "
            "FROM cam_usergroups "
            "WHERE report_order IS NOT NULL "
    ") cu ON cu.id = cuum.group_id "            
    "JOIN ( "
	    "SELECT "
		    "usr.id, "
		    "usr.username, "
		    "usr.name "
	    "FROM cam_users usr "
	    "WHERE usr.block = 0 "
    ") " 
    "AS c on c.id = cuum.user_id "
    "ORDER BY cu.report_order, c.username "
    )

# Make a connection to the DB
def connection():
    return create_engine(
        f"mysql+pymysql://{username}:{password}@{host}:{port}/{database}"
    )
