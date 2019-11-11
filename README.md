# Задача
Написать консольное приложение, которое на выходе создает Excel файл с данными случайных людей.
Данные, которые необходимо сгенерировать:
1. Персональные данные - имя, фамилия, отчество, возраст, пол (М или Ж), дата рождения, место рождения (город);
2. Место проживания - шестизначный почтовый индекс, страна, область, город, улица, дом, квартира.
Требования:
1) В файле один лист с таблицей, в которой сгенерированы данные для n человек, где n - целое число, задается пользователем параметром командной строки или вводом, 1 <= n <= 30;
2) Все текстовые данные на русском языке;
3) Все имена, фамилии и другие значения должны быть адекватными, случайные наборы символов не допускаются;
4) Дату в файл записывать в формате "ДД-ММ-ГГГГ"
5) Имена, фамилии и отчества должны сочетаться с полом, например, женские имена с мужским отчеством не допускаются, как и мужские имена с женским полом;
6) Дата рождения и возраст тоже должны соответствовать друг другу;
7) После того, как файл создан, в лог должно быть выведено сообщение:
"Файл создан. Путь: *здесь выводим полный путь к файлу*".

# Итог
Консольная программа, на вход принимает стандартный ввод.
Добавлена обработка ошибок ввода.
Часть данных считывается из текстовых файлов и затем случайным образом выбирается конечный набор, часть генерируется внутри программы.
Файлы сохраняются в папке проекта: ...\DataTest\src\main\resources.
Вынесла, что смогла, в отдельные методы.
К сожалению, по SOLID не получилось, а сгенерить ещё и pdf не успела :(
