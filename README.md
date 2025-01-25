# Массовый импорт контактов из Excel

## Заказчик

Фамилия Имя Отчество - педагог дополнительного образования.

## Проблема

Каждый год педагоги дополнительного образования должны добавлять контакты учеников и родителей в свою телефонную книгу. 

В IT-Куб г.(изменено) есть единая база данных с номерами учеников и родителей в виде Excel файла, импорт контактов из нее затруднён, особенно на устройствах с IOS

## Цель работы

Разработать Desktop приложение для массового импорта контактов в телефонную книгу.

## Описание работы программы

Пользователь открывает приложение с таким интерфейсом:

![Снимок экрана 2023-10-22 072724](https://github.com/YanIlin22/XlsxDataGet/assets/119660072/fd3a311c-0f64-47a8-87ec-82606661f0d8)

Нажав на кнопку “Открыть” открывается диалоговое окно выбора файлов, необходимо открыть БД учеников IT-Куба.

Тестовая БД представлена в репозитории.

Из БД подгружаются списки групп, пользователю необходимо выбрать группу и затем нажать на кнопку “Ученики” или “Родители”.

При нажатии этих кнопок открывается диалоговое окно с выбором пути для сохранения файла.

Программа формирует файл контактов в формате VCF.

Затем этот файл можно перенести себе в телефон. Открыв его на телефоне, контакты будут импортированы в телефонную книгу.

## Разработчик

Ильин Ян.
