# excel-converter
This is my first serious project with more than 1000 line of codes.\
I wrote this project for my university's website in the second year of study.\
At the beginning of the project, I didn't know anything about Java, PostgreSQL and SQL.

## About the program
The essence of the program is very simple:
1) The program takes some data from .xls files
2) It saves them in PostgreSQL database

## Usage
1) Clear tables in PostgreSQL database<br>
<code>java -jar Converter.jar clear</code>
2) Add data to PostgreSQL database from all .xls files<br>
<code>java -jar Converter.jar update_all year</code>
3) Add data to PostgreSQL database from one specific .xls file<br>
<code>java -jar Converter.jar update fileName year</code>
<br>
<code>fileName</code> - name of an .xls file<br>
<code>year</code> - what's the year of the .xls file (the specifics of my university)
<br>
<br>

- Before using the program, you need to make sure that the data located in the <em>database.properties</em> file is valid.
- In order to parse several files at the same time, you MUST move the files to a folder called <em>"Файлы для парсинга" (Files for parsing).</em>

## What have I learned
- Java Basics
- PostgreSQL
- SQL

## License
(c) 2021 Vyacheslav. [MIT License](LICENSE)