import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellReference;

public class Database {
    static Connection connection = null;
    static Statement statement = null;
    static PreparedStatement ppstatement = null;
    static ResultSet result = null;

    //***************************//
    // Подключение к базе данных //
    //***************************//
    public static void connect(String DB_URL, String USER, String PASS) {
        System.out.println("Проверка подключения к PostgreSQL JDBC...");

        try {
            Class.forName("org.postgresql.Driver");
        } catch (ClassNotFoundException e) {
            System.out.println("PostgreSQL JDBC Driver не найден.");
            e.printStackTrace();
            return;
        }

        System.out.println("PostgreSQL JDBC Driver был успешно подключен!");

        try {
            connection = DriverManager.getConnection(DB_URL, USER, PASS);
        } catch (SQLException e) {
            System.out.println("Ошибка подключения.");
            e.printStackTrace();
            return;
        }

        if (connection != null)
            System.out.println("Вы успешно подключились к базе данных!");
        else
            System.out.println("Не удалось подключиться к базе данных.");
    }

    //*******************************************************************//
    // Проверка таблицы discipline на наличие всех необходимых дисциплин //
    //*******************************************************************//
    public static void checkDiscipline(String fileName) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("План");
        HSSFRow row;

        String checkDiscipline = "SELECT COUNT(id) AS count FROM discipline WHERE LOWER(REPLACE(name, ' ', ''))=?;";
        String addDiscipline = "INSERT INTO discipline(id, name) VALUES(?, ?);";
        String getLastId = "SELECT MAX(ID) FROM discipline;";

        System.out.println("\n\nПроверка таблицы discipline на наличие недостающих дисциплин");
        System.out.println("------------");

        try {
            int rowCount = 5;
            int disciplineCount = 0;

            // В переменную id добавляем значение, которое отсутствует в таблице discipline
            statement = connection.createStatement();
            result = statement.executeQuery(getLastId);
            result.next();
            int idDiscipline = result.getInt("max") + 1;

            while ((row = sheet.getRow(rowCount)) != null) {
                // Пропускаем ненужные строки и переходим к следующим
                if ((row.getCell(1).getStringCellValue()).equals("") ||
                    (row.getCell(2).getStringCellValue()).equals("") ||
                    (row.getCell(2).getStringCellValue()).contains("Дисциплины")) {
                    rowCount++;
                    continue;
                }

                // В переменную discipline добавляем название дисциплины
                String discipline = (row.getCell(2).getStringCellValue());
                ppstatement = connection.prepareStatement(checkDiscipline);
                ppstatement.setString(1, discipline.replaceAll("\\s*", "").toLowerCase());
                result = ppstatement.executeQuery();
                result.next();

                // Если дисциплина отсутствует - добавить в таблицу
                if (result.getInt("count") == 0) {
                    ppstatement = connection.prepareStatement(addDiscipline);
                    ppstatement.setInt(1, idDiscipline);
                    ppstatement.setString(2, discipline);
                    ppstatement.executeUpdate();

                    idDiscipline++;
                    disciplineCount++;
                    System.out.println(disciplineCount + ". В таблицу discipline была добавлена дисциплина: " + discipline);
                }

                rowCount++;
            }

            if (disciplineCount != 0) {
                System.out.println("------------");
                System.out.println("Добавление дисциплин прошло успешно!");
                System.out.println("Всего было добавлено недостающих дисциплин: " + disciplineCount);
            } else {
                System.out.println("------------");
                System.out.println("Недостающих дисциплин не обнаружено!");
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    //***********************************//
    // Заполнение таблицы podrazdelenies //
    //***********************************//
    public static void updatePodrazdelenies(String file) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet sheet = workbook.getSheet("ПланСвод");
        HSSFRow row;

        String sql = "INSERT INTO podrazdelenies(name, id_type_podrazdelenie) VALUES(?, ?);";
        String checkPodrazName = "SELECT id FROM podrazdelenies WHERE name=?;";

        System.out.println("\n\nТаблица podrazdelenie");
        System.out.println("------------");

        try {
            int rowCount = 5;
            int unitCount = 0;

            while ((row = sheet.getRow(rowCount)) != null) {
                // Пропускаем ненужные строки и переходим к следующим
                if ((row.getCell(0).getStringCellValue()).equals("") ||
                    (row.getCell(1).getStringCellValue()).equals("") ||
                    (row.getCell(2).getStringCellValue()).contains("Дисциплины")) {
                    rowCount++;
                    continue;
                }

                int lastCell = row.getLastCellNum() - 1; // Номер ячейки, в которой находятся нужные данные
                String unitName = row.getCell(lastCell).getStringCellValue(); // Название подразделения
                int idPodrazType = 1; // Тип подразделения | 1 - кафедра

                /* Если элемент с таким именем уже присутствует в таблице,
                то пропускаем его и переходим к следующему */
                ppstatement = connection.prepareStatement(checkPodrazName);
                ppstatement.setString(1, unitName);
                result = ppstatement.executeQuery();
                result.next();
                if (result.getRow() != 0)
                {
                    rowCount++;
                    continue;
                }

                // Отправляем все собранные данные в таблицу podrazdelenies
                ppstatement = connection.prepareStatement(sql);
                ppstatement.setString(1, unitName);
                ppstatement.setInt(2, idPodrazType);
                ppstatement.executeUpdate();

                unitCount++;
                rowCount++;
                System.out.println(unitCount + ". В таблицу podrazdelenies был добавлен элемент: " + unitName);
            }
            System.out.println("------------");
            System.out.println("Добавление элементов прошло успешно!");
            System.out.println("Всего было добавлено элементов: " + unitCount);
        } catch (SQLException e) {
            System.out.println("------------");
            System.out.println("Что-то пошло не так. Добавление элементов в таблицу podrazdelenies " +
                "не было завершено на 100%");
            e.printStackTrace();
        }
    }

    //****************************//
    // Заполнение таблицы profile //
    //****************************//
    public static void updateProfile(String fileName, String tableFormat) throws IOException, SQLException {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("Титул");
        HSSFRow row;

        String sql = "INSERT INTO profile(id, code, specialty, profile, id_specialtys, name) VALUES(?, ?, ?, ?, ?, ?);";
        String addSpecialtyInfo = "INSERT INTO specialtys(name, code) VALUES(?, ?);";
        String getSpecialtyInfo = "SELECT * FROM specialtys WHERE REPLACE(code, '.', '')=?;";
        String getNameProfile = "SELECT id FROM profile WHERE LOWER(REPLACE(name, ' ', ''))=?;";
        String getIdProfile = "SELECT id FROM profile ORDER BY id DESC LIMIT 1;";

        System.out.println("\n\nТаблица profiles");
        System.out.println("------------");

        int rowNum;
        int cellNum;

        // Определение формата таблицы
        if (tableFormat.equals("2018")) {
            rowNum = 16;
            cellNum = 2;
        } else {
            rowNum = 15;
            cellNum = 1;
        }

        // Определяем код и название специальности
        row = sheet.getRow(rowNum);
        String codeSpecialtys = row.getCell(cellNum).getStringCellValue();
        row = sheet.getRow(rowNum + 2);
        String fullName = row.getCell(cellNum).getStringCellValue();
        String[] fullNameSpecialty = fullName.split("\n");
        String nameSpecialtys = fullNameSpecialty[0].replaceAll("[^a-яА-Я ]", "");

        // Определяем, присутствует ли в таблице specialtys такая специальность
        int idSpecialtys;
        ppstatement = connection.prepareStatement(getSpecialtyInfo);
        ppstatement.setString(1, codeSpecialtys.replaceAll("[^0-9]", ""));
        result = ppstatement.executeQuery();
        result.next();
        if (result.getRow() == 0) {
            ppstatement = connection.prepareStatement(addSpecialtyInfo);
            ppstatement.setString(1, nameSpecialtys);
            ppstatement.setString(2, codeSpecialtys);
            ppstatement.executeUpdate();

            ppstatement = connection.prepareStatement(getSpecialtyInfo);
            ppstatement.setString(1, codeSpecialtys.replaceAll("[^0-9]", ""));
            result = ppstatement.executeQuery();
            result.next();
            idSpecialtys = result.getInt("id");
        } else {
            idSpecialtys = result.getInt("id");
        }

        // Если такой профиль уже есть в таблице - он не добавляется в таблицу profile
        row = sheet.getRow(rowNum + 3);
        String nameProfile = row.getCell(cellNum).getStringCellValue();
        ppstatement = connection.prepareStatement(getNameProfile);
        ppstatement.setString(1, nameProfile.toLowerCase(Locale.ROOT).replaceAll("[^а-яА-Я]", ""));
        result = ppstatement.executeQuery();
        result.next();
        if (result.getRow() != 0) {
            System.out.println("Новых профилей добавлено не было!");
            System.out.println("------------");
            return;
        }

        // Определяем id таблицы profile
        int idProfile;
        statement = connection.createStatement();
        result = statement.executeQuery(getIdProfile);
        result.next();
        if (result.getRow() == 0)
            idProfile = 1;
        else
            idProfile = result.getInt("id") + 1;

        // Отправляем все собранные данные в таблицу profile
        ppstatement = connection.prepareStatement(sql);
        ppstatement.setInt(1, idProfile);
        ppstatement.setString(2, codeSpecialtys);
        ppstatement.setString(3, nameSpecialtys);
        ppstatement.setString(4, nameProfile);
        ppstatement.setInt(5, idSpecialtys);
        ppstatement.setString(6, nameProfile);
        ppstatement.executeUpdate();

        System.out.println("------------");
        System.out.println("В таблицу был добавлен профиль: " + nameProfile);
        System.out.println("Добавление элементов прошло успешно!");
    }

    //**********************************//
    // Заполнение таблицы module_choose //
    //**********************************//
    public static void updateModuleChoose(String file) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet sheet = workbook.getSheet("План");
        HSSFRow row;

        String sql = "INSERT INTO module_choose(id_type, name) VALUES(?, ?);";
        String getTypeId = "SELECT id FROM type WHERE value=?;";

        System.out.println("\n\nТаблица module_choose");
        System.out.println("------------");

        try {
            int rowCount = 5;
            int moduleCount = 0;
            int disciplineCount = 0;
            boolean physicalCulture = false;
            int idType;
            String name = "";

            while ((row = sheet.getRow(rowCount)) != null) {
                // Если ячейка в Excel содержит слово Блок 2, значит, завершаем парсинг данной таблицы
                if ((row.getCell(0).getStringCellValue()).contains("Блок 2"))
                    break;

                // Пропускаем ненужные строки и переходим к следующим
                if ((row.getCell(0).getStringCellValue()).equals("") ||
                    (row.getCell(1).getStringCellValue()).equals("")) {
                    rowCount++;
                    continue;
                }

                if ((row.getCell(2).getStringCellValue()).contains("Дисциплины") ||
                    row.getCell(2).getStringCellValue().contains("элективные дисциплины")) {
                    if (row.getCell(2).getStringCellValue().contains("элективные дисциплины")) {
                        physicalCulture = true;
                    } else {
                        physicalCulture = false;
                    }

                    moduleCount = 0;
                    rowCount++;
                    name = row.getCell(2).getStringCellValue();
                    continue;
                }

                if (name.equals("")) {
                    rowCount++;
                    continue;
                }

                // Название дисциплины а также номер модуля по выбору
                String discipline = (row.getCell(2).getStringCellValue());
                String index = row.getCell(1).getStringCellValue();
                int fullType = index.indexOf('.');
                String firstIndex = index.substring(0, fullType);

                // Тип модуля по выбору
                ppstatement = connection.prepareStatement(getTypeId);
                ppstatement.setString(1, firstIndex);
                result = ppstatement.executeQuery();
                result.next();
                idType = result.getInt("id");

                // Отправляем все собранные данные в таблицу module_choose
                ppstatement = connection.prepareStatement(sql);
                ppstatement.setInt(1, idType);
                ppstatement.setString(2, name);
                ppstatement.executeUpdate();

                moduleCount++;
                if (moduleCount / 2 == 1 && !physicalCulture) name = "";
                disciplineCount++;
                rowCount++;

                System.out.println(disciplineCount + ". В таблицу была добавлена дисциплина: " + discipline);
            }

            System.out.println("------------");
            System.out.println("Добавление дисциплин прошло успешно!");
            System.out.println("Всего было добавлено дисциплин: " + disciplineCount);
        } catch (SQLException e) {
            System.out.println("------------");
            System.out.println("Что-то пошло не так. Добавление дисциплин в таблицу module_choose " +
                "не было завершено на 100%");
            e.printStackTrace();
        }
    }

    //*******************************//
    // Заполнение таблицы teach_plan //
    //*******************************//
    public static void updateTeachPlan(String fileName, String tableFormat) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workbook.getSheet("Титул");
        HSSFRow row;

        String sql = "INSERT INTO teach_plan(id, course, id_profile, id_form, date_start, date_end) VALUES(?, ?, ?, ?, ?, ?);";
        String getTeachPlanId = "SELECT id FROM teach_plan ORDER BY id DESC LIMIT 1;";
        String getFormId = "SELECT id FROM form_of_training WHERE name=?;";
        String getProfileId = "SELECT id FROM profile WHERE LOWER(REPLACE(name, ' ', ''))=? ORDER BY id DESC LIMIT 1;";

        System.out.println("\n\nТаблица teach_plan");
        System.out.println("------------");

        try {
            /* Если таблица teach_plan пуста, присваиваем переменной id значение 1,
            иначе берём последнее значение id из таблицы и инкрементируем его */
            int id;
            statement = connection.createStatement();
            result = statement.executeQuery(getTeachPlanId);
            result.next();
            if (result.getRow() == 0)
                id = 1;
            else
                id = result.getInt("id") + 1;

            // В зависимости от формата таблицы, программа будет искать искать данные на нужной строке
            String academicYear = ""; // Год начала и год конца определенного курса
            int fistStudyYear = 0; // Год начала 4-ёх или 6-ти годичного обучения
            String trainingForm = "";

            if (tableFormat.equals("2018")) {
                row = sheet.getRow(29);
                fistStudyYear = Integer.parseInt(row.getCell(20).getStringCellValue());
                row = sheet.getRow(30);
                academicYear = row.getCell(18).getStringCellValue();
                row = sheet.getRow(31);
                trainingForm = row.getCell(1).getStringCellValue();
            } else if (tableFormat.equals("2019") || tableFormat.equals("2020")) {
                row = sheet.getRow(28);
                fistStudyYear = Integer.parseInt(row.getCell(19).getStringCellValue());
                row = sheet.getRow(29);
                academicYear = row.getCell(17).getStringCellValue();
                row = sheet.getRow(30);
                trainingForm = row.getCell(0).getStringCellValue();
            } else if (tableFormat.equals("2021")) {
                row = sheet.getRow(28);
                fistStudyYear = Integer.parseInt(row.getCell(19).getStringCellValue());
                row = sheet.getRow(29);
                academicYear = row.getCell(19).getStringCellValue();
                row = sheet.getRow(30);
                trainingForm = row.getCell(0).getStringCellValue();
            }

            int fullType = academicYear.indexOf('-');
            String date = academicYear.substring(0, fullType);

            int yearStart = Integer.parseInt(date); // Год начала определенного курса
            int monthStart = 9;
            int dayStart = 1;
            int monthEnd = 8;
            int dayEnd = 31;

            // Столбик date_start
            String dateStart = (yearStart) + "-" + monthStart + "-" + dayStart;
            SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
            java.util.Date utilDate = format.parse(dateStart);
            java.sql.Date sqlDateStart = new java.sql.Date(utilDate.getTime());

            // Столбик course
            int course = yearStart - fistStudyYear + 1;

            // Столбик date_end
            String dateEnd = (yearStart + 1) + "-" + monthEnd + "-" + dayEnd;
            format = new SimpleDateFormat("yyyy-MM-dd");
            utilDate = format.parse(dateEnd);
            java.sql.Date sqlDateEnd = new java.sql.Date(utilDate.getTime());

            // Столбик id_profile
            int idProfile;
            row = sheet.getRow(18);
            String nameProfile = row.getCell(1).getStringCellValue();
            ppstatement = connection.prepareStatement(getProfileId);
            ppstatement.setString(1, nameProfile.toLowerCase(Locale.ROOT).replaceAll("[^а-яА-Я]", ""));
            result = ppstatement.executeQuery();
            result.next();
            if (result.getRow() != 0)
                idProfile = result.getInt("id");
            else
                idProfile = 0;

            // Столбик id_form
            String[] formName = trainingForm.split((" "));
            String fullFormName;

            if (trainingForm.contains("сокращенная"))
                fullFormName = formName[2] + " сокращенная";
            else if (trainingForm.contains("ускоренная"))
                fullFormName = formName[2] + " ускоренная";
            else
                fullFormName = formName[2];

            ppstatement = connection.prepareStatement(getFormId);
            ppstatement.setString(1, fullFormName);
            result = ppstatement.executeQuery();
            result.next();
            int idForm = result.getInt("id");

            // Отправляем все собранные данные в таблицу teach_plan
            ppstatement = connection.prepareStatement(sql);
            ppstatement.setInt(1, id);
            ppstatement.setInt(2, course);
            ppstatement.setInt(3, idProfile);
            ppstatement.setInt(4, idForm);
            ppstatement.setDate(5, sqlDateStart);
            ppstatement.setDate(6, sqlDateEnd);
            ppstatement.executeUpdate();

            System.out.println("------------");
            System.out.println("Добавление элементов прошло успешно!");
            System.out.println("В таблицу был добавлен учебный год: " + academicYear);
        } catch (SQLException | ParseException | NumberFormatException e) {
            System.out.println("------------");
            System.out.println("Что-то пошло не так. Добавление элементов в таблицу teach_plan " +
                "не было завершено на 100%.");
            System.out.println("Возможно, формат таблицы является не подходящим.");
            System.out.println();
            e.printStackTrace();
        }
    }

    //*************************************//
    // Заполнение таблицы grafik_education //
    //*************************************//
    public static void updateGrafikEducation() {
        String sql = "INSERT INTO grafik_education(id_teach_plan, year_start) VALUES(?, ?);";
        String getTeachPlanId = "SELECT id FROM teach_plan ORDER BY id DESC LIMIT 1;";
        String getTeachPlanYear = "SELECT date_start FROM teach_plan WHERE id=?;";

        System.out.println("\n\nТаблица grafik_education");
        System.out.println("------------");

        try {
            // Определяем последний id в таблице teach_plan
            statement = connection.createStatement();
            result = statement.executeQuery(getTeachPlanId);
            result.next();
            int idTeachPlan = result.getInt("id");

            // Определяем год начала курса в таблице teach_plan
            ppstatement = connection.prepareStatement(getTeachPlanYear);
            ppstatement.setInt(1, idTeachPlan);
            result = ppstatement.executeQuery();

            while (result.next()) {
                String pattern = "yyyy-MM-dd";
                DateFormat df = new SimpleDateFormat(pattern);
                Date year = result.getDate("date_start");
                String yearStart = df.format(year);
                int fullType = yearStart.indexOf('-');
                String date = yearStart.substring(0, fullType);
                int dateStart = Integer.parseInt(date);

                ppstatement = connection.prepareStatement(sql);
                ppstatement.setInt(1, idTeachPlan);
                ppstatement.setInt(2, dateStart);
                ppstatement.executeUpdate();

                System.out.println("------------");
                System.out.println("Добавление элементов прошло успешно!");
                System.out.println("В таблицу был добавлен год начала курса: " + dateStart);
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    //******************************************//
    // Заполнение таблицы grafik_education_days //
    //******************************************//
    public static void updateGrafikEducationDays(String file, String tableFormat) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet sheet = workbook.getSheet("График");
        HSSFRow row;
        HSSFCell cell;

        String sql = "INSERT INTO grafik_education_days(id_grafik_education, day, mounth, year, id_vid_activ) " +
            "VALUES(?, ?, ?, ?, ?);";
        String getGrafikEduID = "SELECT id FROM grafik_education ORDER BY id DESC LIMIT 1;";
        String getGrafikEduIdTP = "SELECT id_teach_plan FROM grafik_education WHERE id=?;";
        String getGrafikEduYear = "SELECT year_start FROM grafik_education WHERE id_teach_plan=?;";
        String getTeachPlanCourse = "SELECT course FROM teach_plan WHERE id=?;";

        System.out.println("\n\nТаблица grafik_education_days");
        System.out.println("------------");

        try {
            // Определяем последний id в таблице grafik_education
            statement = connection.createStatement();
            result = statement.executeQuery(getGrafikEduID);
            result.next();
            int idGrafikEduId = result.getInt("id");

            // Определяем нужный id_teach_plan в таблице grafik_education
            ppstatement = connection.prepareStatement(getGrafikEduIdTP);
            ppstatement.setInt(1, idGrafikEduId);
            result = ppstatement.executeQuery();
            result.next();
            int idTeachPlan = result.getInt("id_teach_plan");

            // Определяем нужный учебный год в таблице grafik_education
            ppstatement = connection.prepareStatement(getGrafikEduYear);
            ppstatement.setInt(1, idTeachPlan);
            result = ppstatement.executeQuery();
            result.next();
            int yearStart = result.getInt("year_start");

            // Определяем номер курса в таблице teach_plan
            ppstatement = connection.prepareStatement(getTeachPlanCourse);
            ppstatement.setInt(1, idTeachPlan);
            result = ppstatement.executeQuery();
            result.next();
            int course = result.getInt("course") - 1;

            int monthCount = 1; // 1 - Сентябрь | 12 - Август
            int weekCount = 0;
            int lastWeek = 0;

            // В зависимости от типа Excel таблицы выбираем необходимые значения переменных weekCount и lastWeek
            if (tableFormat.equals("2019") || tableFormat.equals("2020") || tableFormat.equals("2021")) {
                weekCount = 1;
                lastWeek = 53;
            } else if (tableFormat.equals("2018")) {
                weekCount = 2;
                lastWeek = 54;
            }

            while (weekCount < lastWeek) {
                ppstatement = connection.prepareStatement(sql);

                boolean newMonth = false;
                int tableRow = 12 + (7 * course);
                int firstDay;
                int lastDay;
                int whichMonth = 0;

                row = sheet.getRow(2); // 2
                String[] days;
                String weekDays;
                weekDays = row.getCell(weekCount).getStringCellValue();

                if (weekDays.equals("")) {
                    row = sheet.getRow(1);
                    weekDays = row.getCell(weekCount).getStringCellValue();
                }
                days = weekDays.split(" ");
                firstDay = Integer.parseInt(days[0]);
                lastDay = Integer.parseInt(days[days.length - 1]);

                // Если первый день недели начинается в конце одного месяца, а последний день недели в начале следующего месяца
                if (firstDay > lastDay) {
                    if (firstDay + 6 - 30 == lastDay)
                        whichMonth = 30;
                    else if (firstDay + 6 - 31 == lastDay)
                        whichMonth = 31;
                    else
                        whichMonth = 28;
                }

                boolean cellsMerged = false;
                int vidActiv = 0;
                row = sheet.getRow(12 + (course * 7));

                while (tableRow < 18 + (course * 7)) {
                    row = sheet.getRow(tableRow);

                    if (firstDay > lastDay || newMonth) {
                        ppstatement.setInt(3, monthCount * 10 + (monthCount + 1));
                        newMonth = true;
                    } else {
                        ppstatement.setInt(3, monthCount);
                    }

                    ppstatement.setInt(1, idGrafikEduId);
                    ppstatement.setInt(2, firstDay);
                    ppstatement.setInt(4, yearStart);

                    String activ = row.getCell(weekCount).getStringCellValue();

                    if (!cellsMerged) {
                        switch (activ) {
                            case "":
                                vidActiv = 1;
                                break;
                            case "Э":
                                vidActiv = 2;
                                break;
                            case "У":
                                vidActiv = 3;
                                break;
                            case "П":
                                vidActiv = 4;
                                break;
                            case "Д":
                                vidActiv = 5;
                                break;
                            case "К":
                                vidActiv = 6;
                                break;
                            case "*":
                                vidActiv = 7;
                                break;
                            default:
                                break;
                        }
                    }

                    cell = row.getCell(weekCount);
                    int rowIndex = cell.getRowIndex();
                    int tempColumnIndex = cell.getColumnIndex();
                    String columnIndex = CellReference.convertNumToColString(tempColumnIndex);
                    String firstFullIndex = columnIndex + "" + rowIndex;
                    String secondFullIndex = columnIndex + "" + (rowIndex + 5);

                    for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                        String firstIndex = CellReference.convertNumToColString(sheet.getMergedRegion(i).getFirstColumn()) + sheet.getMergedRegion(i).getFirstRow();
                        String secondIndex = CellReference.convertNumToColString(sheet.getMergedRegion(i).getLastColumn()) + sheet.getMergedRegion(i).getLastRow();

                        if (firstIndex.equals(firstFullIndex) && secondIndex.equals(secondFullIndex))
                            cellsMerged = true;
                    }

                    if (firstDay == whichMonth && monthCount != 3)
                        firstDay = 0;

                    if (firstDay == whichMonth && monthCount == 3)
                        newMonth = true;

                    ppstatement.setInt(5, vidActiv);
                    ppstatement.executeUpdate();

                    firstDay++;
                    tableRow++;
                }

                if (!newMonth && lastDay == 30 || lastDay == 31) monthCount++;
                if (newMonth) monthCount++;

                weekCount++;
            }

            System.out.println("------------");
            System.out.println("Добавление элементов прошло успешно!");
        } catch (SQLException e) {
            System.out.println("------------");
            System.out.println("Что-то пошло не так. Добавление элементов в таблицу grafik_education " +
                "не было завершено на 100%");
            e.printStackTrace();
        }
    }

    //*******************************************//
    // Заполнение таблицы grafik_education_days //
    //******************************************//
    public static void updateGrafikEducationDaysNew(String file, String tableFormat) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet sheet = workbook.getSheet("График");
        HSSFRow row;
        HSSFCell cell;

        String sql = "INSERT INTO grafik_education_days(id_grafik_education, day, mounth, year, id_vid_activ) " +
            "VALUES(?, ?, ?, ?, ?);";
        String getGrafikEduID = "SELECT id FROM grafik_education ORDER BY id DESC LIMIT 1;";
        String getGrafikEduIdTP = "SELECT id_teach_plan FROM grafik_education WHERE id=?;";
        String getGrafikEduYear = "SELECT year_start FROM grafik_education WHERE id_teach_plan=?;";
        String getTeachPlanCourse = "SELECT course FROM teach_plan WHERE id=?;";

        System.out.println("\n\nТаблица grafik_education_days");
        System.out.println("------------");

        try {
            // Определяем последний id в таблице grafik_education
            statement = connection.createStatement();
            result = statement.executeQuery(getGrafikEduID);
            result.next();
            int idGrafikEduId = result.getInt("id");

            // Определяем нужный id_teach_plan в таблице grafik_education
            ppstatement = connection.prepareStatement(getGrafikEduIdTP);
            ppstatement.setInt(1, idGrafikEduId);
            result = ppstatement.executeQuery();
            result.next();
            int idTeachPlan = result.getInt("id_teach_plan");

            // Определяем нужный учебный год в таблице grafik_education
            ppstatement = connection.prepareStatement(getGrafikEduYear);
            ppstatement.setInt(1, idTeachPlan);
            result = ppstatement.executeQuery();
            result.next();
            int yearStart = result.getInt("year_start");

            // Определяем номер курса в таблице teach_plan
            ppstatement = connection.prepareStatement(getTeachPlanCourse);
            ppstatement.setInt(1, idTeachPlan);
            result = ppstatement.executeQuery();
            result.next();
            int course = result.getInt("course") - 1;

            int monthCount = 0; // Столбик mounth
            int weekCount = 1;
            int lastWeek = 53;
            int daysCount = 1;

            String[] months = {"Сентябрь", "Октябрь", "Ноябрь", "Декабрь", "Январь", "Февраль",
                "Март", "Апрель", "Май", "Июнь", "Июль", "Август"};

            while (weekCount <= lastWeek) {
                row = sheet.getRow((course * 17) + 2);
                String month = row.getCell(weekCount).getStringCellValue();

                for (int i = 0; i < 12; i++) {
                    if (month.equals(months[i]))
                        monthCount = i + 1;
                }

                int firstWeekDay = 0;
                int fd = (course * 17) + 3;
                row = sheet.getRow(fd);
                while (row.getCell(weekCount).getStringCellValue().equals("")) {
                    fd++;
                    firstWeekDay++;
                    row = sheet.getRow(fd);
                }
                String firstDay = row.getCell(weekCount).getStringCellValue();

                int ld = (course * 17) + 8;
                row = sheet.getRow(ld);
                while (row.getCell(weekCount).getStringCellValue().equals("")) {
                    ld--;
                    row = sheet.getRow(ld);
                }
                String lastDay = row.getCell(weekCount).getStringCellValue();

                int firstDayDate = Integer.parseInt(firstDay);
                int lastDayDate = Integer.parseInt(lastDay);

                if (firstDayDate > lastDayDate) {
                    if (monthCount < 9)
                        monthCount = monthCount * 10 + (monthCount + 1);
                    else
                        monthCount = monthCount * 100 + (monthCount + 1);
                }

                int tableRow = 12 + firstWeekDay + (course * 17);
                int vidActiv = 0; // Столбик id_vid_activ
                boolean cellsMerged = false;

                while (tableRow < (18 + (course * 17))) {
                    ppstatement = connection.prepareStatement(sql);
                    row = sheet.getRow(tableRow  - 9);
                    daysCount = Integer.parseInt(row.getCell(weekCount).getStringCellValue());
                    row = sheet.getRow(tableRow);
                    String activ = row.getCell(weekCount).getStringCellValue();

                    if (!cellsMerged) {
                        switch (activ) {
                            case "":
                                vidActiv = 1;
                                break;
                            case "Э":
                                vidActiv = 2;
                                break;
                            case "У":
                                vidActiv = 3;
                                break;
                            case "П":
                                vidActiv = 4;
                                break;
                            case "Д":
                                vidActiv = 5;
                                break;
                            case "К":
                                vidActiv = 6;
                                break;
                            case "*":
                                vidActiv = 7;
                                break;
                            default:
                                break;
                        }
                    }

                    cell = row.getCell(weekCount);
                    int rowIndex = cell.getRowIndex();
                    int tempColumnIndex = cell.getColumnIndex();
                    String columnIndex = CellReference.convertNumToColString(tempColumnIndex);
                    String firstFullIndex = columnIndex + "" + rowIndex;
                    String secondFullIndex = columnIndex + "" + (rowIndex + 5);

                    for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                        String firstIndex = CellReference.convertNumToColString(sheet.getMergedRegion(i).getFirstColumn()) + sheet.getMergedRegion(i).getFirstRow();
                        String secondIndex = CellReference.convertNumToColString(sheet.getMergedRegion(i).getLastColumn()) + sheet.getMergedRegion(i).getLastRow();
                        if (firstIndex.equals(firstFullIndex) && secondIndex.equals(secondFullIndex))
                            cellsMerged = true;
                    }

                    ppstatement.setInt(1, idGrafikEduId);
                    ppstatement.setInt(2, daysCount);
                    ppstatement.setInt(3, monthCount);
                    ppstatement.setInt(4, yearStart);
                    ppstatement.setInt(5, vidActiv);
                    ppstatement.executeUpdate();

                    if (weekCount == 53 && daysCount == lastDayDate)
                        break;

                    tableRow++;
                }

                weekCount++;
            }

            System.out.println("------------");
            System.out.println("Добавление элементов прошло успешно!");
        } catch (SQLException e) {
            System.out.println("------------");
            System.out.println("Что-то пошло не так. Добавление элементов в таблицу grafik_education " +
                "не было завершено на 100%");
            e.printStackTrace();
        }
    }

    //***************************************//
    // Заполнение таблицы discipline_plan_ed //
    //***************************************//
    public static void updateDisciplinePlanEd(String file) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet sheet = workbook.getSheet("План");
        HSSFRow row;

        String sql = "INSERT INTO discipline_plan_ed" +
            "(id_discipline, id_course_dis_plan_ed, id_type, id_type_part, id_module_choose," +
            "number, id_teach_plan, id_podrazdelenie) " + "VALUES(?, ?, ?, ?, ?, ?, ?, ?);";
        String getIdTP = "SELECT id FROM teach_plan ORDER BY id DESC LIMIT 1;";
        String getLastId = "SELECT id_course_dis_plan_ed FROM discipline_plan_ed ORDER BY id_course_dis_plan_ed DESC LIMIT 1;";
        String getTypeId = "SELECT id FROM type WHERE value=?;";
        String getTypePartId = "SELECT id FROM type_part WHERE value=?;";
        String getDisciplineId = "SELECT id FROM discipline WHERE LOWER(REPLACE(name, ' ', ''))=?;";
        String getPodrazId = "SELECT id FROM podrazdelenies WHERE LOWER(REPLACE(name, ' ', ''))=?;";

        try {
            // Определение последнего id в таблице teach_plan
            statement = connection.createStatement();
            result = statement.executeQuery(getIdTP);
            result.next();
            int idTeachPlan = result.getInt("id");

            int number = 0;
            int rowCount = 5;
            int idModule = 0;
            int disciplineCount = 0;
            boolean isModuleChoose = false;

            /* Если таблица discipline_plan_ed пуста, присваиваем переменной idCDPE значение 1,
            иначе берём последнее значение id_course_dis_plan_ed из таблицы discipline_plan_ed и инкрементируем его */
            int idCDPE;
            statement = connection.createStatement();
            result = statement.executeQuery(getLastId);
            result.next();
            if (result.getRow() == 0)
                idCDPE = 1;
            else
                idCDPE = result.getInt("id_course_dis_plan_ed") + 1;

            System.out.println("\n\nТаблица discipline_plan_ed");
            System.out.println("------------");

            while ((row = sheet.getRow(rowCount)) != null) {
                // Если ячейка в Excel содержит слово ФТД, значит, завершаем парсинг данной таблицы
                if ((row.getCell(0).getStringCellValue()).contains("ФТД"))
                    break;

                // Пропускаем ненужные строки и переходим к следующим
                if ((row.getCell(0).getStringCellValue()).equals("") ||
                    (row.getCell(1).getStringCellValue()).equals("") ||
                    (row.getCell(2).getStringCellValue()).contains("специализации")) {
                    isModuleChoose = false;
                    number = 0;
                    rowCount++;
                    continue;
                }

                String index = row.getCell(1).getStringCellValue();
                int fullType = index.indexOf('.');
                String type = index.substring(0, fullType);

                String typePart;
                if (index.contains("В.")) typePart = "В";
                else typePart = "Б";

                // Определяем название дисциплины. Если в названии дисциплины находятся определённые слова, пропускаем строчку
                String discipline = (row.getCell(2).getStringCellValue());
                if (discipline.contains("Дисциплины") ||
                    discipline.contains("элективные")) {
                    rowCount++;
                    isModuleChoose = true;
                    idModule += 2;
                    number = 0;
                    continue;
                }

                if (discipline.contains("Производственная практика")) {
                    number = 0;
                    isModuleChoose = false;
                }

                // Ищем нужную дисциплину в таблице discipline и определяем её id
                ppstatement = connection.prepareStatement(getDisciplineId);
                ppstatement.setString(1, discipline.replaceAll("\\s*", "").toLowerCase());
                result = ppstatement.executeQuery();
                result.next();
                int idDiscipline = result.getInt("id");

                // Ищем нужный тип блока в таблице type и определяем его id
                ppstatement = connection.prepareStatement(getTypeId);
                ppstatement.setString(1,type);
                result = ppstatement.executeQuery();
                result.next();
                int idType = result.getInt("id");

                // Ищем к какой части относится дисциплина и определяем id этой части
                ppstatement = connection.prepareStatement(getTypePartId);
                ppstatement.setString(1, typePart);
                result = ppstatement.executeQuery();
                result.next();
                int idTypePart = result.getInt("id");

                int idModuleChoose = 0;
                if (isModuleChoose) {
                    idModuleChoose = idModule;
                    number++;
                }

                // Определяем название и id кафедры
                sheet = workbook.getSheet("ПланСвод");
                row = sheet.getRow(rowCount);
                int lastColumn = row.getLastCellNum() - 1;
                String namePodraz = row.getCell(lastColumn).getStringCellValue();
                ppstatement = connection.prepareStatement(getPodrazId);
                ppstatement.setString(1, namePodraz.replaceAll("\\s*", "").toLowerCase());
                result = ppstatement.executeQuery();
                result.next();
                int idPodrazdelenie = result.getInt("id");

                // Отправляем все собранные данные в таблицу discipline_plan_ed
                sheet = workbook.getSheet("План");
                ppstatement = connection.prepareStatement(sql);
                ppstatement.setInt(1, idDiscipline);
                ppstatement.setInt(2, idCDPE);
                ppstatement.setInt(3, idType);
                ppstatement.setInt(4, idTypePart);
                ppstatement.setInt(5, idModuleChoose);
                ppstatement.setInt(6, number);
                ppstatement.setInt(7, idTeachPlan);
                ppstatement.setInt(8, idPodrazdelenie);
                ppstatement.executeUpdate();

                disciplineCount++;
                idCDPE++;
                rowCount++;
                System.out.println(disciplineCount + (". В таблицу была добавлена дисциплина: ") + discipline);
            }

            System.out.println("------------");
            System.out.println("Добавление дисциплин прошло успешно!");
            System.out.println("Всего было добавлено дисциплин: " + disciplineCount);
        } catch (SQLException e) {
            System.out.println("------------");
            System.out.println("Что-то пошло не так. Добавление дисциплин в таблицу " +
                "discipline_plan_ed не было завершено на 100%.");
            e.printStackTrace();
        }
    }

    //****************************************//
    // Заполнение таблицы course_disc_plan_ed //
    //****************************************//
    public static void updateCourseDiscPlanEd(String file, String tableFormat) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet sheet = workbook.getSheet("Титул");
        HSSFRow row;

        String sql = "INSERT INTO course_disc_plan_ed"
                + "(id_zach_ed, zach_ed, semester, lec, lab, prac, sr, control, id_course_disc_plan_ed) "
                + "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?);";
        String getDPE = "SELECT * FROM discipline_plan_ed ORDER BY id_teach_plan DESC LIMIT 1;";

        System.out.println("\n\nТаблица course_disk_plan_ed");
        System.out.println("------------");

        try {
            int currentSemester = 1;
            int semesterCount = 0;
            int disciplineCount = 1;

            // Определение количества семестров
            String term = "";
            String studyForm = "";
            if (tableFormat.equals("2019") || tableFormat.equals("2020") || tableFormat.equals("2021")) {
                row = sheet.getRow(30);
                studyForm = row.getCell(0).getStringCellValue();
                row = sheet.getRow(31);
                term = row.getCell(0).getStringCellValue();
            } else if (tableFormat.equals("2018")) {
                row = sheet.getRow(31);
                studyForm = row.getCell(1).getStringCellValue();
                row = sheet.getRow(32);
                term = row.getCell(1).getStringCellValue();
            }

            if (term.contains("2г"))
                semesterCount = 4;
            if (term.contains("4г"))
                semesterCount = 8;
            if (term.contains("5л"))
                semesterCount = 10;
            if (term.contains("2г 6м"))
                semesterCount = 6;
            if (term.contains("4г 6м"))
                semesterCount = 10;
            if (term.contains("5л 6м"))
                semesterCount = 12;

            /* Расчёт количества столбцов для одного семестра,
            а также получение номера столбца, на котором начинается 1-ый семестр */
            sheet = workbook.getSheet("План");
            row = sheet.getRow(0);
            int cellCourseOne = 0;
            int cellCourseTwo = cellCourseOne;

            while (!row.getCell(cellCourseOne).getStringCellValue().contains("Курс 1"))
                cellCourseOne++;

            while (!row.getCell(cellCourseTwo).getStringCellValue().contains("Курс 2"))
                cellCourseTwo++;

            int totalCell;
            if (studyForm.contains("Заочная") && !studyForm.contains("Очно-заочная")) {
                totalCell = cellCourseTwo - cellCourseOne;
                semesterCount /= 2;
            } else {
                totalCell =  (cellCourseTwo - cellCourseOne) / 2;
            }

            while (currentSemester <= semesterCount) {
                int rowCount = 5;

                // Определение id_course_dis_plan_ed в таблице discipline_plan_ed
                statement = connection.createStatement();
                result = statement.executeQuery(getDPE);
                result.next();
                int idCDPE = result.getInt("id_course_dis_plan_ed");

                while ((row = sheet.getRow(rowCount)) != null) {
                    // Если ячейка в Excel содержит слово ФТД, значит, завершаем парсинг данной таблицы
                    if ((row.getCell(0).getStringCellValue()).contains("ФТД"))
                        break;

                    // Пропускаем ненужные строки и переходим к следующим
                    if ((row.getCell(0).getStringCellValue()).equals("") ||
                            (row.getCell(1).getStringCellValue()).equals("") ||
                            (row.getCell(2).getStringCellValue()).contains("Дисциплины") ||
                            (row.getCell(2).getStringCellValue()).contains("элективные дисциплины") ||
                            (row.getCell(2).getStringCellValue()).contains("специализации")) {
                        rowCount++;
                        continue;
                    }

                    int idZachEd = 1;
                    int zachEd = 0;
                    int lec = 0;
                    int lab = 0;
                    int prac = 0;
                    int sr = 0;
                    int control = 0;

                    if (totalCell == 9) {
                        zachEd = Integer.parseInt(0 + row.getCell(cellCourseOne + ((currentSemester - 1) * 9)).getStringCellValue());
                        lec = Integer.parseInt(0 + row.getCell(cellCourseOne + 3 + ((currentSemester - 1) * 9)).getStringCellValue());
                        lab = Integer.parseInt(0 + row.getCell(cellCourseOne + 4 + ((currentSemester - 1) * 9)).getStringCellValue());
                        prac = Integer.parseInt(0 + row.getCell(cellCourseOne + 5 + ((currentSemester - 1) * 9)).getStringCellValue());
                        sr = Integer.parseInt(0 + row.getCell(cellCourseOne + 6 + ((currentSemester - 1) * 9)).getStringCellValue());
                        control = Integer.parseInt(0 + row.getCell(cellCourseOne + 7 + ((currentSemester - 1) * 9)).getStringCellValue());
                    }
                    else if (totalCell == 8) {
                        zachEd = Integer.parseInt(0 + row.getCell(cellCourseOne + ((currentSemester - 1) * 8)).getStringCellValue());
                        lec = Integer.parseInt(0 + row.getCell(cellCourseOne + 2 + ((currentSemester - 1) * 8)).getStringCellValue());
                        lab = Integer.parseInt(0 + row.getCell(cellCourseOne + 3 + ((currentSemester - 1) * 8)).getStringCellValue());
                        prac = Integer.parseInt(0 + row.getCell(cellCourseOne + 4 + ((currentSemester - 1) * 8)).getStringCellValue());
                        sr = Integer.parseInt(0 + row.getCell(cellCourseOne + 5 + ((currentSemester - 1) * 8)).getStringCellValue());
                        control = Integer.parseInt(0 + row.getCell(cellCourseOne + 6 + ((currentSemester - 1) * 8)).getStringCellValue());
                    }
                    else if (totalCell == 7) {
                        zachEd = Integer.parseInt(0 + row.getCell(cellCourseOne + ((currentSemester - 1) * 7)).getStringCellValue());
                        lec = Integer.parseInt(0 + row.getCell(cellCourseOne + 2 + ((currentSemester - 1) * 7)).getStringCellValue());
                        lab = Integer.parseInt(0 + row.getCell(cellCourseOne + 3 + ((currentSemester - 1) * 7)).getStringCellValue());
                        prac = Integer.parseInt(0 + row.getCell(cellCourseOne + 4 + ((currentSemester - 1) * 7)).getStringCellValue());
                        sr = Integer.parseInt(0 + row.getCell(cellCourseOne + 5 + ((currentSemester - 1) * 7)).getStringCellValue());
                        control = Integer.parseInt(0 + row.getCell(cellCourseOne + 6 + ((currentSemester - 1) * 7)).getStringCellValue());
                    } else if (totalCell == 6) {
                        zachEd = Integer.parseInt(0 + row.getCell(cellCourseOne + ((currentSemester - 1) * 6)).getStringCellValue());
                        lec = Integer.parseInt(0 + row.getCell(cellCourseOne + 1 + ((currentSemester - 1) * 6)).getStringCellValue());
                        lab = Integer.parseInt(0 + row.getCell(cellCourseOne + 2 + ((currentSemester - 1) * 6)).getStringCellValue());
                        prac = Integer.parseInt(0 + row.getCell(cellCourseOne + 3 + ((currentSemester - 1) * 6)).getStringCellValue());
                        sr = Integer.parseInt(0 + row.getCell(cellCourseOne + 4 + ((currentSemester - 1) * 6)).getStringCellValue());
                        control = Integer.parseInt(0 + row.getCell(cellCourseOne + 5 + ((currentSemester - 1) * 6)).getStringCellValue());
                    } else if (totalCell == 5) {
                        zachEd = Integer.parseInt(0 + row.getCell(cellCourseOne + ((currentSemester - 1) * 5)).getStringCellValue());
                        lec = Integer.parseInt(0 + row.getCell(cellCourseOne + 1 + ((currentSemester - 1) * 5)).getStringCellValue());
                        lab = Integer.parseInt(0 + row.getCell(cellCourseOne + 2 + ((currentSemester - 1) * 5)).getStringCellValue());
                        prac = Integer.parseInt(0 + row.getCell(cellCourseOne + 3 + ((currentSemester - 1) * 5)).getStringCellValue());
                        sr = Integer.parseInt(0 + row.getCell(cellCourseOne + 4 + ((currentSemester - 1) * 5)).getStringCellValue());
                        control = 0;
                    }

                    // Отправляем все собранные данные в таблицу course_disc_plan_ed
                    ppstatement = connection.prepareStatement(sql);
                    ppstatement.setInt(1, idZachEd);
                    ppstatement.setInt(2, zachEd);
                    ppstatement.setInt(3, currentSemester);
                    ppstatement.setInt(4, lec);
                    ppstatement.setInt(5, lab);
                    ppstatement.setInt(6, prac);
                    ppstatement.setInt(7, sr);
                    ppstatement.setInt(8, control);
                    ppstatement.setInt(9, idCDPE);
                    ppstatement.executeUpdate();

                    if (currentSemester == 1) {
                        System.out.println(disciplineCount + ". В таблицу была добавлена дисциплина: " + row.getCell(2).getStringCellValue());
                        disciplineCount++;
                    }

                    rowCount++;
                    idCDPE++;
                }

                currentSemester++;
            }

            System.out.println("------------");
            System.out.println("Добавление дисциплин прошло успешно!");
            System.out.println("Всего было добавлено дисциплин: " + (disciplineCount - 1));
        } catch (SQLException e) {
            System.out.println("------------");
            System.out.println("Что-то пошло не так. Добавление дисциплин в таблицу  " +
                    "course_disk_plan_ed не было завершено на 100%.");
            e.printStackTrace();
        }
    }

    //*********************************//
    // Заполнение таблицы form_control //
    //*********************************//
    public static void updateFormControl(String file) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet sheet = workbook.getSheet("План");
        HSSFRow row;

        String sql = "INSERT INTO form_control(id_discipline_plan_ed, semester, id_type_control) " +
                "VALUES(?, ?, ?);";
        String getTeachPlanId = "SELECT id_teach_plan FROM discipline_plan_ed ORDER BY id_teach_plan DESC LIMIT 1;";
        String getDpeId = "SELECT id FROM discipline_plan_ed WHERE id_teach_plan=?;";

        System.out.println("\n\nТаблица form_control");
        System.out.println("------------");

        try {
            int rowCount = 5;
            int disciplineCount = 0;

            // Определяем значение последнего id_teach_plan в таблице discipline_plan_ed
            statement = connection.createStatement();
            result = statement.executeQuery(getTeachPlanId);
            result.next();
            int idTeachPlan = result.getInt("id_teach_plan");

            // Определяем количество видов форм контроля (экзамен, зачёт, зачёт с оценкой etc)
            row = sheet.getRow(0);
            int j = 3;
            int formControlCount = 3;
            while (formControlCount != 80) {
                if (row.getCell(j).getStringCellValue().contains("з.е.") ||
                        row.getCell(j).getStringCellValue().contains("Итого акад.часов"))
                    break;
                formControlCount++;
                j++;
            }

            // Определяем нужный id из таблицы discipline_plan_ed
            ppstatement = connection.prepareStatement(getDpeId);
            ppstatement.setInt(1, idTeachPlan);
            result = ppstatement.executeQuery();
            result.next();
            int idDPE = result.getInt("id");

            while ((row = sheet.getRow(rowCount)) != null) {
                // Если ячейка в Excel содержит слово ФТД, значит, завершаем парсинг данной таблицы
                if ((row.getCell(0).getStringCellValue()).contains("ФТД"))
                    break;

                // Пропускаем ненужные строки и переходим к следующим
                if ((row.getCell(1).getStringCellValue()).equals("")||
                        (row.getCell(2).getStringCellValue()).equals("") ||
                        (row.getCell(2).getStringCellValue()).contains("Дисциплины") ||
                        (row.getCell(2).getStringCellValue()).contains("элективные дисциплины") ||
                        (row.getCell(2).getStringCellValue()).contains("специализации")) {
                    rowCount++;
                    continue;
                }

                String discipline = row.getCell(2).getStringCellValue();

                for (int i = 3; i < formControlCount; i++) {
                    String semester = row.getCell(i).getStringCellValue();
                    int semesterCount = 0;
                    int semesterLength = semester.length();

                    for (int x = 0; x < semesterLength; x++) {
                        if (semester.charAt(x) >= '1' && semester.charAt(x) <='9') {
                            semesterCount = Integer.parseInt(String.valueOf(semester.charAt(x)));
                            
                            ppstatement = connection.prepareStatement(sql);
                            ppstatement.setInt(1, idDPE);
                            ppstatement.setInt(2, semesterCount);
                            ppstatement.setInt(3, i - 2);
                            ppstatement.executeUpdate();
                        } else if (semester.charAt(x) >= 'A' && semester.charAt(x) <= 'C') {
                            String semesterLetter = String.valueOf(semester.charAt(x));
                            if (semesterLetter.contains("A")) semesterCount = 10;
                            else if (semesterLetter.contains("B")) semesterCount = 11;
                            else if (semesterLetter.contains("C")) semesterCount = 12;

                            ppstatement = connection.prepareStatement(sql);
                            ppstatement.setInt(1, idDPE);
                            ppstatement.setInt(2, semesterCount);
                            ppstatement.setInt(3, i - 2);
                            ppstatement.executeUpdate();
                        }
                    }
                }

                idDPE++;
                disciplineCount++;
                rowCount++;
                System.out.println(disciplineCount + ". В таблицу была добавлена дисциплина: " + discipline);
            }

            System.out.println("------------");
            System.out.println("Добавление дисциплин прошло успешно!");
            System.out.println("Всего было добавлено дисциплин: " + disciplineCount);
        } catch (SQLException e) {
            System.out.println("------------");
            System.out.println("Что-то пошло не так. Добавление дисциплин в таблицу form_control " +
                    "не было завершено на 100%");
            e.printStackTrace();
        }
    }

    //*******************************************************//
    // Определение типа графиков в таблице График .xls файла //
    //*******************************************************//
    public  static String checkGrafikType(String file) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet sheet = workbook.getSheet("График");
        HSSFRow row = sheet.getRow(2);
        HSSFCell cell = row.getCell(0);

        if (cell.getStringCellValue().contains("Числа"))
            return "Old table";
        else
            return "New table";
    }

    //*******************************************//
    // Добавление данных в CSV файл - Дисциплины //
    //*******************************************//
    public static void checkDisciplines(String fileName) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Дисциплины");
        HSSFRow row;
        HSSFCell cell;
        sheet.setDefaultRowHeight((short) 400);

        String countDisciplines = "SELECT COUNT(id_discipline) AS count FROM discipline_plan_ed WHERE id_teach_plan=?;";
        String getCourseInfo = "SELECT * FROM course_disc_plan_ed WHERE id_course_disc_plan_ed=?;";
        String getFormControl = "SELECT * FROM form_control WHERE id_discipline_plan_ed=?;";
        String getDisciplineName = "SELECT name FROM discipline WHERE id=?;";
        String getDPEIdTP = "SELECT id_teach_plan FROM discipline_plan_ed ORDER BY id_teach_plan DESC LIMIT 1;";
        String getDPEIdCDPE = "SELECT id_course_dis_plan_ed FROM discipline_plan_ed WHERE id_teach_plan=?;";
        String getDPEIdDisc = "SELECT id_discipline FROM discipline_plan_ed WHERE id=?;";
        String getDPEId = "SELECT id FROM discipline_plan_ed WHERE id_teach_plan=?;";

        try {
            statement = connection.createStatement();
            result = statement.executeQuery(getDPEIdTP);
            result.next();
            int idTeachPlan = result.getInt("id_teach_plan");

            ppstatement = connection.prepareStatement(countDisciplines);
            ppstatement.setInt(1, idTeachPlan);
            result = ppstatement.executeQuery();
            result.next();
            int disciplineCount = result.getInt("count");

            ppstatement = connection.prepareStatement(getDPEId);
            ppstatement.setInt(1, idTeachPlan);
            result = ppstatement.executeQuery();
            result.next();
            int idDPE = result.getInt("id");

            ppstatement = connection.prepareStatement(getDPEIdCDPE);
            ppstatement.setInt(1, idTeachPlan);
            result = ppstatement.executeQuery();
            result.next();
            int idCDPE = result.getInt("id_course_dis_plan_ed");

            int i = 0;
            int j = 1;
            sheet.addMergedRegion(CellRangeAddress.valueOf("A1:E1"));
            row = sheet.createRow(0);
            cell = row.createCell(0);
            cell.setCellValue("Название дисциплин");
            cell = row.createCell(6);
            cell.setCellValue("Семестр");
            sheet.addMergedRegion(CellRangeAddress.valueOf("H1:I1"));
            cell = row.createCell(7);
            cell.setCellValue("Форма контроля");
            CellStyle cellStyle = row.getSheet().getWorkbook().createCellStyle();
            cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
            cell = row.createCell(10);
            cell.setCellValue("З.Е.");
            cell = row.createCell(11);
            cell.setCellValue("Лек");
            cell = row.createCell(12);
            cell.setCellValue("Лаб");
            cell = row.createCell(13);
            cell.setCellValue("Прак");
            cell = row.createCell(14);
            cell.setCellValue("СР");
            cell = row.createCell(15);
            cell.setCellValue("Контроль");

            while (i < disciplineCount) {
                ppstatement = connection.prepareStatement(getDPEIdDisc);
                ppstatement.setInt(1, idDPE);
                result = ppstatement.executeQuery();
                result.next();
                int idDiscipline = result.getInt("id_discipline");

                ppstatement = connection.prepareStatement(getDisciplineName);
                ppstatement.setInt(1, idDiscipline);
                result = ppstatement.executeQuery();
                result.next();
                String nameDiscipline = result.getString("name");

                ppstatement = connection.prepareStatement(getCourseInfo);
                ppstatement.setInt(1, idCDPE);
                result = ppstatement.executeQuery();
                result.next();
                int zachEd = result.getInt("zach_ed");
                int lec = result.getInt("lec");
                int lab =  result.getInt("lab");
                int prac = result.getInt("prac");
                int sr =  result.getInt("sr");
                int control =  result.getInt("control");

                ppstatement = connection.prepareStatement(getFormControl);
                ppstatement.setInt(1, idDPE);
                result = ppstatement.executeQuery();

                while (result.next()) {
                    sheet.addMergedRegion(new CellRangeAddress(j, j, 0, 4));
                    row = sheet.createRow(j);
                    cell = row.createCell(0);
                    cell.setCellValue(nameDiscipline);
                    int semFormControl = result.getInt("semester");
                    int idTypeControl = result.getInt("id_type_control");
                    String nameTypeControl;

                    switch (idTypeControl) {
                        case 1:
                            nameTypeControl = "Экзамен";
                            break;
                        case 2:
                            nameTypeControl = "Зачёт";
                            break;
                        case 3:
                            nameTypeControl = "Зачёт с оценкой";
                            break;
                        case 4:
                            nameTypeControl = "Курсовая работа";
                            break;
                        case 5:
                            nameTypeControl = "Реферат";
                            break;
                        default:
                            nameTypeControl = "РГР";
                            break;
                    }

                    cell = row.createCell(6);
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue(semFormControl);
                    sheet.addMergedRegion(new CellRangeAddress(j, j, 7, 8));
                    cell = row.createCell(7);
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue(nameTypeControl);
                    cell = row.createCell(10);
                    cell.setCellValue(zachEd);
                    cell = row.createCell(11);
                    cell.setCellValue(lec);
                    cell = row.createCell(12);
                    cell.setCellValue(lab);
                    cell = row.createCell(13);
                    cell.setCellValue(prac);
                    cell = row.createCell(14);
                    cell.setCellValue(sr);
                    cell = row.createCell(15);
                    cell.setCellValue(control);

                    j++;
                }

                idCDPE++;
                idDPE++;
                i++;
            }

            FileOutputStream fileOut = new FileOutputStream("Review - " + fileName);
            workbook.write(fileOut);
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    //***************************************//
    // Добавление данных в CSV файл - График //
    //***************************************//
    public static void checkGrafik(String fileName) throws IOException, SQLException {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream("Review - " + fileName));
        HSSFSheet sheet = workbook.createSheet("График");
        HSSFRow row;
        HSSFCell cell;
        sheet.setDefaultRowHeight((short) 400);

        String getTeachPlanInfo = "SELECT * FROM teach_plan ORDER BY id DESC LIMIT 1;";
        String getProfileName = "SELECT name FROM profile WHERE id=?";
        String getGrafikEduInfo = "SELECT * FROM grafik_education WHERE id_teach_plan=?;";
        String getDay = "SELECT * FROM grafik_education_days WHERE day=? and mounth=? and year=? and id_grafik_education=?;";
        String getVidActiv = "SELECT name FROM vid_activ WHERE id=?;";

        statement = connection.createStatement();
        result = statement.executeQuery(getTeachPlanInfo);
        result.next();
        int teachPlanId = result.getInt("id");
        int profileId = result.getInt("id_profile");
        int course = result.getInt("course");

        ppstatement = connection.prepareStatement(getProfileName);
        ppstatement.setInt(1, profileId);
        result = ppstatement.executeQuery();
        result.next();
        String profileName = result.getString("name");

        ppstatement = connection.prepareStatement(getGrafikEduInfo);
        ppstatement.setInt(1, teachPlanId);
        result = ppstatement.executeQuery();
        result.next();
        int grafikEduId = result.getInt("id");
        String year = result.getString("year_start");
        int yearStart = Integer.parseInt(year) - course + 1;

        sheet.addMergedRegion(CellRangeAddress.valueOf("A1:C1"));
        row = sheet.createRow(0);
        cell = row.createCell(0);
        cell.setCellValue("Профиль: ");
        sheet.addMergedRegion(CellRangeAddress.valueOf("D1:H1"));
        cell = row.createCell(3);
        cell.setCellValue(profileName);

        sheet.addMergedRegion(CellRangeAddress.valueOf("A2:C2"));
        row = sheet.createRow(1);
        cell = row.createCell(0);
        cell.setCellValue("Год начала подготовки: ");
        cell = row.createCell(3);
        cell.setCellValue(yearStart);

        row = sheet.createRow(3);
        cell = row.createCell(0);
        cell.setCellValue("Год: " + Integer.parseInt(year));

        int[] days = {5, 7, 8, 10};
        int[] months = {5, 6, 10, 11};
        String[] nameMonths = {"Январь", "Февраль", "Июнь", "Июль"};

        int i = 0;
        while (i < 4) {
            ppstatement = connection.prepareStatement(getDay);
            ppstatement.setInt(1, days[i]);
            ppstatement.setInt(2, months[i]);
            ppstatement.setInt(3, Integer.parseInt(year));
            ppstatement.setInt(4, grafikEduId);
            result = ppstatement.executeQuery();
            result.next();
            int dayType = result.getInt("id_vid_activ");

            ppstatement = connection.prepareStatement(getVidActiv);
            ppstatement.setInt(1, dayType);
            result = ppstatement.executeQuery();
            result.next();
            String vidActiv = result.getString("name");

            sheet.addMergedRegion(CellRangeAddress.valueOf("C" + (6+ i) + ":" + "H" + (6 + i)));
            sheet.addMergedRegion(CellRangeAddress.valueOf("A" + (6+ i) + ":" + "B" + (6 + i)));
            row = sheet.createRow(i + 5);
            cell = row.createCell(0);
            cell.setCellValue(days[i] + " " + nameMonths[i]);
            cell = row.createCell(2);
            cell.setCellValue(vidActiv);

            i++;
        }

        FileOutputStream fileOut = new FileOutputStream("Review - " + fileName);
        workbook.write(fileOut);
        fileOut.close();
    }

    //*********************//
    // Очистка всех таблиц //
    //*********************//
    public static void clearTables() {
        String clearPodraz = "DELETE FROM podrazdelenies; ALTER SEQUENCE podrazdelenies_id_seq RESTART WITH 1;";
        String clearModuleChoose = "DELETE FROM module_choose; ALTER SEQUENCE module_choose_id_seq RESTART WITH 1";
        String clearFormControl = "DELETE FROM form_control; ALTER SEQUENCE form_control_id_seq RESTART WITH 1";
        String clearTeachPlan = "DELETE FROM teach_plan;";
        String clearGrafikEdu = "DELETE FROM grafik_education; ALTER SEQUENCE grafik_education_id_seq RESTART WITH 1";
        String clearGrafikEduDays = "DELETE FROM grafik_education_days; ALTER SEQUENCE grafik_education_days_id_seq RESTART WITH 1";
        String clearCDPE = "DELETE FROM course_disc_plan_ed; ALTER SEQUENCE course_disc_plan_ed_id_seq RESTART WITH 1";
        String clearDPE = "DELETE FROM discipline_plan_ed; ALTER SEQUENCE discipline_plan_ed_id_seq RESTART WITH 1";
        String clearProfile = "DELETE FROM profile;";

        try {
            statement = connection.createStatement();
            int podrazCount = statement.executeUpdate(clearPodraz);
            int moduleChooseCount = statement.executeUpdate(clearModuleChoose);
            int teachPlanCount = statement.executeUpdate(clearTeachPlan);
            int grafikEduCount = statement.executeUpdate(clearGrafikEdu);
            int grafukEduDaysCount = statement.executeUpdate(clearGrafikEduDays);
            int dpeCount = statement.executeUpdate(clearDPE);
            int cdpeCount = statement.executeUpdate(clearCDPE);
            int formControlCount = statement.executeUpdate(clearFormControl);
            int profileCount = statement.executeUpdate(clearProfile);

            System.out.println("Количество удалённых строк в таблице podrazdelenies: " + podrazCount);
            System.out.println("Количество удалённых строк в таблице profile: " + profileCount);
            System.out.println("Количество удалённых строк в таблице module_choose: " + moduleChooseCount);
            System.out.println("Количество удалённых строк в таблице teach_plan: " + teachPlanCount);
            System.out.println("Количество удалённых строк в таблице grafik_education: " + grafikEduCount);
            System.out.println("Количество удалённых строк в таблице grafik_education_days: " + grafukEduDaysCount);
            System.out.println("Количество удалённых строк в таблице discipline_plan_ed: " + dpeCount);
            System.out.println("Количество удалённых строк в таблице course_disc_plan_ed: " + cdpeCount);
            System.out.println("Количество удалённых строк в таблице form_control: " + formControlCount);
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }
}