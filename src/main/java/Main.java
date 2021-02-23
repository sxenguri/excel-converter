import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.SQLException;
import java.util.Properties;

public class Main {
    public static void main(String[] args) throws IOException {
        try {
            String actionType = args[0]; // Тип действия, который хочет выполнить пользователь

            //---------------------------//
            // Подключение к базе данных //
            //---------------------------//
            Properties props = new Properties();
            try (InputStream in = Files.newInputStream(Paths.get("database.properties"))) {
                props.load(in);
            }

            String DATABASE = props.getProperty("DATABASE");
            String PASS = props.getProperty("PASSWORD");
            String USER = props.getProperty("LOGIN");
            String PORT = props.getProperty("PORT");
            String IP = props.getProperty("IP");

            String DB_URL = "jdbc:postgresql://" + IP + ":" + PORT + "/" + DATABASE;

            Database.connect(DB_URL, USER, PASS);

            //------------------------------//
            // Очистка таблиц в базе данных //
            //------------------------------//
            if (actionType.equals("clear")) {
                Database.clearTables();
            }

            //----------------------------------------------------//
            // Добавление данных одного Excel файла в базу данных //
            //----------------------------------------------------//
            else if (actionType.equals("update")) {
                String fileName = args[1]; // Название .xls файла
                String tableFormat = args[2]; //

                //--------------------------------//
                // Проверка недостающих дисциплин //
                //--------------------------------//
                Database.checkDiscipline(fileName);

                //------------------------//
                // Таблица podrazdelenies //
                //------------------------//
                Database.updatePodrazdelenies(fileName);

                //-----------------//
                // Таблица profile //
                //-----------------//
                Database.updateProfile(fileName, tableFormat);

                //-----------------------//
                // Таблица module_choose //
                //-----------------------//
                Database.updateModuleChoose(fileName);

                //--------------------//
                // Таблица teach_plan //
                //--------------------//
                Database.updateTeachPlan(fileName, tableFormat);

                //--------------------------//
                // Таблица grafik_education //
                //--------------------------//
                Database.updateGrafikEducation();

                //-------------------------------//
                // Таблица grafik_education_days //
                //-------------------------------//
                String grafikFormat = Database.checkGrafikType(fileName);
                if (grafikFormat.contains("Old table"))
                    Database.updateGrafikEducationDays(fileName, tableFormat);
                else
                    Database.updateGrafikEducationDaysNew(fileName, tableFormat);

                //----------------------------//
                // Таблица discipline_plan_ed //
                //----------------------------//
                Database.updateDisciplinePlanEd(fileName);

                //-----------------------------//
                // Таблица course_disk_plan_ed //
                //-----------------------------//
                Database.updateCourseDiscPlanEd(fileName, tableFormat);

                //----------------------//
                // Таблица form_control //
                //----------------------//
                Database.updateFormControl(fileName);

                //----------------------------------------------------//
                // Создание нового excel файла с проверочными данными //
                //----------------------------------------------------//
                Database.checkDisciplines(fileName);
                Database.checkGrafik(fileName);
            }

            //---------------------------------------------------//
            // Добавление данных всех Excel файлов в базу данных //
            //---------------------------------------------------//
            else if (actionType.equals("update_all")) {
                String tableFormat = args[1];
                File folder = new File("Файлы для парсинга");
                String path = folder.getAbsolutePath();
                File[] listOfFiles = folder.listFiles();
                assert listOfFiles != null;

                for (File listOfFile : listOfFiles) {
                    String fileName = path + "/" + listOfFile.getName();;

                    Database.checkDiscipline(fileName); // Проверка недостающих дисциплин
                    Database.updatePodrazdelenies(fileName); // Добавление данных в таблицу podrazdelenies
                    Database.updateProfile(fileName, tableFormat); // Добавление данных в таблицу profile
                    Database.updateModuleChoose(fileName); // Добавление данных в таблицу module_choose
                    Database.updateTeachPlan(fileName, tableFormat); // Добавление данных в таблицу teach_plan
                    Database.updateGrafikEducation(); // Добавление данных в таблицу grafik_education
                    String grafikFormat = Database.checkGrafikType(fileName);
                    // В зависимости от типа графика в таблице График .xls файла применяем тот или иной метод
                    if (grafikFormat.contains("Old table"))
                        Database.updateGrafikEducationDays(fileName, tableFormat);
                    else
                        Database.updateGrafikEducationDaysNew(fileName, tableFormat);
                    Database.updateDisciplinePlanEd(fileName); // Добавление данных в таблицу discipline_plan_ed
                    Database.updateCourseDiscPlanEd(fileName, tableFormat); // Добавление данных в таблицу course_disc_plan_ed
                    Database.updateFormControl(fileName); // Добавление данных в таблицу form_control
                }
            }
        } catch (ArrayIndexOutOfBoundsException | SQLException e) {
            e.printStackTrace();
        }
    }
}