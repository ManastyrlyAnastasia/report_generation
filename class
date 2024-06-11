<?php

require 'vendor/autoload.php'; // Подключение автозагрузчика Composer

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

interface Report {
    public function xlsxHeaders();
    public function Query($filterId = null);
    public function extraFields();
}

class FacultyReport implements Report {
    public function xlsxHeaders() {
        return ['Faculty ID', 'Faculty Name', 'Specialization Name', 'Student ID', 'Student Name', 'Student Surname'];
    }

    public function Query($facultyId = null) {
        if ($facultyId) {
            return "SELECT faculty.id AS faculty_id, faculty.name AS faculty_name, specialization.name AS specialization_name, students.id AS student_id, students.name AS student_name, students.surname AS student_surname
                    FROM faculty
                    JOIN specialization ON faculty.id = specialization.faculty_id
                    JOIN groups ON specialization.id = groups.specialization_id
                    JOIN students ON groups.id = students.groups_id
                    WHERE faculty.id = $facultyId";
        }
        return "SELECT faculty.id AS faculty_id, faculty.name AS faculty_name, specialization.name AS specialization_name, students.id AS student_id, students.name AS student_name, students.surname AS student_surname
                FROM faculty
                JOIN specialization ON faculty.id = specialization.faculty_id
                JOIN groups ON specialization.id = groups.specialization_id
                JOIN students ON groups.id = students.groups_id";
    }

    public function extraFields() {
        return [];
    }
}

class GroupReport implements Report {
    public function xlsxHeaders() {
        return ['Group ID', 'Group Name', 'Start Year', 'End Year', 'Student ID', 'Student Name', 'Student Surname'];
    }

    public function Query($groupId = null) {
        if ($groupId) {
            return "SELECT groups.id AS group_id, groups.name AS group_name, groups.start_year, groups.end_year, students.id AS student_id, students.name AS student_name, students.surname AS student_surname
                    FROM groups 
                    LEFT JOIN students ON groups.id = students.groups_id 
                    WHERE groups.id = $groupId";
        }
        return "SELECT groups.id AS group_id, groups.name AS group_name, groups.start_year, groups.end_year, students.id AS student_id, students.name AS student_name, students.surname AS student_surname
                FROM groups 
                LEFT JOIN students ON groups.id = students.groups_id";
    }

    public function extraFields() {
        return [];
    }
}

class SpecializationReport implements Report {
    public function xlsxHeaders() {
        return ['Specialization ID', 'Specialization Name', 'Student Name', 'Student Surname', 'Group Name', 'Faculty Name'];
    }

    public function Query($specializationId = null) {
        if ($specializationId) {
            return "SELECT specialization.id AS specialization_id, specialization.name AS specialization_name, students.name AS student_name, students.surname AS student_surname, groups.name AS group_name, faculty.name AS faculty_name
                    FROM specialization
                    JOIN groups ON specialization.id = groups.specialization_id
                    JOIN students ON groups.id = students.groups_id
                    JOIN faculty ON specialization.faculty_id = faculty.id
                    WHERE specialization.id = $specializationId";
        }
        return "SELECT specialization.id AS specialization_id, specialization.name AS specialization_name, students.name AS student_name, students.surname AS student_surname, groups.name AS group_name, faculty.name AS faculty_name
                FROM specialization
                JOIN groups ON specialization.id = groups.specialization_id
                JOIN students ON groups.id = students.groups_id
                JOIN faculty ON specialization.faculty_id = faculty.id";
    }

    public function extraFields() {
        return [];
    }
}

class StudentReport implements Report {
    public function xlsxHeaders() {
        return ['id', 'name', 'surname', 'birth', 'address', 'group_name']; // Изменено 'groups_id' на 'group_name'
    }

    public function Query($specializationId = null) {
        return "SELECT students.id, students.name, students.surname, students.birth, students.address, groups.name AS group_name 
                FROM students 
                JOIN groups ON students.groups_id = groups.id"; // Изменено соединение с таблицей groups
    }

    public function extraFields() {
        return [];
    }
}

class ReportGenerate {
    private $report;
    private $reportId;

    public function __construct(Report $report, $reportId) {
        $this->report = $report;
        $this->reportId = $reportId;
    }

    public function getData($filterId = null) {
        $servername = 'localhost';
        $username = 'root';
        $password = '';
        $dbname = 'data';
        $conn = new mysqli($servername, $username, $password, $dbname);

        if ($conn->connect_error) {
            die("Connection failed: " . $conn->connect_error);
        }

        $query = $this->report->Query($filterId);
        $result = $conn->query($query);

        if (!$result) {
            die("Ошибка выполнения запроса: " . $conn->error);
        }

        $data = [];
        if ($result->num_rows > 0) {
            while ($row = $result->fetch_assoc()) {
                $data[] = $row;
            }
        }

        $conn->close();
        return $data;
    }

    public function generateXlsx($filterId = null) {
        $headers = $this->report->xlsxHeaders();
        $data = $this->getData($filterId);

        // Создание XLSX файла
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // Запись заголовков
        foreach ($headers as $index => $header) {
            $cell = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($index + 1) . '1';
            $sheet->setCellValue($cell, $header);
        }

        // Центрирование заголовков
        $headerStyle = [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ]
        ];

        foreach (range('A', \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex(count($headers))) as $columnID) {
            $sheet->getStyle($columnID . '1')->applyFromArray($headerStyle);
        }

        // Запись данных
        $rightAlignStyle = [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_RIGHT,
            ]
        ];

        foreach ($data as $rowIndex => $row) {
            $rowNumber = $rowIndex + 2; // Строки данных начинаются со второй строки
            $colIndex = 1;
            foreach ($row as $value) {
                $cell = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colIndex) . $rowNumber;
                $sheet->setCellValue($cell, $value);
                
                // Применение выравнивания по правому краю для всех столбцов
                $sheet->getStyle($cell)->applyFromArray($rightAlignStyle);

                $colIndex++;
            }
        }

        // Установка ширины столбцов
        $sheet->getColumnDimension('A')->setWidth(20);
        $sheet->getColumnDimension('B')->setWidth(25);
        $sheet->getColumnDimension('C')->setWidth(20);
        $sheet->getColumnDimension('D')->setWidth(20);
        $sheet->getColumnDimension('E')->setWidth(25);
        $sheet->getColumnDimension('F')->setWidth(25);

        // Генерация имени файла
        $timestamp = date('YmdHis');
        $directory = __DIR__ . '/DATA';
        $filename = "{$directory}/{$timestamp}.xlsx";

        // Убедитесь, что директория существует
        if (!file_exists($directory)) {
            mkdir($directory, 0777, true);
        }

        // Сохранение XLSX файла
        $writer = new Xlsx($spreadsheet);
        $writer->save($filename);

        // Запись информации о файле в базу данных
        $this->saveFileInfo($timestamp, $filename);

        return $filename;
    }

    private function saveFileInfo($timestamp, $filename) {
        // Подключение к базе данных
        $servername = 'localhost';
        $username = 'root';
        $password = '';
        $dbname = 'data';
        $conn = new mysqli($servername, $username, $password, $dbname);

        if ($conn->connect_error) {
            die("Connection failed: " . $conn->connect_error);
        }

        $time = date('Y-m-d H:i:s', strtotime($timestamp));

        // Убедитесь, что status_id существует
        $status_id = 1; // Пример ID статуса
        $status_check = $conn->query("SELECT id FROM status WHERE id = $status_id");

        if ($status_check->num_rows == 0) {
            die("Статус с ID $status_id не существует.");
        }

        // Убедитесь, что user_id существует
        $user_id = 1; // Пример ID пользователя
        $user_check = $conn->query("SELECT id FROM user WHERE id = $user_id");

        if ($user_check->num_rows == 0) {
            die("Пользователь с ID $user_id не существует.");
        }

        // Убедитесь, что report_id существует
        $report_id_check = $conn->query("SELECT id FROM report WHERE id = " . intval($this->reportId));

        if ($report_id_check->num_rows == 0) {
            die("Отчет с ID " . intval($this->reportId) . " не существует.");
        }

        $stmt = $conn->prepare("INSERT INTO report_history (report_id, status_id, user_id, time, filename, message) VALUES (?, ?, ?, ?, ?, ?)");
        $message = "Report generated successfully";
        $stmt->bind_param('iiisss', $this->reportId, $status_id, $user_id, $time, $filename, $message);
        $stmt->execute();
        $stmt->close();
        $conn->close();
    }
}

// Получение списка отчетов из базы данных для формы
function getReports() {
    $servername = 'localhost';
    $username = 'root';
    $password = '';
    $dbname = 'data';

    $conn = new mysqli($servername, $username, $password, $dbname);

    if ($conn->connect_error) {
        die("Connection failed: " . $conn->connect_error);
    }

    $result = $conn->query("SELECT id, string_id, name FROM report");
    $reports = [];
    if ($result->num_rows > 0) {
        while ($row = $result->fetch_assoc()) {
            $reports[] = $row;
        }
    }

    $conn->close();
    return $reports;
}

// Получение списка групп из базы данных для формы
function getGroups() {
    $servername = 'localhost';
    $username = 'root';
    $password = '';
    $dbname = 'data';

    $conn = new mysqli($servername, $username, $password, $dbname);

    if ($conn->connect_error) {
        die("Connection failed: " . $conn->connect_error);
    }

    $result = $conn->query("SELECT id, name FROM groups");
    $groups = [];
    if ($result->num_rows > 0) {
        while ($row = $result->fetch_assoc()) {
            $groups[] = $row;
        }
    }

    $conn->close();
    return $groups;
}

// Получение списка специальностей из базы данных для формы
function getSpecializations() {
    $servername = 'localhost';
    $username = 'root';
    $password = '';
    $dbname = 'data';

    $conn = new mysqli($servername, $username, $password, $dbname);

    if ($conn->connect_error) {
        die("Connection failed: " . $conn->connect_error);
    }

    $result = $conn->query("SELECT id, name FROM specialization");
    $specializations = [];
    if ($result->num_rows > 0) {
        while ($row = $result->fetch_assoc()) {
            $specializations[] = $row;
        }
    }

    $conn->close();
    return $specializations;
}

// Получение списка факультетов из базы данных для формы
function getFaculties() {
    $servername = 'localhost';
    $username = 'root';
    $password = '';
    $dbname = 'data';

    $conn = new mysqli($servername, $username, $password, $dbname);

    if ($conn->connect_error) {
        die("Connection failed: " . $conn->connect_error);
    }

    $result = $conn->query("SELECT id, name FROM faculty");
    $faculties = [];
    if ($result->num_rows > 0) {
        while ($row = $result->fetch_assoc()) {
            $faculties[] = $row;
        }
    }

    $conn->close();
    return $faculties;
}

// HTML форма
$reports = getReports();
$groups = getGroups();
$specializations = getSpecializations();
$faculties = getFaculties();
?>
<!DOCTYPE html>
<html>
<head>
    <title>Выбор отчета</title>
    <script>
        function showExtraDropdown() {
            var reportType = document.getElementById("report_type").value;
            var groupDropdown = document.getElementById("group_dropdown");
            var specializationDropdown = document.getElementById("specialization_dropdown");
            var facultyDropdown = document.getElementById("faculty_dropdown");
            if (reportType == "GRP_STATS") { // Assuming 'GRP_STATS' is the string_id for GroupReport
                groupDropdown.style.display = "block";
                specializationDropdown.style.display = "none";
                facultyDropdown.style.display = "none";
            } else if (reportType == "SPE_STATS") { // Assuming 'SPE_STATS' is the string_id for SpecializationReport
                groupDropdown.style.display = "none";
                specializationDropdown.style.display = "block";
                facultyDropdown.style.display = "none";
            } else if (reportType == "FAC_STATS") { // Assuming 'FAC_STATS' is the string_id for FacultyReport
                groupDropdown.style.display = "none";
                specializationDropdown.style.display = "none";
                facultyDropdown.style.display = "block";
            } else {
                groupDropdown.style.display = "none";
                specializationDropdown.style.display = "none";
                facultyDropdown.style.display = "none";
            }
        }
    </script>
</head>
<body>
    <form method="post" action="">
        <label for="report_type">Выберите тип отчета:</label>
        <select name="report_type" id="report_type" onchange="showExtraDropdown()">
            <option value="">Выберите тип отчета</option>
            <?php foreach ($reports as $report): ?>
                <option value="<?php echo $report['string_id']; ?>"><?php echo $report['name']; ?></option>
            <?php endforeach; ?>
        </select>

        <div id="group_dropdown" style="display:none;">
            <label for="group_id">Выберите группу:</label>
            <select name="group_id" id="group_id">
                <?php foreach ($groups as $group): ?>
                    <option value="<?php echo $group['id']; ?>"><?php echo $group['name']; ?></option>
                <?php endforeach; ?>
            </select>
        </div>

        <div id="specialization_dropdown" style="display:none;">
            <label for="specialization_id">Выберите специальность:</label>
            <select name="specialization_id" id="specialization_id">
                <?php foreach ($specializations as $specialization): ?>
                    <option value="<?php echo $specialization['id']; ?>"><?php echo $specialization['name']; ?></option>
                <?php endforeach; ?>
            </select>
        </div>

        <div id="faculty_dropdown" style="display:none;">
            <label for="faculty_id">Выберите факультет:</label>
            <select name="faculty_id" id="faculty_id">
                <?php foreach ($faculties as $faculty): ?>
                    <option value="<?php echo $faculty['id']; ?>"><?php echo $faculty['name']; ?></option>
                <?php endforeach; ?>
            </select>
        </div>

        <button type="submit" name="generate_report">Сгенерировать отчет</button>
    </form>

    <?php
    if (isset($_POST['generate_report'])) {
        $reportType = $_POST['report_type'];
        $groupId = isset($_POST['group_id']) ? $_POST['group_id'] : null;
        $specializationId = isset($_POST['specialization_id']) ? $_POST['specialization_id'] : null;
        $facultyId = isset($_POST['faculty_id']) ? $_POST['faculty_id'] : null;

        switch ($reportType) {
            case 'FAC_STATS':
                $report = new FacultyReport();
                break;
            case 'GRP_STATS':
                $report = new GroupReport();
                break;
            case 'SPE_STATS':
                $report = new SpecializationReport();
                break;
            case 'UNI_STATS':
                $report = new StudentReport();
                break;
            default:
                throw new Exception("Invalid report type");
        }

        $reportGenerator = new ReportGenerate($report, $reportType);
        $filename = $reportGenerator->generateXlsx($reportType == 'GRP_STATS' ? $groupId : ($reportType == 'SPE_STATS' ? $specializationId : ($reportType == 'FAC_STATS' ? $facultyId : null)));

        if ($filename) {
            // Установим заголовки для скачивания файла
            header('Content-Description: File Transfer');
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment; filename="'.basename($filename).'"');
            header('Expires: 0');
            header('Cache-Control: must-revalidate');
            header('Pragma: public');
            header('Content-Length: ' . filesize($filename));
            readfile($filename);
            exit;
        } else {
            echo "Ошибка генерации отчета.";
        }
    }
    ?>
</body>
</html>
