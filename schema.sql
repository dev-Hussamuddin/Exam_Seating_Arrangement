CREATE DATABASE IF NOT EXISTS exam_seating_db;
USE exam_seating_db;

-- Master data: class and department normally remain unchanged each year.
CREATE TABLE IF NOT EXISTS classes (
    class_id INT PRIMARY KEY AUTO_INCREMENT,
    class_name VARCHAR(50) NOT NULL UNIQUE,
    department VARCHAR(100) NOT NULL
) ENGINE=InnoDB;

-- Excluded roll numbers are stored as comma-separated text, such as 5,12,27.
CREATE TABLE IF NOT EXISTS student_batches (
    batch_id INT PRIMARY KEY AUTO_INCREMENT,
    class_id INT NOT NULL,
    academic_year VARCHAR(20),
    roll_start INT NOT NULL,
    roll_end INT NOT NULL,
    number_of_students INT NOT NULL,
    non_included_rolls VARCHAR(255),
    subject VARCHAR(150),
    roll_numbers TEXT,
    semester VARCHAR(50),
    FOREIGN KEY (class_id) REFERENCES classes(class_id)
) ENGINE=InnoDB;

-- Capacity is the number of benches: one bench seats one student.
CREATE TABLE IF NOT EXISTS classrooms (
    classroom_id INT PRIMARY KEY AUTO_INCREMENT,
    classroom_no VARCHAR(20) NOT NULL UNIQUE,
    capacity INT NOT NULL
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS timetable (
    timetable_id INT PRIMARY KEY AUTO_INCREMENT,
    class_id INT NOT NULL,
    subject VARCHAR(150) NOT NULL,
    semester VARCHAR(50),
    exam_date DATE,
    start_time TIME,
    end_time TIME,
    time_text VARCHAR(255),
    FOREIGN KEY (class_id) REFERENCES classes(class_id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS seating_arrangements (
    arrangement_id INT PRIMARY KEY AUTO_INCREMENT,
    timetable_id INT NOT NULL,
    classroom_id INT NOT NULL,
    roll_start INT NOT NULL,
    roll_end INT NOT NULL,
    allocated_count INT NOT NULL,
    roll_numbers VARCHAR(500) NOT NULL,
    block_number INT NULL,
    seat_numbers TEXT NULL,
    FOREIGN KEY (timetable_id) REFERENCES timetable(timetable_id),
    FOREIGN KEY (classroom_id) REFERENCES classrooms(classroom_id)
) ENGINE=InnoDB;
