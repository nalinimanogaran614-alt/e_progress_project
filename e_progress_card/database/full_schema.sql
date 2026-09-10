-- E-Progress Card complete database schema.
-- Import with: mysql -u root -p < full_schema.sql

CREATE DATABASE IF NOT EXISTS e_progress_card
  CHARACTER SET utf8mb4
  COLLATE utf8mb4_unicode_ci;
USE e_progress_card;

CREATE TABLE IF NOT EXISTS users (
    id INT AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(100) NOT NULL UNIQUE,
    password TEXT NOT NULL,
    role VARCHAR(50) NOT NULL DEFAULT 'student',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS departments (
    id INT AUTO_INCREMENT PRIMARY KEY,
    department_name VARCHAR(150) NOT NULL,
    department_code VARCHAR(50) DEFAULT NULL,
    hod_name VARCHAR(150) DEFAULT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS classes (
    id INT AUTO_INCREMENT PRIMARY KEY,
    department_id INT NOT NULL,
    class_name VARCHAR(100) NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS faculty (
    id INT AUTO_INCREMENT PRIMARY KEY,
    user_id INT NOT NULL UNIQUE,
    faculty_name VARCHAR(150) NOT NULL,
    email VARCHAR(150) DEFAULT NULL,
    phone VARCHAR(30) DEFAULT NULL,
    department_id INT DEFAULT NULL,
    class_id INT DEFAULT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS students (
    id INT AUTO_INCREMENT PRIMARY KEY,
    user_id INT DEFAULT NULL UNIQUE,
    register_number VARCHAR(100) NOT NULL UNIQUE,
    student_name VARCHAR(150) NOT NULL,
    email VARCHAR(150) DEFAULT NULL,
    phone VARCHAR(30) DEFAULT NULL,
    department_id INT DEFAULT NULL,
    class_id INT DEFAULT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS subjects (
    id INT AUTO_INCREMENT PRIMARY KEY,
    subject_code VARCHAR(100) NOT NULL UNIQUE,
    subject_name VARCHAR(200) NOT NULL,
    department_id INT DEFAULT NULL,
    class_id INT DEFAULT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS class_subjects (
    id INT AUTO_INCREMENT PRIMARY KEY,
    class_id INT NOT NULL,
    subject_id INT NOT NULL,
    UNIQUE KEY uq_class_subject (class_id, subject_id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS marks (
    id INT AUTO_INCREMENT PRIMARY KEY,
    student_id INT NOT NULL,
    subject_id INT NOT NULL,
    internal DECIMAL(6,2) DEFAULT NULL,
    `external` DECIMAL(6,2) DEFAULT NULL,
    total DECIMAL(6,2) DEFAULT NULL,
    grade VARCHAR(10) DEFAULT NULL,
    result VARCHAR(20) DEFAULT NULL,
    appreciation VARCHAR(200) DEFAULT NULL,
    ca1 DECIMAL(6,2) DEFAULT NULL,
    ca2 DECIMAL(6,2) DEFAULT NULL,
    assignment DECIMAL(6,2) DEFAULT NULL,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    UNIQUE KEY uq_marks_student_subject (student_id, subject_id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS ep_faculty_details (
    faculty_id INT PRIMARY KEY,
    gender VARCHAR(20) DEFAULT NULL,
    dob DATE DEFAULT NULL,
    address TEXT DEFAULT NULL
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS ep_student_details (
    student_id INT PRIMARY KEY,
    gender VARCHAR(20) DEFAULT NULL,
    dob DATE DEFAULT NULL,
    address TEXT DEFAULT NULL
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS result_settings (
    id INT PRIMARY KEY,
    visibility TINYINT(1) NOT NULL DEFAULT 0,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS ep_academic_years (
    id INT AUTO_INCREMENT PRIMARY KEY,
    year_name VARCHAR(20) NOT NULL UNIQUE,
    is_active TINYINT(1) NOT NULL DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS ep_department_batches (
    id INT AUTO_INCREMENT PRIMARY KEY,
    department_id INT NOT NULL UNIQUE,
    course_type ENUM('UG','PG') NOT NULL,
    start_year INT NOT NULL,
    end_year INT NOT NULL,
    course_duration VARCHAR(20) NOT NULL,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS ep_semesters (
    id INT AUTO_INCREMENT PRIMARY KEY,
    course_type ENUM('UG','PG') NOT NULL,
    semester_no INT NOT NULL,
    year_no INT NOT NULL,
    semester_label VARCHAR(40) NOT NULL,
    UNIQUE KEY uq_ep_sem (course_type, semester_no)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS ep_class_semesters (
    id INT AUTO_INCREMENT PRIMARY KEY,
    class_id INT NOT NULL,
    academic_year_id INT NOT NULL,
    course_type VARCHAR(5) NOT NULL,
    year_no INT NOT NULL,
    semester_no INT NOT NULL,
    section VARCHAR(30) DEFAULT 'A',
    UNIQUE KEY uq_ep_class_year_sem (class_id, academic_year_id, semester_no)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS ep_subject_semesters (
    id INT AUTO_INCREMENT PRIMARY KEY,
    subject_id INT NOT NULL,
    class_id INT NOT NULL,
    academic_year_id INT NOT NULL,
    semester_no INT NOT NULL,
    UNIQUE KEY uq_ep_sub_sem (subject_id, class_id, academic_year_id, semester_no)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS ep_faculty_assignments (
    id INT AUTO_INCREMENT PRIMARY KEY,
    faculty_id INT NOT NULL,
    academic_year_id INT NOT NULL,
    semester_no INT NOT NULL,
    class_id INT NOT NULL,
    subject_id INT NOT NULL,
    section VARCHAR(30) DEFAULT 'A',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE KEY uq_ep_fac_assign (faculty_id, academic_year_id, semester_no, class_id, subject_id)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS ep_student_semesters (
    id INT AUTO_INCREMENT PRIMARY KEY,
    student_id INT NOT NULL,
    academic_year_id INT NOT NULL,
    semester_no INT NOT NULL,
    year_no INT NOT NULL,
    class_id INT NOT NULL,
    section VARCHAR(30) DEFAULT 'A',
    status VARCHAR(20) DEFAULT 'Current',
    UNIQUE KEY uq_ep_student_sem (student_id, academic_year_id, semester_no)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS ep_question_papers (
    id INT AUTO_INCREMENT PRIMARY KEY,
    faculty_id INT NOT NULL,
    subject_id INT NOT NULL,
    academic_year_id INT NOT NULL,
    semester_no INT NOT NULL,
    cia VARCHAR(10) NOT NULL,
    file_path TEXT,
    paper_json LONGTEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE KEY uq_ep_qp (faculty_id, subject_id, academic_year_id, semester_no, cia)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS ep_assignment_marks (
    id INT AUTO_INCREMENT PRIMARY KEY,
    student_id INT NOT NULL,
    subject_id INT NOT NULL,
    academic_year_id INT NOT NULL,
    semester_no INT NOT NULL,
    marks DECIMAL(6,2) NOT NULL DEFAULT 0,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    UNIQUE KEY uq_ep_assignment_mark (student_id, subject_id, academic_year_id, semester_no)
) ENGINE=InnoDB;

CREATE TABLE IF NOT EXISTS ep_question_marks (
    id INT AUTO_INCREMENT PRIMARY KEY,
    student_id INT NOT NULL,
    subject_id INT NOT NULL,
    academic_year_id INT NOT NULL,
    semester_no INT NOT NULL,
    exam_type VARCHAR(10) NOT NULL,
    question_no VARCHAR(30) NOT NULL,
    question_text TEXT NULL,
    max_marks DECIMAL(6,2) NOT NULL DEFAULT 0,
    obtained_marks DECIMAL(6,2) NULL,
    choice VARCHAR(5) NULL,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    UNIQUE KEY uq_ep_qmark (student_id, subject_id, academic_year_id, semester_no, exam_type, question_no)
) ENGINE=InnoDB;

INSERT INTO result_settings (id, visibility)
VALUES (1, 0)
ON DUPLICATE KEY UPDATE id = id;

INSERT INTO ep_semesters (course_type, semester_no, year_no, semester_label)
VALUES
  ('UG', 1, 1, 'Semester 1'), ('UG', 2, 1, 'Semester 2'),
  ('UG', 3, 2, 'Semester 3'), ('UG', 4, 2, 'Semester 4'),
  ('UG', 5, 3, 'Semester 5'), ('UG', 6, 3, 'Semester 6'),
  ('PG', 1, 1, 'Semester 1'), ('PG', 2, 1, 'Semester 2'),
  ('PG', 3, 2, 'Semester 3'), ('PG', 4, 2, 'Semester 4')
ON DUPLICATE KEY UPDATE semester_label = VALUES(semester_label), year_no = VALUES(year_no);