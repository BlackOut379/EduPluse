
import sqlite3

DATABASE = "edupluse.db"


# -------------------------
# DATABASE CONNECTION
# -------------------------

def get_connection():
    connection = sqlite3.connect(DATABASE)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON")
    return connection


# -------------------------
# DATABASE INITIALIZATION
# -------------------------

def initialize_database():

    connection = get_connection()

    try:
        connection.executescript("""

        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE,
            password TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('teacher', 'student')),
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );


        CREATE TABLE IF NOT EXISTS classes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE
        );


        CREATE TABLE IF NOT EXISTS subjects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE
        );


        CREATE TABLE IF NOT EXISTS teacher_assignments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_id INTEGER NOT NULL,
            class_id INTEGER NOT NULL,
            subject_id INTEGER NOT NULL,

            FOREIGN KEY (teacher_id) REFERENCES users(id),
            FOREIGN KEY (class_id) REFERENCES classes(id),
            FOREIGN KEY (subject_id) REFERENCES subjects(id),

            UNIQUE (teacher_id, class_id, subject_id)
        );


        CREATE TABLE IF NOT EXISTS enrollments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            class_id INTEGER NOT NULL,

            FOREIGN KEY (student_id) REFERENCES users(id),
            FOREIGN KEY (class_id) REFERENCES classes(id),

            UNIQUE (student_id, class_id)
        );


        CREATE TABLE IF NOT EXISTS teaching_materials (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_id INTEGER NOT NULL,
            class_id INTEGER NOT NULL,
            subject_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            content TEXT,
            file_name TEXT NOT NULL,
            file_path TEXT NOT NULL,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (teacher_id) REFERENCES users(id),
            FOREIGN KEY (class_id) REFERENCES classes(id),
            FOREIGN KEY (subject_id) REFERENCES subjects(id)
        );


        CREATE TABLE IF NOT EXISTS syllabuses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_id INTEGER NOT NULL,
            class_id INTEGER NOT NULL,
            subject_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            file_name TEXT NOT NULL,
            file_path TEXT NOT NULL,
            uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (teacher_id) REFERENCES users(id),
            FOREIGN KEY (class_id) REFERENCES classes(id),
            FOREIGN KEY (subject_id) REFERENCES subjects(id),

            UNIQUE (class_id, subject_id)
        );


        CREATE TABLE IF NOT EXISTS assessments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_id INTEGER NOT NULL,
            class_id INTEGER NOT NULL,
            subject_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            start_time TEXT,
            end_time TEXT,
            status TEXT DEFAULT 'draft'
                CHECK(status IN ('draft', 'published')),
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (teacher_id) REFERENCES users(id),
            FOREIGN KEY (class_id) REFERENCES classes(id),
            FOREIGN KEY (subject_id) REFERENCES subjects(id)
        );


        CREATE TABLE IF NOT EXISTS questions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            assessment_id INTEGER NOT NULL,
            question TEXT NOT NULL,
            option_a TEXT NOT NULL,
            option_b TEXT NOT NULL,
            option_c TEXT NOT NULL,
            option_d TEXT NOT NULL,
            correct_answer TEXT NOT NULL,
            explanation TEXT,
            concept TEXT,

            FOREIGN KEY (assessment_id)
                REFERENCES assessments(id)
                ON DELETE CASCADE
        );


        CREATE TABLE IF NOT EXISTS attempts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            assessment_id INTEGER NOT NULL,
            score INTEGER DEFAULT 0,
            submitted_at TEXT DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (student_id) REFERENCES users(id),
            FOREIGN KEY (assessment_id)
                REFERENCES assessments(id),

            UNIQUE (student_id, assessment_id)
        );


        CREATE TABLE IF NOT EXISTS responses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            attempt_id INTEGER NOT NULL,
            question_id INTEGER NOT NULL,
            selected_answer TEXT NOT NULL,
            is_correct INTEGER NOT NULL DEFAULT 0,

            FOREIGN KEY (attempt_id)
                REFERENCES attempts(id)
                ON DELETE CASCADE,

            FOREIGN KEY (question_id)
                REFERENCES questions(id),

            UNIQUE (attempt_id, question_id)
        );


        CREATE TABLE IF NOT EXISTS attendance_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_id INTEGER NOT NULL,
            class_id INTEGER NOT NULL,
            subject_id INTEGER NOT NULL,
            lecture_date TEXT NOT NULL,
            lecture_number INTEGER NOT NULL,

            FOREIGN KEY (teacher_id) REFERENCES users(id),
            FOREIGN KEY (class_id) REFERENCES classes(id),
            FOREIGN KEY (subject_id) REFERENCES subjects(id),

            UNIQUE (
                teacher_id,
                class_id,
                subject_id,
                lecture_date,
                lecture_number
            )
        );


        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER NOT NULL,
            student_id INTEGER NOT NULL,
            status TEXT NOT NULL
                CHECK(status IN ('Present', 'Absent')),

            FOREIGN KEY (session_id)
                REFERENCES attendance_sessions(id)
                ON DELETE CASCADE,

            FOREIGN KEY (student_id) REFERENCES users(id),

            UNIQUE (session_id, student_id)
        );


        CREATE TABLE IF NOT EXISTS academic_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            subject_id INTEGER NOT NULL,
            record_type TEXT NOT NULL,
            score REAL NOT NULL,
            maximum_score REAL NOT NULL,
            recorded_at TEXT DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (student_id) REFERENCES users(id),
            FOREIGN KEY (subject_id) REFERENCES subjects(id),

            UNIQUE (student_id, subject_id, record_type)
        );


        CREATE TABLE IF NOT EXISTS ai_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_id INTEGER NOT NULL,
            feature TEXT NOT NULL,
            request_key TEXT NOT NULL UNIQUE,
            model TEXT,
            status TEXT DEFAULT 'pending',
            error_message TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            completed_at TEXT,

            FOREIGN KEY (teacher_id) REFERENCES users(id)
        );


        """)

        connection.commit()

    except Exception:
        connection.rollback()
        raise

    finally:
        connection.close()


# -------------------------
# AI REQUEST TRACKING
# -------------------------

def start_ai_request(teacher_id, feature, request_key, model=None):

    connection = get_connection()

    try:
        cursor = connection.execute("""
            INSERT INTO ai_requests (
                teacher_id,
                feature,
                request_key,
                model,
                status
            )
            VALUES (?, ?, ?, ?, 'pending')
        """, (
            teacher_id,
            feature,
            request_key,
            model
        ))

        connection.commit()
        return cursor.lastrowid

    finally:
        connection.close()


def complete_ai_request(request_key):

    connection = get_connection()

    try:
        connection.execute("""
            UPDATE ai_requests
            SET status = 'completed',
                completed_at = CURRENT_TIMESTAMP,
                error_message = NULL
            WHERE request_key = ?
        """, (request_key,))

        connection.commit()

    finally:
        connection.close()


def fail_ai_request(request_key, error_message=None):

    connection = get_connection()

    try:
        connection.execute("""
            UPDATE ai_requests
            SET status = 'failed',
                error_message = ?,
                completed_at = CURRENT_TIMESTAMP
            WHERE request_key = ?
        """, (
            str(error_message) if error_message else None,
            request_key
        ))

        connection.commit()

    finally:
        connection.close()


# -------------------------
# INITIALIZE DATABASE
# -------------------------

if __name__ == "__main__":
    initialize_database()
    print("EduPluse database initialized successfully.")
