
import uuid
import os
import sqlite3

from flask import (
    Flask, render_template, request, redirect,
    url_for, session, send_file
)
from database import (
    get_connection,
    start_ai_request,
    complete_ai_request,
    fail_ai_request
)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from datetime import date, datetime

from pypdf import PdfReader
from pptx import Presentation
from docx import Document
from ai_service import generate_questions


app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY")

if not app.secret_key:
    raise ValueError("FLASK_SECRET_KEY not found in environment.")


# -------------------------
# UPLOAD SETTINGS
# -------------------------

UPLOAD_FOLDER = "uploads/materials"
SYLLABUS_FOLDER = "uploads/syllabuses"

ALLOWED_EXTENSIONS = {"pdf", "pptx", "docx", "txt"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["SYLLABUS_FOLDER"] = SYLLABUS_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024


# -------------------------
# HELPER FUNCTIONS
# -------------------------

def allowed_file(filename):
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS
    )


def extract_text(file_path, extension):

    extracted_text = []

    if extension == "pdf":
        reader = PdfReader(file_path)

        for page in reader.pages:
            extracted_text.append(page.extract_text() or "")

    elif extension == "pptx":
        presentation = Presentation(file_path)

        for slide in presentation.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    extracted_text.append(shape.text)

    elif extension == "docx":
        document = Document(file_path)

        for paragraph in document.paragraphs:
            extracted_text.append(paragraph.text)

        for table in document.tables:
            for row in table.rows:
                extracted_text.append(
                    " | ".join(cell.text for cell in row.cells)
                )

    elif extension == "txt":
        with open(file_path, "r", encoding="utf-8") as file:
            extracted_text.append(file.read())

    return "\n".join(extracted_text)


def safe_file_path(file_path, folder):

    if not file_path:
        return None

    absolute_path = os.path.abspath(file_path)
    absolute_folder = os.path.abspath(folder)

    try:
        if os.path.commonpath(
            [absolute_path, absolute_folder]
        ) != absolute_folder:
            return None
    except ValueError:
        return None

    return absolute_path


def get_teacher_assignments(connection, teacher_id):

    return connection.execute("""
        SELECT
            classes.id AS class_id,
            classes.name AS class_name,
            subjects.id AS subject_id,
            subjects.name AS subject_name
        FROM teacher_assignments
        JOIN classes
            ON teacher_assignments.class_id = classes.id
        JOIN subjects
            ON teacher_assignments.subject_id = subjects.id
        WHERE teacher_assignments.teacher_id = ?
        ORDER BY classes.name, subjects.name
    """, (teacher_id,)).fetchall()


# -------------------------
# HOME PAGE
# -------------------------

@app.route("/")
def home():
    return render_template("index.html")


# -------------------------
# LOGIN PAGES
# -------------------------

@app.route("/teacher/login", methods=["GET", "POST"])
def teacher_login():

    if request.method == "POST":
        return login_user("teacher")

    return render_template("login.html", role="Teacher")


@app.route("/student/login", methods=["GET", "POST"])
def student_login():

    if request.method == "POST":
        return login_user("student")

    return render_template("login.html", role="Student")


# -------------------------
# LOGIN AUTHENTICATION
# -------------------------

def login_user(required_role):

    email = request.form.get("email", "").strip().lower()
    password = request.form.get("password", "")

    if not email or not password:
        return "Email and password are required."

    connection = get_connection()

    try:
        user = connection.execute(
            """
            SELECT * FROM users
            WHERE email = ? AND role = ?
            """,
            (email, required_role)
        ).fetchone()

        if user is None:
            return "Account not found. Check your email and account type."

        if not check_password_hash(user["password"], password):
            return "Incorrect password."

        session.clear()
        session["user_id"] = user["id"]
        session["name"] = user["name"]
        session["role"] = user["role"]

        if required_role == "student":
            return redirect(url_for("student_dashboard"))

        return redirect(url_for("teacher_dashboard"))

    finally:
        connection.close()


# -------------------------
# STUDENT DASHBOARD
# -------------------------

@app.route("/student/dashboard")
def student_dashboard():

    if "user_id" not in session or session.get("role") != "student":
        return redirect(url_for("student_login"))

    student_id = session["user_id"]
    connection = get_connection()

    try:
        subject_count = connection.execute("""
            SELECT COUNT(DISTINCT teacher_assignments.subject_id)
            FROM enrollments
            JOIN teacher_assignments
                ON enrollments.class_id = teacher_assignments.class_id
            WHERE enrollments.student_id = ?
        """, (student_id,)).fetchone()[0]

        attendance_data = connection.execute("""
            SELECT
                COUNT(*) AS total_lectures,
                SUM(
                    CASE WHEN status = 'Present'
                    THEN 1 ELSE 0 END
                ) AS attended
            FROM attendance
            WHERE student_id = ?
        """, (student_id,)).fetchone()

        total_lectures = attendance_data["total_lectures"]
        attended = attendance_data["attended"] or 0

        overall_attendance = (
            round((attended / total_lectures) * 100, 1)
            if total_lectures > 0 else 0
        )

        attendance_summary = connection.execute("""
            SELECT
                subjects.name AS subject_name,
                COUNT(attendance.id) AS total_lectures,
                SUM(
                    CASE WHEN attendance.status = 'Present'
                    THEN 1 ELSE 0 END
                ) AS attended,
                SUM(
                    CASE WHEN attendance.status = 'Absent'
                    THEN 1 ELSE 0 END
                ) AS missed
            FROM attendance
            JOIN attendance_sessions
                ON attendance.session_id = attendance_sessions.id
            JOIN subjects
                ON attendance_sessions.subject_id = subjects.id
            WHERE attendance.student_id = ?
            GROUP BY subjects.id
            ORDER BY subjects.name
        """, (student_id,)).fetchall()

        marks_data = connection.execute("""
            SELECT
                COALESCE(SUM(score), 0) AS total_marks,
                COALESCE(SUM(maximum_score), 0) AS maximum_marks
            FROM academic_records
            WHERE student_id = ?
        """, (student_id,)).fetchone()

        internal_marks_obtained = marks_data["total_marks"]
        internal_marks_maximum = marks_data["maximum_marks"]

        learning_materials = connection.execute("""
            SELECT
                teaching_materials.id,
                teaching_materials.title,
                teaching_materials.file_name,
                teaching_materials.created_at,
                classes.name AS class_name,
                subjects.name AS subject_name
            FROM teaching_materials
            JOIN classes
                ON teaching_materials.class_id = classes.id
            JOIN subjects
                ON teaching_materials.subject_id = subjects.id
            WHERE teaching_materials.class_id IN (
                SELECT class_id
                FROM enrollments
                WHERE student_id = ?
            )
            ORDER BY teaching_materials.created_at DESC
        """, (student_id,)).fetchall()

        upcoming_assessments = connection.execute("""
            SELECT
                assessments.id,
                assessments.title,
                assessments.start_time,
                assessments.end_time,
                classes.name AS class_name,
                subjects.name AS subject_name,
                COUNT(questions.id) AS question_count
            FROM assessments
            JOIN classes
                ON assessments.class_id = classes.id
            JOIN subjects
                ON assessments.subject_id = subjects.id
            LEFT JOIN questions
                ON assessments.id = questions.assessment_id
            WHERE assessments.status = 'published'
            AND assessments.class_id IN (
                SELECT class_id
                FROM enrollments
                WHERE student_id = ?
            )
            GROUP BY assessments.id
            ORDER BY assessments.start_time ASC
        """, (student_id,)).fetchall()

        current_time = datetime.now().isoformat(timespec="minutes")

        assessment_list = []

        for assessment in upcoming_assessments:

            assessment = dict(assessment)

            existing_attempt = connection.execute("""
                SELECT id
                FROM attempts
                WHERE student_id = ?
                AND assessment_id = ?
            """, (student_id, assessment["id"])).fetchone()

            if existing_attempt:
                assessment["availability"] = "Completed"

            elif current_time < assessment["start_time"]:
                assessment["availability"] = "Upcoming"

            elif current_time <= assessment["end_time"]:
                assessment["availability"] = "Available"

            else:
                assessment["availability"] = "Closed"

            assessment_list.append(assessment)

        return render_template(
            "student_dashboard.html",
            name=session["name"],
            enrolled_subjects=subject_count,
            overall_attendance=overall_attendance,
            attendance_summary=attendance_summary,
            internal_marks_obtained=internal_marks_obtained,
            internal_marks_maximum=internal_marks_maximum,
            learning_materials=learning_materials,
            upcoming_assessments=assessment_list
        )

    finally:
        connection.close()


# -------------------------
# STUDENT SUBJECTS
# -------------------------

@app.route("/student/subjects")
def student_subjects():

    if "user_id" not in session or session.get("role") != "student":
        return redirect(url_for("student_login"))

    student_id = session["user_id"]
    connection = get_connection()

    try:
        subjects = connection.execute("""
            SELECT
                subjects.id AS id,
                subjects.name AS name,
                subjects.id AS subject_id,
                subjects.name AS subject_name,
                classes.id AS class_id,
                classes.name AS class_name,

                (
                    SELECT COUNT(*)
                    FROM teaching_materials
                    WHERE teaching_materials.class_id = classes.id
                    AND teaching_materials.subject_id = subjects.id
                ) AS material_count,

                (
                    SELECT COUNT(*)
                    FROM assessments
                    WHERE assessments.class_id = classes.id
                    AND assessments.subject_id = subjects.id
                    AND assessments.status = 'published'
                ) AS assessment_count,

                (
                    SELECT COUNT(*)
                    FROM syllabuses
                    WHERE syllabuses.class_id = classes.id
                    AND syllabuses.subject_id = subjects.id
                ) AS syllabus_count,

                (
                    SELECT users.name
                    FROM teacher_assignments AS ta
                    JOIN users ON ta.teacher_id = users.id
                    WHERE ta.class_id = classes.id
                    AND ta.subject_id = subjects.id
                    LIMIT 1
                ) AS teacher_name

            FROM enrollments

            JOIN classes
                ON enrollments.class_id = classes.id

            JOIN teacher_assignments
                ON teacher_assignments.class_id = classes.id

            JOIN subjects
                ON teacher_assignments.subject_id = subjects.id

            WHERE enrollments.student_id = ?

            ORDER BY classes.name, subjects.name
        """, (student_id,)).fetchall()

        return render_template(
            "student_subjects.html",
            name=session["name"],
            subjects=subjects
        )

    finally:
        connection.close()


# -------------------------
# STUDENT SYLLABUS DOWNLOAD
# -------------------------

@app.route("/student/syllabus/<int:subject_id>")
def student_syllabus(subject_id):

    if "user_id" not in session or session.get("role") != "student":
        return redirect(url_for("student_login"))

    student_id = session["user_id"]
    connection = get_connection()

    try:
        syllabus = connection.execute("""
            SELECT
                syllabuses.file_path,
                syllabuses.file_name,
                syllabuses.title
            FROM syllabuses
            JOIN enrollments
                ON syllabuses.class_id = enrollments.class_id
            WHERE enrollments.student_id = ?
            AND syllabuses.subject_id = ?
        """, (student_id, subject_id)).fetchone()

        if syllabus is None:
            return "Official syllabus not found.", 404

        file_path = safe_file_path(
            syllabus["file_path"],
            app.config["SYLLABUS_FOLDER"]
        )

        if file_path is None:
            return "Invalid file location.", 403

        if not os.path.isfile(file_path):
            return "Syllabus file not found on server.", 404

        return send_file(
            file_path,
            as_attachment=True,
            download_name=syllabus["file_name"]
        )

    finally:
        connection.close()


# -------------------------
# STUDENT SYLLABUS VIEW
# -------------------------

@app.route("/student/syllabus/view/<int:subject_id>")
def view_student_syllabus(subject_id):

    if "user_id" not in session or session.get("role") != "student":
        return redirect(url_for("student_login"))

    student_id = session["user_id"]
    connection = get_connection()

    try:
        syllabus = connection.execute("""
            SELECT
                syllabuses.file_path,
                syllabuses.file_name
            FROM syllabuses
            JOIN enrollments
                ON syllabuses.class_id = enrollments.class_id
            WHERE enrollments.student_id = ?
            AND syllabuses.subject_id = ?
        """, (student_id, subject_id)).fetchone()

        if syllabus is None:
            return "Official syllabus not found.", 404

        if not syllabus["file_name"].lower().endswith(".pdf"):
            return "Only PDF syllabuses can be viewed online.", 400

        file_path = safe_file_path(
            syllabus["file_path"],
            app.config["SYLLABUS_FOLDER"]
        )

        if file_path is None:
            return "Invalid file location.", 403

        if not os.path.isfile(file_path):
            return "Syllabus file not found on server.", 404

        return send_file(
            file_path,
            mimetype="application/pdf",
            as_attachment=False,
            download_name=syllabus["file_name"]
        )

    finally:
        connection.close()



# -------------------------
# STUDENT LEARNING MATERIALS
# -------------------------

@app.route("/student/materials")
def student_materials():

    if "user_id" not in session or session.get("role") != "student":
        return redirect(url_for("student_login"))

    student_id = session["user_id"]
    subject_id = request.args.get("subject", type=int)
    connection = get_connection()

    try:
        query = """
            SELECT
                teaching_materials.id,
                teaching_materials.title,
                teaching_materials.content,
                teaching_materials.file_name,
                teaching_materials.file_path,
                teaching_materials.created_at,
                classes.name AS class_name,
                subjects.name AS subject_name

            FROM teaching_materials

            JOIN classes
                ON teaching_materials.class_id = classes.id

            JOIN subjects
                ON teaching_materials.subject_id = subjects.id

            WHERE teaching_materials.class_id IN (
                SELECT class_id
                FROM enrollments
                WHERE student_id = ?
            )
        """

        parameters = [student_id]

        if subject_id:
            query += " AND teaching_materials.subject_id = ?"
            parameters.append(subject_id)

        query += """
            ORDER BY teaching_materials.created_at DESC
        """

        learning_materials = connection.execute(
            query, parameters
        ).fetchall()

        return render_template(
            "student_materials.html",
            name=session["name"],
            learning_materials=learning_materials,
            materials=learning_materials
        )

    finally:
        connection.close()


# -------------------------
# STUDENT ASSESSMENTS LIST
# -------------------------

@app.route("/student/assessments")
def student_assessments():

    if "user_id" not in session or session.get("role") != "student":
        return redirect(url_for("student_login"))

    student_id = session["user_id"]
    subject_id = request.args.get("subject", type=int)
    connection = get_connection()

    try:
        query = """
            SELECT
                assessments.id,
                assessments.title,
                assessments.start_time,
                assessments.end_time,
                classes.name AS class_name,
                subjects.name AS subject_name,
                COUNT(questions.id) AS question_count,
                attempts.id AS attempt_id,
                attempts.score,
                attempts.submitted_at

            FROM assessments

            JOIN classes
                ON assessments.class_id = classes.id

            JOIN subjects
                ON assessments.subject_id = subjects.id

            LEFT JOIN questions
                ON assessments.id = questions.assessment_id

            LEFT JOIN attempts
                ON attempts.assessment_id = assessments.id
                AND attempts.student_id = ?

            WHERE assessments.status = 'published'

            AND assessments.class_id IN (
                SELECT class_id
                FROM enrollments
                WHERE student_id = ?
            )
        """

        parameters = [student_id, student_id]

        if subject_id:
            query += " AND assessments.subject_id = ?"
            parameters.append(subject_id)

        query += """
            GROUP BY assessments.id
            ORDER BY assessments.start_time ASC
        """

        assessments = connection.execute(
            query, parameters
        ).fetchall()

        current_time = datetime.now().isoformat(timespec="minutes")

        assessment_list = []

        for assessment in assessments:

            assessment = dict(assessment)

            if assessment["attempt_id"]:
                assessment["availability"] = "Completed"

            elif current_time < assessment["start_time"]:
                assessment["availability"] = "Upcoming"

            elif current_time <= assessment["end_time"]:
                assessment["availability"] = "Available"

            else:
                assessment["availability"] = "Closed"

            assessment_list.append(assessment)

        return render_template(
            "student_assessments.html",
            name=session["name"],
            assessments=assessment_list
        )

    finally:
        connection.close()


# -------------------------
# STUDENT ASSESSMENT ATTEMPT
# -------------------------

@app.route("/student/assessments/<int:assessment_id>", methods=["GET", "POST"])
def student_assessment_attempt(assessment_id):

    if "user_id" not in session or session.get("role") != "student":
        return redirect(url_for("student_login"))

    student_id = session["user_id"]
    connection = get_connection()

    try:
        assessment = connection.execute("""
            SELECT
                assessments.id,
                assessments.title,
                assessments.class_id,
                assessments.subject_id,
                assessments.start_time,
                assessments.end_time,
                subjects.name AS subject_name
            FROM assessments
            JOIN subjects
                ON assessments.subject_id = subjects.id
            WHERE assessments.id = ?
            AND assessments.status = 'published'
            AND assessments.class_id IN (
                SELECT class_id
                FROM enrollments
                WHERE student_id = ?
            )
        """, (assessment_id, student_id)).fetchone()

        if assessment is None:
            return "Assessment not found.", 404

        existing_attempt = connection.execute("""
            SELECT id, score, submitted_at
            FROM attempts
            WHERE student_id = ?
            AND assessment_id = ?
        """, (student_id, assessment_id)).fetchone()

        if existing_attempt:
            return render_template(
                "student_assessment_result.html",
                assessment=assessment,
                attempt=existing_attempt,
                already_submitted=True
            )

        try:
            start_time = datetime.fromisoformat(assessment["start_time"])
            end_time = datetime.fromisoformat(assessment["end_time"])
        except (ValueError, TypeError):
            return "Assessment schedule is invalid.", 400

        current_time = datetime.now()

        if current_time < start_time:
            return "This assessment is not available yet.", 403

        if current_time > end_time:
            return "The assessment submission period has ended.", 403

        questions = connection.execute("""
            SELECT
                id,
                question,
                option_a,
                option_b,
                option_c,
                option_d,
                correct_answer
            FROM questions
            WHERE assessment_id = ?
            ORDER BY id
        """, (assessment_id,)).fetchall()

        if not questions:
            return "This assessment has no questions.", 400

        if request.method == "POST":

            answers = []

            for question in questions:

                selected = request.form.get(
                    f"question_{question['id']}", ""
                ).strip().upper()

                if selected not in ["A", "B", "C", "D"]:
                    return "Please answer every question.", 400

                answers.append((question["id"], selected))

            try:
                connection.execute("BEGIN")

                cursor = connection.execute("""
                    INSERT INTO attempts
                    (student_id, assessment_id, score)
                    VALUES (?, ?, 0)
                """, (student_id, assessment_id))

                attempt_id = cursor.lastrowid
                score = 0

                for question, (_, selected) in zip(questions, answers):

                    correct_answer = str(
                        question["correct_answer"]
                    ).strip()

                    correct_letter = correct_answer.upper().replace(
                        "OPTION ", ""
                    )

                    if correct_letter in ["A", "B", "C", "D"]:
                        is_correct = selected == correct_letter

                    else:
                        option_column = f"option_{selected.lower()}"

                        is_correct = (
                            correct_answer.casefold()
                            == str(question[option_column]).strip().casefold()
                        )

                    if is_correct:
                        score += 1

                    connection.execute("""
                        INSERT INTO responses
                        (attempt_id, question_id, selected_answer, is_correct)
                        VALUES (?, ?, ?, ?)
                    """, (
                        attempt_id,
                        question["id"],
                        selected,
                        int(is_correct)
                    ))

                connection.execute("""
                    UPDATE attempts
                    SET score = ?
                    WHERE id = ?
                """, (score, attempt_id))

                connection.commit()

                attempt = connection.execute("""
                    SELECT id, score, submitted_at
                    FROM attempts
                    WHERE id = ?
                """, (attempt_id,)).fetchone()

                return render_template(
                    "student_assessment_result.html",
                    assessment=assessment,
                    attempt=attempt,
                    already_submitted=False
                )

            except sqlite3.IntegrityError:
                connection.rollback()
                return "This assessment has already been submitted.", 409

            except Exception:
                connection.rollback()
                raise

        return render_template(
            "student_assessment_attempt.html",
            assessment=assessment,
            questions=questions
        )

    finally:
        connection.close()


# -------------------------
# STUDENT MATERIAL DOWNLOAD
# -------------------------

@app.route("/student/materials/download/<int:material_id>")
def download_student_material(material_id):

    if "user_id" not in session or session.get("role") != "student":
        return redirect(url_for("student_login"))

    student_id = session["user_id"]
    connection = get_connection()

    try:
        material = connection.execute("""
            SELECT
                teaching_materials.file_path,
                teaching_materials.file_name
            FROM teaching_materials
            WHERE teaching_materials.id = ?
            AND teaching_materials.class_id IN (
                SELECT class_id
                FROM enrollments
                WHERE student_id = ?
            )
        """, (material_id, student_id)).fetchone()

        if material is None:
            return "Learning material not found.", 404

        file_path = safe_file_path(
            material["file_path"],
            app.config["UPLOAD_FOLDER"]
        )

        if file_path is None:
            return "Invalid file location.", 403

        if not os.path.isfile(file_path):
            return "File not found on server.", 404

        return send_file(
            file_path,
            as_attachment=True,
            download_name=material["file_name"]
        )

    finally:
        connection.close()


# -------------------------
# STUDENT MATERIAL VIEW
# -------------------------

@app.route("/student/materials/view/<int:material_id>")
def view_student_material(material_id):

    if "user_id" not in session or session.get("role") != "student":
        return redirect(url_for("student_login"))

    student_id = session["user_id"]
    connection = get_connection()

    try:
        material = connection.execute("""
            SELECT
                file_path,
                file_name
            FROM teaching_materials
            WHERE id = ?
            AND class_id IN (
                SELECT class_id
                FROM enrollments
                WHERE student_id = ?
            )
        """, (material_id, student_id)).fetchone()

        if material is None:
            return "Learning material not found.", 404

        if not material["file_name"].lower().endswith(".pdf"):
            return (
                "Only PDF materials can be viewed online. "
                "Please download this file.",
                400
            )

        file_path = safe_file_path(
            material["file_path"],
            app.config["UPLOAD_FOLDER"]
        )

        if file_path is None:
            return "Invalid file location.", 403

        if not os.path.isfile(file_path):
            return "File not found on server.", 404

        return send_file(
            file_path,
            mimetype="application/pdf",
            as_attachment=False,
            download_name=material["file_name"]
        )

    finally:
        connection.close()


# -------------------------
# STUDENT ATTENDANCE
# -------------------------

@app.route("/student/attendance")
def student_attendance():

    if "user_id" not in session or session.get("role") != "student":
        return redirect(url_for("student_login"))

    student_id = session["user_id"]
    connection = get_connection()

    try:
        attendance_summary = connection.execute("""
            SELECT
                subjects.name AS subject_name,
                COUNT(attendance.id) AS total_lectures,
                SUM(
                    CASE
                        WHEN attendance.status = 'Present'
                        THEN 1 ELSE 0
                    END
                ) AS attended,
                SUM(
                    CASE
                        WHEN attendance.status = 'Absent'
                        THEN 1 ELSE 0
                    END
                ) AS missed
            FROM attendance
            JOIN attendance_sessions
                ON attendance.session_id = attendance_sessions.id
            JOIN subjects
                ON attendance_sessions.subject_id = subjects.id
            WHERE attendance.student_id = ?
            GROUP BY subjects.id
            ORDER BY subjects.name
        """, (student_id,)).fetchall()

        return render_template(
            "student_attendance.html",
            attendance_summary=attendance_summary
        )

    finally:
        connection.close()


# -------------------------
# STUDENT INTERNAL MARKS
# -------------------------

@app.route("/student/marks")
def student_marks():

    if "user_id" not in session or session.get("role") != "student":
        return redirect(url_for("student_login"))

    student_id = session["user_id"]
    connection = get_connection()

    try:
        marks = connection.execute("""
            SELECT
                subjects.name AS subject_name,
                academic_records.record_type,
                academic_records.score,
                academic_records.maximum_score,
                academic_records.recorded_at
            FROM academic_records
            JOIN subjects
                ON academic_records.subject_id = subjects.id
            WHERE academic_records.student_id = ?
            ORDER BY
                subjects.name,
                academic_records.recorded_at DESC
        """, (student_id,)).fetchall()

        return render_template(
            "student_marks.html",
            marks=marks,
            name=session["name"]
        )

    finally:
        connection.close()



# -------------------------
# STUDENT TOPIC ANALYSIS
# -------------------------

@app.route("/student/topic-analysis")
def student_topic_analysis():

    if "user_id" not in session or session.get("role") != "student":
        return redirect(url_for("student_login"))

    student_id = session["user_id"]
    connection = get_connection()

    try:
        topic_analysis = connection.execute("""
            SELECT
                questions.concept AS concept,
                COUNT(responses.id) AS attempted,
                SUM(
                    CASE
                        WHEN responses.is_correct = 1
                        THEN 1 ELSE 0
                    END
                ) AS correct
            FROM responses

            JOIN attempts
                ON responses.attempt_id = attempts.id

            JOIN questions
                ON responses.question_id = questions.id

            WHERE attempts.student_id = ?
            AND questions.concept IS NOT NULL
            AND TRIM(questions.concept) != ''

            GROUP BY TRIM(questions.concept)
            ORDER BY concept
        """, (student_id,)).fetchall()

        analysis = []

        total_questions_attempted = 0
        total_correct_answers = 0

        for topic in topic_analysis:

            topic = dict(topic)

            attempted = topic["attempted"] or 0
            correct = topic["correct"] or 0

            accuracy = (
                round((correct / attempted) * 100, 1)
                if attempted > 0 else 0
            )

            topic["attempted"] = attempted
            topic["correct"] = correct
            topic["accuracy"] = accuracy

            if accuracy >= 75:
                topic["status"] = "Strong"

            elif accuracy >= 50:
                topic["status"] = "Needs Improvement"

            else:
                topic["status"] = "Needs Attention"

            total_questions_attempted += attempted
            total_correct_answers += correct

            analysis.append(topic)

        overall_accuracy = (
            round(
                (total_correct_answers / total_questions_attempted) * 100,
                1
            )
            if total_questions_attempted > 0 else 0
        )

        strong_topics = [
            topic for topic in analysis
            if topic["accuracy"] >= 75
        ]

        weak_topics = [
            topic for topic in analysis
            if topic["accuracy"] < 50
        ]

        return render_template(
            "student_topic_analysis.html",
            name=session["name"],
            topic_analysis=analysis,
            total_questions_attempted=total_questions_attempted,
            overall_accuracy=overall_accuracy,
            strong_topics=strong_topics,
            weak_topics=weak_topics
        )

    finally:
        connection.close()



@app.route("/teacher/dashboard")
def teacher_dashboard():

    if "user_id" not in session or session.get("role") != "teacher":
        return redirect(url_for("teacher_login"))

    teacher_id = session["user_id"]
    connection = get_connection()

    try:
        total_classes = connection.execute("""
            SELECT COUNT(DISTINCT class_id)
            FROM teacher_assignments
            WHERE teacher_id = ?
        """, (teacher_id,)).fetchone()[0]

        total_subjects = connection.execute("""
            SELECT COUNT(DISTINCT subject_id)
            FROM teacher_assignments
            WHERE teacher_id = ?
        """, (teacher_id,)).fetchone()[0]

        total_students = connection.execute("""
            SELECT COUNT(DISTINCT enrollments.student_id)
            FROM enrollments
            JOIN teacher_assignments
                ON enrollments.class_id = teacher_assignments.class_id
            WHERE teacher_assignments.teacher_id = ?
        """, (teacher_id,)).fetchone()[0]

        active_assessments = connection.execute("""
            SELECT COUNT(*)
            FROM assessments
            WHERE teacher_id = ?
            AND status = 'published'
        """, (teacher_id,)).fetchone()[0]

        material_count = connection.execute("""
            SELECT COUNT(*)
            FROM teaching_materials
            WHERE teacher_id = ?
        """, (teacher_id,)).fetchone()[0]

        recent_assessments = connection.execute("""
            SELECT
                assessments.id,
                assessments.title,
                assessments.status,
                assessments.start_time,
                classes.name AS class_name,
                subjects.name AS subject_name
            FROM assessments
            JOIN classes
                ON assessments.class_id = classes.id
            JOIN subjects
                ON assessments.subject_id = subjects.id
            WHERE assessments.teacher_id = ?
            ORDER BY assessments.id DESC
            LIMIT 5
        """, (teacher_id,)).fetchall()

        return render_template(
            "teacher_dashboard.html",
            name=session["name"],
            total_classes=total_classes,
            total_subjects=total_subjects,
            total_students=total_students,
            active_assessments=active_assessments,
            material_count=material_count,
            recent_assessments=recent_assessments
        )

    finally:
        connection.close()



# -------------------------
# TEACHER SUBJECTS
# -------------------------

@app.route("/teacher/subjects")
def teacher_subjects():

    if "user_id" not in session or session.get("role") != "teacher":
        return redirect(url_for("teacher_login"))

    teacher_id = session["user_id"]
    connection = get_connection()

    try:
        subjects = connection.execute("""
            SELECT
                classes.id AS class_id,
                classes.name AS class_name,
                subjects.id AS subject_id,
                subjects.name AS subject_name,

                (
                    SELECT COUNT(*)
                    FROM enrollments
                    WHERE enrollments.class_id = classes.id
                ) AS student_count,

                (
                    SELECT COUNT(*)
                    FROM teaching_materials
                    WHERE teaching_materials.class_id = classes.id
                    AND teaching_materials.subject_id = subjects.id
                ) AS material_count,

                (
                    SELECT COUNT(*)
                    FROM assessments
                    WHERE assessments.class_id = classes.id
                    AND assessments.subject_id = subjects.id
                ) AS assessment_count,

                (
                    SELECT COUNT(*)
                    FROM syllabuses
                    WHERE syllabuses.class_id = classes.id
                    AND syllabuses.subject_id = subjects.id
                ) AS syllabus_count

            FROM teacher_assignments

            JOIN classes
                ON teacher_assignments.class_id = classes.id

            JOIN subjects
                ON teacher_assignments.subject_id = subjects.id

            WHERE teacher_assignments.teacher_id = ?

            ORDER BY classes.name, subjects.name
        """, (teacher_id,)).fetchall()

        syllabus_files = connection.execute("""
            SELECT
                syllabuses.id,
                syllabuses.title,
                syllabuses.file_name,
                syllabuses.uploaded_at,
                classes.name AS class_name,
                subjects.name AS subject_name

            FROM syllabuses

            JOIN classes
                ON syllabuses.class_id = classes.id

            JOIN subjects
                ON syllabuses.subject_id = subjects.id

            WHERE syllabuses.teacher_id = ?

            ORDER BY syllabuses.uploaded_at DESC
        """, (teacher_id,)).fetchall()

        return render_template(
            "teacher_subjects.html",
            name=session["name"],
            subjects=subjects,
            syllabus_files=syllabus_files
        )

    finally:
        connection.close()

# -------------------------
# TEACHER SYLLABUS UPLOAD
# -------------------------

@app.route("/teacher/syllabus/upload", methods=["POST"])
def upload_teacher_syllabus():

    if "user_id" not in session or session.get("role") != "teacher":
        return redirect(url_for("teacher_login"))

    teacher_id = session["user_id"]

    title = request.form.get("title", "").strip()
    class_id = request.form.get("class_id", type=int)
    subject_id = request.form.get("subject_id", type=int)
    uploaded_file = request.files.get("syllabus_file")

    if not title:
        return "Syllabus title is required.", 400

    if not class_id or not subject_id:
        return "Please select a class and subject.", 400

    if not uploaded_file or not uploaded_file.filename:
        return "Please select a syllabus file.", 400

    if not allowed_file(uploaded_file.filename):
        return "Unsupported file format.", 400

    connection = get_connection()

    try:
        assignment = connection.execute("""
            SELECT id
            FROM teacher_assignments
            WHERE teacher_id = ?
            AND class_id = ?
            AND subject_id = ?
        """, (teacher_id, class_id, subject_id)).fetchone()

        if assignment is None:
            return "You are not assigned to this class and subject.", 403

        original_filename = secure_filename(uploaded_file.filename)

        if "." not in original_filename:
            return "Invalid filename.", 400

        extension = original_filename.rsplit(".", 1)[1].lower()
        stored_filename = f"{uuid.uuid4().hex}.{extension}"

        os.makedirs(app.config["SYLLABUS_FOLDER"], exist_ok=True)

        file_path = os.path.join(
            app.config["SYLLABUS_FOLDER"],
            stored_filename
        )

        uploaded_file.save(file_path)

        try:
            existing = connection.execute("""
                SELECT id, file_path
                FROM syllabuses
                WHERE class_id = ?
                AND subject_id = ?
            """, (class_id, subject_id)).fetchone()

            if existing:

                old_file_path = existing["file_path"]

                connection.execute("""
                    UPDATE syllabuses
                    SET
                        teacher_id = ?,
                        title = ?,
                        file_name = ?,
                        file_path = ?,
                        uploaded_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                """, (
                    teacher_id,
                    title,
                    original_filename,
                    file_path,
                    existing["id"]
                ))

            else:

                old_file_path = None

                connection.execute("""
                    INSERT INTO syllabuses
                    (
                        teacher_id,
                        class_id,
                        subject_id,
                        title,
                        file_name,
                        file_path
                    )
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (
                    teacher_id,
                    class_id,
                    subject_id,
                    title,
                    original_filename,
                    file_path
                ))

            connection.commit()

            if old_file_path:
                old_file_path = safe_file_path(
                    old_file_path,
                    app.config["SYLLABUS_FOLDER"]
                )

                if old_file_path and os.path.isfile(old_file_path):
                    if old_file_path != os.path.abspath(file_path):
                        os.remove(old_file_path)

            return redirect(url_for("teacher_subjects"))

        except Exception:
            connection.rollback()

            if os.path.exists(file_path):
                os.remove(file_path)

            raise

    finally:
        connection.close()


# -------------------------
# TEACHER SYLLABUS DOWNLOAD
# -------------------------

@app.route("/teacher/syllabus/download/<int:syllabus_id>")
def download_teacher_syllabus(syllabus_id):

    if "user_id" not in session or session.get("role") != "teacher":
        return redirect(url_for("teacher_login"))

    connection = get_connection()

    try:
        syllabus = connection.execute("""
            SELECT file_path, file_name
            FROM syllabuses
            WHERE id = ?
            AND teacher_id = ?
        """, (syllabus_id, session["user_id"])).fetchone()

        if syllabus is None:
            return "Syllabus not found.", 404

        file_path = safe_file_path(
            syllabus["file_path"],
            app.config["SYLLABUS_FOLDER"]
        )

        if file_path is None:
            return "Invalid file location.", 403

        if not os.path.isfile(file_path):
            return "Syllabus file not found.", 404

        return send_file(
            file_path,
            as_attachment=True,
            download_name=syllabus["file_name"]
        )

    finally:
        connection.close()



@app.route("/teacher/class-analysis")
def teacher_class_analysis():

    if "user_id" not in session or session.get("role") != "teacher":
        return redirect(url_for("teacher_login"))

    teacher_id = session["user_id"]
    selected_class_id = request.args.get("class_id", type=int)

    connection = get_connection()

    try:
        # Get classes assigned to this teacher
        classes = connection.execute("""
            SELECT DISTINCT
                classes.id,
                classes.name
            FROM teacher_assignments
            JOIN classes
                ON teacher_assignments.class_id = classes.id
            WHERE teacher_assignments.teacher_id = ?
            ORDER BY classes.name
        """, (teacher_id,)).fetchall()

        valid_class_ids = [item["id"] for item in classes]

        if selected_class_id not in valid_class_ids:
            selected_class_id = None

        # Class and subject performance
        class_analysis = connection.execute("""
            SELECT
                classes.id AS class_id,
                classes.name AS class_name,
                subjects.id AS subject_id,
                subjects.name AS subject_name,

                (
                    SELECT COUNT(DISTINCT enrollments.student_id)
                    FROM enrollments
                    WHERE enrollments.class_id = classes.id
                ) AS student_count,

                (
                    SELECT COUNT(*)
                    FROM attendance_sessions
                    WHERE attendance_sessions.class_id = classes.id
                    AND attendance_sessions.subject_id = subjects.id
                ) AS total_lectures,

                (
                    SELECT COUNT(*)
                    FROM attendance
                    JOIN attendance_sessions
                        ON attendance.session_id = attendance_sessions.id
                    WHERE attendance_sessions.class_id = classes.id
                    AND attendance_sessions.subject_id = subjects.id
                    AND attendance.status = 'Present'
                ) AS present_records,

                (
                    SELECT COUNT(DISTINCT attempts.id)
                    FROM assessments
                    JOIN attempts
                        ON attempts.assessment_id = assessments.id
                    WHERE assessments.class_id = classes.id
                    AND assessments.subject_id = subjects.id
                    AND assessments.status = 'published'
                ) AS total_attempts,

                (
                    SELECT AVG(attempts.score)
                    FROM assessments
                    JOIN attempts
                        ON attempts.assessment_id = assessments.id
                    WHERE assessments.class_id = classes.id
                    AND assessments.subject_id = subjects.id
                    AND assessments.status = 'published'
                ) AS average_score

            FROM teacher_assignments

            JOIN classes
                ON teacher_assignments.class_id = classes.id

            JOIN subjects
                ON teacher_assignments.subject_id = subjects.id

            WHERE teacher_assignments.teacher_id = ?
            AND (? IS NULL OR classes.id = ?)

            GROUP BY classes.id, subjects.id

            ORDER BY classes.name, subjects.name
        """, (
            teacher_id,
            selected_class_id,
            selected_class_id
        )).fetchall()

        # Topic-wise analysis for the selected class
        topic_analysis = []

        if selected_class_id is not None:

            topic_analysis = connection.execute("""
                SELECT
                    subjects.name AS subject_name,

                    COALESCE(
                        NULLIF(TRIM(questions.concept), ''),
                        'Uncategorized'
                    ) AS topic,

                    COUNT(DISTINCT attempts.id) AS total_attempts,

                    COUNT(responses.id) AS total_responses,

                    SUM(
                        CASE
                            WHEN responses.is_correct = 1 THEN 1
                            ELSE 0
                        END
                    ) AS correct_responses,

                    ROUND(
                        100.0 * SUM(
                            CASE
                                WHEN responses.is_correct = 1 THEN 1
                                ELSE 0
                            END
                        ) / NULLIF(COUNT(responses.id), 0),
                        1
                    ) AS accuracy

                FROM assessments

                JOIN subjects
                    ON assessments.subject_id = subjects.id

                JOIN questions
                    ON questions.assessment_id = assessments.id

                JOIN attempts
                    ON attempts.assessment_id = assessments.id

                JOIN responses
                    ON responses.attempt_id = attempts.id
                    AND responses.question_id = questions.id

                WHERE assessments.teacher_id = ?
                AND assessments.class_id = ?
                AND assessments.status = 'published'

                GROUP BY subjects.id, questions.concept

                ORDER BY accuracy ASC, subjects.name, topic
            """, (
                teacher_id,
                selected_class_id
            )).fetchall()

        return render_template(
            "teacher_class_analysis.html",
            name=session["name"],
            classes=classes,
            selected_class_id=selected_class_id,
            class_analysis=class_analysis,
            topic_analysis=topic_analysis
        )

    finally:
        connection.close()



# -------------------------
# REGISTRATION
# -------------------------

@app.route("/register", methods=["GET", "POST"])
def register():

    if request.method == "POST":

        name = request.form.get("name", "").strip()
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        role = request.form.get("role", "")

        if not name or not email or not password:
            return "All fields are required."

        if role not in ["teacher", "student"]:
            return "Invalid account type."

        if len(password) < 8:
            return "Password must contain at least 8 characters."

        hashed_password = generate_password_hash(password)
        connection = get_connection()

        try:
            connection.execute("""
                INSERT INTO users (name, email, password, role)
                VALUES (?, ?, ?, ?)
            """, (name, email, hashed_password, role))

            connection.commit()

            if role == "student":
                return redirect(url_for("student_login"))

            return redirect(url_for("teacher_login"))

        except sqlite3.IntegrityError:
            return "This email is already registered."

        finally:
            connection.close()

    return render_template("register.html")


# -------------------------
# LOGOUT
# -------------------------

@app.route("/logout")
def logout():

    session.clear()
    return redirect(url_for("home"))


# -------------------------
# DATABASE TEST
# -------------------------

@app.route("/db-test")
def db_test():

    connection = get_connection()

    try:
        connection.execute("SELECT 1")
        return "EduPluse database connected successfully!"

    finally:
        connection.close()


# -------------------------
# TEACHER ATTENDANCE
# -------------------------

@app.route("/teacher/attendance")
def teacher_attendance():

    if "user_id" not in session or session.get("role") != "teacher":
        return redirect(url_for("teacher_login"))

    connection = get_connection()

    try:
        teacher_id = session["user_id"]

        assignments = get_teacher_assignments(connection, teacher_id)

        selected_assignment = None
        students = []

        class_id = request.args.get("class_id", type=int)
        subject_id = request.args.get("subject_id", type=int)

        if class_id is not None or subject_id is not None:

            if class_id is None or subject_id is None:
                return "Both class and subject are required.", 400

            selected_assignment = connection.execute("""
                SELECT
                    classes.id AS class_id,
                    classes.name AS class_name,
                    subjects.id AS subject_id,
                    subjects.name AS subject_name
                FROM teacher_assignments
                JOIN classes
                    ON teacher_assignments.class_id = classes.id
                JOIN subjects
                    ON teacher_assignments.subject_id = subjects.id
                WHERE teacher_assignments.teacher_id = ?
                AND teacher_assignments.class_id = ?
                AND teacher_assignments.subject_id = ?
            """, (teacher_id, class_id, subject_id)).fetchone()

            if selected_assignment is None:
                return "You are not assigned to this class and subject.", 403

            students = connection.execute("""
                SELECT
                    users.id,
                    users.name,
                    users.email
                FROM enrollments
                JOIN users
                    ON enrollments.student_id = users.id
                WHERE enrollments.class_id = ?
                AND users.role = 'student'
                ORDER BY users.name
            """, (class_id,)).fetchall()

        return render_template(
            "teacher_attendance.html",
            assignments=assignments,
            selected_assignment=selected_assignment,
            students=students
        )

    finally:
        connection.close()


@app.route("/teacher/attendance/save", methods=["POST"])
def save_attendance():

    if "user_id" not in session or session.get("role") != "teacher":
        return redirect(url_for("teacher_login"))

    teacher_id = session["user_id"]

    class_id = request.form.get("class_id", type=int)
    subject_id = request.form.get("subject_id", type=int)
    lecture_date = request.form.get("lecture_date", "")
    lecture_number = request.form.get("lecture_number", type=int)

    if not class_id or not subject_id:
        return "Invalid class or subject.", 400

    if not lecture_number or lecture_number < 1:
        return "Lecture number must be at least 1.", 400

    try:
        date.fromisoformat(lecture_date)
    except (ValueError, TypeError):
        return "Please enter a valid lecture date.", 400

    connection = get_connection()

    try:
        assignment = connection.execute("""
            SELECT id
            FROM teacher_assignments
            WHERE teacher_id = ?
            AND class_id = ?
            AND subject_id = ?
        """, (teacher_id, class_id, subject_id)).fetchone()

        if assignment is None:
            return "You are not assigned to this class and subject.", 403

        students = connection.execute("""
            SELECT users.id
            FROM enrollments
            JOIN users
                ON enrollments.student_id = users.id
            WHERE enrollments.class_id = ?
            AND users.role = 'student'
        """, (class_id,)).fetchall()

        if not students:
            return "No students are enrolled in this class.", 400

        attendance_records = []

        for student in students:
            student_id = student["id"]
            status = request.form.get(f"student_{student_id}")

            if status not in ["Present", "Absent"]:
                return "Please mark attendance for every student.", 400

            attendance_records.append((student_id, status))

        cursor = connection.execute("""
            INSERT INTO attendance_sessions
            (teacher_id, class_id, subject_id, lecture_date, lecture_number)
            VALUES (?, ?, ?, ?, ?)
        """, (
            teacher_id,
            class_id,
            subject_id,
            lecture_date,
            lecture_number
        ))

        session_id = cursor.lastrowid

        for student_id, status in attendance_records:
            connection.execute("""
                INSERT INTO attendance (session_id, student_id, status)
                VALUES (?, ?, ?)
            """, (session_id, student_id, status))

        connection.commit()

        return redirect(url_for(
            "teacher_attendance",
            class_id=class_id,
            subject_id=subject_id,
            saved=1
        ))

    except sqlite3.IntegrityError:
        connection.rollback()
        return "Attendance for this timetable lecture has already been recorded.", 409

    finally:
        connection.close()


@app.route("/teacher/attendance/history")
def teacher_attendance_history():

    if "user_id" not in session or session.get("role") != "teacher":
        return redirect(url_for("teacher_login"))

    teacher_id = session["user_id"]

    class_id = request.args.get("class_id", type=int)
    subject_id = request.args.get("subject_id", type=int)

    if not class_id or not subject_id:
        return "Class and subject are required.", 400

    connection = get_connection()

    try:
        assignment = connection.execute("""
            SELECT
                classes.name AS class_name,
                subjects.name AS subject_name
            FROM teacher_assignments
            JOIN classes
                ON teacher_assignments.class_id = classes.id
            JOIN subjects
                ON teacher_assignments.subject_id = subjects.id
            WHERE teacher_assignments.teacher_id = ?
            AND teacher_assignments.class_id = ?
            AND teacher_assignments.subject_id = ?
        """, (teacher_id, class_id, subject_id)).fetchone()

        if assignment is None:
            return "You are not assigned to this class and subject.", 403

        attendance_sessions = connection.execute("""
            SELECT
                attendance_sessions.id,
                attendance_sessions.lecture_date,
                attendance_sessions.lecture_number,
                COUNT(attendance.id) AS total_students,
                SUM(
                    CASE WHEN attendance.status = 'Present'
                    THEN 1 ELSE 0 END
                ) AS present_count,
                SUM(
                    CASE WHEN attendance.status = 'Absent'
                    THEN 1 ELSE 0 END
                ) AS absent_count
            FROM attendance_sessions
            LEFT JOIN attendance
                ON attendance_sessions.id = attendance.session_id
            WHERE attendance_sessions.teacher_id = ?
            AND attendance_sessions.class_id = ?
            AND attendance_sessions.subject_id = ?
            GROUP BY attendance_sessions.id
            ORDER BY
                attendance_sessions.lecture_date DESC,
                attendance_sessions.lecture_number DESC
        """, (teacher_id, class_id, subject_id)).fetchall()

        return render_template(
            "teacher_attendance_history.html",
            assignment=assignment,
            class_id=class_id,
            subject_id=subject_id,
            attendance_sessions=attendance_sessions
        )

    finally:
        connection.close()


@app.route("/teacher/attendance/session/<int:session_id>")
def teacher_attendance_session(session_id):

    if "user_id" not in session or session.get("role") != "teacher":
        return redirect(url_for("teacher_login"))

    teacher_id = session["user_id"]
    connection = get_connection()

    try:
        session_info = connection.execute("""
            SELECT
                attendance_sessions.id,
                attendance_sessions.class_id,
                attendance_sessions.subject_id,
                attendance_sessions.lecture_date,
                attendance_sessions.lecture_number,
                classes.name AS class_name,
                subjects.name AS subject_name
            FROM attendance_sessions
            JOIN classes
                ON attendance_sessions.class_id = classes.id
            JOIN subjects
                ON attendance_sessions.subject_id = subjects.id
            WHERE attendance_sessions.id = ?
            AND attendance_sessions.teacher_id = ?
        """, (session_id, teacher_id)).fetchone()

        if session_info is None:
            return "Attendance session not found.", 404

        students = connection.execute("""
            SELECT
                users.name,
                users.email,
                attendance.status
            FROM attendance
            JOIN users
                ON attendance.student_id = users.id
            WHERE attendance.session_id = ?
            ORDER BY users.name
        """, (session_id,)).fetchall()

        return render_template(
            "teacher_attendance_session.html",
            session_info=session_info,
            students=students
        )

    finally:
        connection.close()


# -------------------------
# TEACHER INTERNAL MARKS
# -------------------------

@app.route("/teacher/marks", methods=["GET", "POST"])
def teacher_marks():

    if "user_id" not in session or session.get("role") != "teacher":
        return redirect(url_for("teacher_login"))

    teacher_id = session["user_id"]
    connection = get_connection()

    try:
        assignments = get_teacher_assignments(connection, teacher_id)

        class_id = request.values.get("class_id", type=int)
        subject_id = request.values.get("subject_id", type=int)

        selected_assignment = None
        students = []
        existing_marks = {}
        record_type = request.values.get("record_type", "").strip()
        maximum_score = request.values.get("maximum_score", type=float)

        if class_id is not None or subject_id is not None:

            if class_id is None or subject_id is None:
                return "Select both class and subject.", 400

            selected_assignment = connection.execute("""
                SELECT
                    classes.id AS class_id,
                    classes.name AS class_name,
                    subjects.id AS subject_id,
                    subjects.name AS subject_name
                FROM teacher_assignments
                JOIN classes
                    ON teacher_assignments.class_id = classes.id
                JOIN subjects
                    ON teacher_assignments.subject_id = subjects.id
                WHERE teacher_assignments.teacher_id = ?
                AND teacher_assignments.class_id = ?
                AND teacher_assignments.subject_id = ?
            """, (teacher_id, class_id, subject_id)).fetchone()

            if selected_assignment is None:
                return "You are not assigned to this subject.", 403

            students = connection.execute("""
                SELECT
                    users.id,
                    users.name,
                    users.email
                FROM enrollments
                JOIN users
                    ON enrollments.student_id = users.id
                WHERE enrollments.class_id = ?
                AND users.role = 'student'
                ORDER BY users.name
            """, (class_id,)).fetchall()

            if record_type:

                saved_records = connection.execute("""
                    SELECT
                        academic_records.student_id,
                        academic_records.score,
                        academic_records.maximum_score
                    FROM academic_records
                    JOIN enrollments
                        ON academic_records.student_id = enrollments.student_id
                    WHERE academic_records.subject_id = ?
                    AND academic_records.record_type = ?
                    AND enrollments.class_id = ?
                """, (subject_id, record_type, class_id)).fetchall()

                for record in saved_records:
                    existing_marks[record["student_id"]] = record["score"]

                if saved_records and maximum_score is None:
                    maximum_score = saved_records[0]["maximum_score"]

        if request.method == "POST":

            if selected_assignment is None:
                return "Invalid class or subject.", 400

            if not record_type:
                return "Examination type is required.", 400

            if maximum_score is None or maximum_score <= 0:
                return "Maximum marks must be greater than zero.", 400

            try:
                connection.execute("BEGIN")

                for student in students:

                    score = request.form.get(
                        f"score_{student['id']}", type=float
                    )

                    if score is None or score < 0 or score > maximum_score:
                        raise ValueError(
                            f"Invalid marks for {student['name']}."
                        )

                    existing = connection.execute("""
                        SELECT id
                        FROM academic_records
                        WHERE student_id = ?
                        AND subject_id = ?
                        AND record_type = ?
                    """, (
                        student["id"],
                        subject_id,
                        record_type
                    )).fetchone()

                    if existing:
                        connection.execute("""
                            UPDATE academic_records
                            SET score = ?,
                                maximum_score = ?,
                                recorded_at = CURRENT_TIMESTAMP
                            WHERE id = ?
                        """, (
                            score,
                            maximum_score,
                            existing["id"]
                        ))

                    else:
                        connection.execute("""
                            INSERT INTO academic_records
                            (
                                student_id,
                                subject_id,
                                record_type,
                                score,
                                maximum_score
                            )
                            VALUES (?, ?, ?, ?, ?)
                        """, (
                            student["id"],
                            subject_id,
                            record_type,
                            score,
                            maximum_score
                        ))

                connection.commit()

                return redirect(url_for(
                    "teacher_marks",
                    class_id=class_id,
                    subject_id=subject_id,
                    record_type=record_type
                ))

            except ValueError as error:
                connection.rollback()
                return str(error), 400

            except Exception:
                connection.rollback()
                raise

        return render_template(
            "teacher_marks.html",
            assignments=assignments,
            selected_assignment=selected_assignment,
            students=students,
            record_type=record_type,
            maximum_score=maximum_score,
            existing_marks=existing_marks
        )

    finally:
        connection.close()


# -------------------------
# TEACHER LEARNING MATERIALS
# -------------------------

@app.route("/teacher/materials", methods=["GET", "POST"])
def teacher_materials():

    if "user_id" not in session or session.get("role") != "teacher":
        return redirect(url_for("teacher_login"))

    teacher_id = session["user_id"]
    connection = get_connection()

    try:
        assignments = get_teacher_assignments(connection, teacher_id)

        if request.method == "POST":

            title = request.form.get("title", "").strip()
            class_id = request.form.get("class_id", type=int)
            subject_id = request.form.get("subject_id", type=int)
            uploaded_file = request.files.get("material_file")

            if not title:
                return "Material title is required.", 400

            if not class_id or not subject_id:
                return "Please select a class and subject.", 400

            if not uploaded_file or not uploaded_file.filename:
                return "Please select a file.", 400

            if not allowed_file(uploaded_file.filename):
                return "Unsupported file format.", 400

            assignment = connection.execute("""
                SELECT id
                FROM teacher_assignments
                WHERE teacher_id = ?
                AND class_id = ?
                AND subject_id = ?
            """, (teacher_id, class_id, subject_id)).fetchone()

            if assignment is None:
                return "You are not assigned to this class and subject.", 403

            original_filename = secure_filename(uploaded_file.filename)

            if "." not in original_filename:
                return "Invalid filename.", 400

            extension = original_filename.rsplit(".", 1)[1].lower()
            stored_filename = f"{uuid.uuid4().hex}.{extension}"

            os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

            file_path = os.path.join(
                app.config["UPLOAD_FOLDER"],
                stored_filename
            )

            try:
                uploaded_file.save(file_path)

                extracted_content = extract_text(file_path, extension)

                if not extracted_content.strip():
                    os.remove(file_path)
                    return (
                        "No readable text was found in this file. "
                        "Please upload a text-based document.",
                        400
                    )

                connection.execute("""
                    INSERT INTO teaching_materials
                    (
                        teacher_id,
                        class_id,
                        subject_id,
                        title,
                        content,
                        file_name,
                        file_path
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (
                    teacher_id,
                    class_id,
                    subject_id,
                    title,
                    extracted_content,
                    original_filename,
                    file_path
                ))

                connection.commit()

                return redirect(url_for("teacher_materials"))

            except Exception as error:

                connection.rollback()

                if os.path.exists(file_path):
                    os.remove(file_path)

                print("MATERIAL UPLOAD ERROR:", error)
                return "Unable to process the uploaded file.", 400

        materials = connection.execute("""
            SELECT
                teaching_materials.id,
                teaching_materials.title,
                teaching_materials.file_name,
                teaching_materials.created_at,
                classes.name AS class_name,
                subjects.name AS subject_name
            FROM teaching_materials
            JOIN classes
                ON teaching_materials.class_id = classes.id
            JOIN subjects
                ON teaching_materials.subject_id = subjects.id
            WHERE teaching_materials.teacher_id = ?
            ORDER BY teaching_materials.created_at DESC
        """, (teacher_id,)).fetchall()

        return render_template(
            "teacher_materials.html",
            name=session["name"],
            assignments=assignments,
            materials=materials
        )

    finally:
        connection.close()



# -------------------------
# TEACHER ASSESSMENTS
# -------------------------

@app.route("/teacher/assessments", methods=["GET", "POST"])
def teacher_assessments():

    if "user_id" not in session or session.get("role") != "teacher":
        return redirect(url_for("teacher_login"))

    teacher_id = session["user_id"]
    connection = get_connection()

    review_questions = None
    draft_assessment_id = None
    selected_material = None
    error = None
    edit_mode = False

    try:
        if request.method == "POST":

            material_ids = request.form.getlist("material_ids")

            try:
                material_ids = list(set(int(mid) for mid in material_ids))
            except ValueError:
                raise ValueError("Invalid learning material selection.")

            if not material_ids:
                raise ValueError("Please select at least one learning material.")

            placeholders = ",".join("?" for _ in material_ids)

            materials_selected = connection.execute(f"""
                SELECT id, title, content, class_id, subject_id
                FROM teaching_materials
                WHERE teacher_id = ?
                AND id IN ({placeholders})
            """, [teacher_id] + material_ids).fetchall()

            if len(materials_selected) != len(material_ids):
                raise ValueError("One or more selected materials were not found.")

            class_ids = set(m["class_id"] for m in materials_selected)
            subject_ids = set(m["subject_id"] for m in materials_selected)

            if len(class_ids) != 1 or len(subject_ids) != 1:
                raise ValueError(
                    "Select materials belonging to the same class and subject."
                )

            selected_material = ", ".join(
                m["title"] for m in materials_selected
            )

            combined_text = ""

            for index, material in enumerate(materials_selected, start=1):
                combined_text += (
                    f"\n\nMATERIAL {index}: {material['title']}\n"
                    f"{material['content']}"
                )

            generated_questions = generate_questions(combined_text)

            if not generated_questions or not generated_questions.get("mcqs"):
                raise ValueError("The AI did not generate any MCQs.")

            cursor = connection.execute("""
                INSERT INTO assessments
                (teacher_id, class_id, subject_id, title, status)
                VALUES (?, ?, ?, ?, 'draft')
            """, (
                teacher_id,
                materials_selected[0]["class_id"],
                materials_selected[0]["subject_id"],
                "AI Generated Assessment"
            ))

            draft_assessment_id = cursor.lastrowid

            for question in generated_questions["mcqs"]:

                options = question.get("options", [])

                if len(options) != 4:
                    continue

                connection.execute("""
                    INSERT INTO questions (
                        assessment_id,
                        question,
                        option_a,
                        option_b,
                        option_c,
                        option_d,
                        correct_answer,
                        explanation,
                        concept
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    draft_assessment_id,
                    question["question"],
                    options[0],
                    options[1],
                    options[2],
                    options[3],
                    question["answer"],
                    question.get("explanation", ""),
                    question.get("topic", "")
                ))

            connection.commit()

            review_questions = connection.execute("""
                SELECT *
                FROM questions
                WHERE assessment_id = ?
                ORDER BY id
            """, (draft_assessment_id,)).fetchall()

    except Exception as e:
        connection.rollback()
        error = f"Question generation failed: {e}"

    try:
        assignments = get_teacher_assignments(connection, teacher_id)

        materials = connection.execute("""
            SELECT id, title, class_id, subject_id, created_at
            FROM teaching_materials
            WHERE teacher_id = ?
            ORDER BY created_at DESC
        """, (teacher_id,)).fetchall()

        assessments = connection.execute("""
            SELECT
                assessments.id,
                assessments.title,
                assessments.start_time,
                assessments.end_time,
                assessments.status,
                classes.name AS class_name,
                subjects.name AS subject_name,
                COUNT(questions.id) AS question_count
            FROM assessments
            JOIN classes
                ON assessments.class_id = classes.id
            JOIN subjects
                ON assessments.subject_id = subjects.id
            LEFT JOIN questions
                ON assessments.id = questions.assessment_id
            WHERE assessments.teacher_id = ?
            GROUP BY assessments.id
            ORDER BY assessments.id DESC
        """, (teacher_id,)).fetchall()

        return render_template(
            "teacher_assessments.html",
            name=session["name"],
            assignments=assignments,
            materials=materials,
            assessments=assessments,
            review_questions=review_questions,
            draft_assessment_id=draft_assessment_id,
            selected_material=selected_material,
            error=error,
            edit_mode=edit_mode,
            message=session.pop("assessment_message", None)
        )

    finally:
        connection.close()




# -------------------------
# PUBLISH ASSESSMENT
# -------------------------

@app.route("/teacher/assessments/publish/<int:assessment_id>", methods=["POST"])
def publish_teacher_assessment(assessment_id):

    if "user_id" not in session or session.get("role") != "teacher":
        return redirect(url_for("teacher_login"))

    teacher_id = session["user_id"]
    connection = get_connection()

    try:
        assessment = connection.execute("""
            SELECT id
            FROM assessments
            WHERE id = ?
            AND teacher_id = ?
            AND status = 'draft'
        """, (assessment_id, teacher_id)).fetchone()

        if assessment is None:
            return "Assessment not found.", 404

        title = request.form.get("title", "").strip()
        start_time = request.form.get("start_time", "").strip()
        end_time = request.form.get("end_time", "").strip()

        if not title:
            return "Please enter an assessment title.", 400

        try:
            start = datetime.fromisoformat(start_time)
            end = datetime.fromisoformat(end_time)
        except ValueError:
            return "Please enter valid start and end times.", 400

        if start <= datetime.now():
            return "The assessment start time must be in the future.", 400

        if end <= start:
            return "The end time must be after the start time.", 400

        selected_ids = request.form.getlist("question_ids")

        try:
            selected_ids = list(dict.fromkeys(
                int(qid) for qid in selected_ids
            ))
        except ValueError:
            return "Invalid question selection.", 400

        if not 15 <= len(selected_ids) <= 20:
            return "Select between 15 and 20 questions.", 400

        placeholders = ",".join("?" for _ in selected_ids)

        selected_questions = connection.execute(f"""
            SELECT *
            FROM questions
            WHERE assessment_id = ?
            AND id IN ({placeholders})
        """, [assessment_id] + selected_ids).fetchall()

        if len(selected_questions) != len(selected_ids):
            return "Invalid question selection.", 400

        # Validate questions stored in the database
        for question in selected_questions:

            if not all([
                question["question"],
                question["option_a"],
                question["option_b"],
                question["option_c"],
                question["option_d"]
            ]):
                return "A selected question has missing options.", 400

            correct_answer = str(question["correct_answer"] or "").strip()

            if not correct_answer:
                return "A selected question has a missing correct answer.", 400

            valid_letters = ["A", "B", "C", "D"]

            if correct_answer.upper().replace("OPTION ", "") not in valid_letters:

                valid_options = [
                    str(question["option_a"]).strip(),
                    str(question["option_b"]).strip(),
                    str(question["option_c"]).strip(),
                    str(question["option_d"]).strip()
                ]

                if correct_answer.casefold() not in [
                    option.casefold() for option in valid_options
                ]:
                    return "A selected question has an invalid correct answer.", 400

        connection.execute("BEGIN")

        # Remove questions not selected by the teacher
        connection.execute(f"""
            DELETE FROM questions
            WHERE assessment_id = ?
            AND id NOT IN ({placeholders})
        """, [assessment_id] + selected_ids)

        # Publish directly
        connection.execute("""
            UPDATE assessments
            SET
                title = ?,
                start_time = ?,
                end_time = ?,
                status = 'published'
            WHERE id = ?
            AND teacher_id = ?
            AND status = 'draft'
        """, (
            title,
            start_time,
            end_time,
            assessment_id,
            teacher_id
        ))

        connection.commit()

        session["assessment_message"] = "Assessment published successfully."

        return redirect(url_for("teacher_assessments"))

    except Exception as e:
        connection.rollback()
        print("PUBLISH ERROR:", repr(e))
        return "Unable to publish assessment.", 500

    finally:
        connection.close()



# -------------------------
# RUN APPLICATION
# -------------------------

if __name__ == "__main__":
    app.run(
        host="127.0.0.1",
        port=5001,
        debug=False
    )
