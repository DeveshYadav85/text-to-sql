# builds data/eval_set.json. add more (q, gold, level) rows here to reach 80-100 questions.
import json

J_SD = "JOIN departments d ON s.dept_id=d.dept_id"
J_SM = "JOIN marks m ON m.student_id=s.student_id"
J_MC = "JOIN courses c ON m.course_id=c.course_id"

rows = [
 # easy
 ("How many students are there?", "SELECT COUNT(*) FROM students;", "easy"),
 ("List all department names.", "SELECT dept_name FROM departments;", "easy"),
 ("How many courses are there?", "SELECT COUNT(*) FROM courses;", "easy"),
 ("Show all courses with more than 3 credits.", "SELECT course_name FROM courses WHERE credits > 3;", "easy"),
 ("How many female students are there?", "SELECT COUNT(*) FROM students WHERE gender='F';", "easy"),
 ("How many first year students are there?", "SELECT COUNT(*) FROM students WHERE year=1;", "easy"),
 ("What is the highest score in the marks table?", "SELECT MAX(score) FROM marks;", "easy"),
 ("What is the average score across all marks?", "SELECT AVG(score) FROM marks;", "easy"),
 ("How many marks records are in semester 3?", "SELECT COUNT(*) FROM marks WHERE semester=3;", "easy"),
 ("List the names of fourth year students.", "SELECT name FROM students WHERE year=4;", "easy"),
 ("How many marks records have a score above 90?", "SELECT COUNT(*) FROM marks WHERE score > 90;", "easy"),
 ("What is the dept code of Mechanical Engineering?", "SELECT dept_code FROM departments WHERE dept_name='Mechanical Engineering';", "easy"),
 ("What is the lowest score in the marks table?", "SELECT MIN(score) FROM marks;", "easy"),
 ("How many marks records are fails?", "SELECT COUNT(*) FROM marks WHERE score < 40;", "easy"),
 ("List the course names that have exactly 4 credits.", "SELECT course_name FROM courses WHERE credits = 4;", "easy"),
 # medium
 ("How many students are in CSE?", f"SELECT COUNT(*) FROM students s {J_SD} WHERE d.dept_code='CSE';", "medium"),
 ("List the names of students in Mechanical Engineering.", f"SELECT s.name FROM students s {J_SD} WHERE d.dept_name='Mechanical Engineering';", "medium"),
 ("List the courses offered by the CSE department.", "SELECT c.course_name FROM courses c JOIN departments d ON c.dept_id=d.dept_id WHERE d.dept_code='CSE';", "medium"),
 ("What is the average score in DBMS?", f"SELECT AVG(m.score) FROM marks m {J_MC} WHERE c.course_name='DBMS';", "medium"),
 ("How many students passed in AI?", f"SELECT COUNT(*) FROM marks m {J_MC} WHERE c.course_name='AI' AND m.score >= 40;", "medium"),
 ("Show the names of students who scored above 90 in any course.", f"SELECT DISTINCT s.name FROM students s {J_SM} WHERE m.score > 90;", "medium"),
 ("How many female students are in ECE?", f"SELECT COUNT(*) FROM students s {J_SD} WHERE d.dept_code='ECE' AND s.gender='F';", "medium"),
 ("Show the number of students in each department.", f"SELECT d.dept_name, COUNT(*) FROM students s {J_SD} GROUP BY d.dept_id;", "medium"),
 ("Show the average score for each course with the course name.", f"SELECT c.course_name, AVG(m.score) FROM marks m {J_MC} GROUP BY c.course_id;", "medium"),
 ("Show the names of CSE students scoring above 80 in any course.", f"SELECT DISTINCT s.name FROM students s {J_SD} {J_SM} WHERE d.dept_code='CSE' AND m.score > 80;", "medium"),
 ("How many marks records are there for each semester?", "SELECT semester, COUNT(*) FROM marks GROUP BY semester;", "medium"),
 ("Show the total credits of courses in each department.", "SELECT d.dept_name, SUM(c.credits) FROM courses c JOIN departments d ON c.dept_id=d.dept_id GROUP BY d.dept_id;", "medium"),
 ("How many second year CSE students are there?", f"SELECT COUNT(*) FROM students s {J_SD} WHERE d.dept_code='CSE' AND s.year=2;", "medium"),
 ("List the names of students who failed in Operating Systems.", f"SELECT DISTINCT s.name FROM students s {J_SM} {J_MC} WHERE c.course_name='Operating Systems' AND m.score < 40;", "medium"),
 ("What is the maximum score in Thermodynamics?", f"SELECT MAX(m.score) FROM marks m {J_MC} WHERE c.course_name='Thermodynamics';", "medium"),
 # hard
 ("Show the top 5 CSE students by average score.", f"SELECT s.name, AVG(m.score) a FROM students s {J_SD} {J_SM} WHERE d.dept_code='CSE' GROUP BY s.student_id ORDER BY a DESC LIMIT 5;", "hard"),
 ("Which departments have more than 25 students?", f"SELECT d.dept_name FROM students s {J_SD} GROUP BY d.dept_id HAVING COUNT(*) > 25;", "hard"),
 ("List the names of students whose average score is above 70.", f"SELECT s.name FROM students s {J_SM} GROUP BY s.student_id HAVING AVG(m.score) > 70;", "hard"),
 ("Which course has the highest average score?", f"SELECT c.course_name FROM marks m {J_MC} GROUP BY c.course_id ORDER BY AVG(m.score) DESC LIMIT 1;", "hard"),
 ("How many students failed at least one course?", "SELECT COUNT(DISTINCT student_id) FROM marks WHERE score < 40;", "hard"),
 ("Show the average score of CSE students for each semester.", f"SELECT m.semester, AVG(m.score) FROM students s {J_SD} {J_SM} WHERE d.dept_code='CSE' GROUP BY m.semester;", "hard"),
 ("Show the names of students who scored above the overall average in any course.", f"SELECT DISTINCT s.name FROM students s {J_SM} WHERE m.score > (SELECT AVG(score) FROM marks);", "hard"),
 ("Which department has the highest average score?", f"SELECT d.dept_name FROM students s {J_SD} {J_SM} GROUP BY d.dept_id ORDER BY AVG(m.score) DESC LIMIT 1;", "hard"),
 ("Compare the average score of male and female students.", f"SELECT s.gender, AVG(m.score) FROM students s {J_SM} GROUP BY s.gender;", "hard"),
 ("Which departments have an average score above 60?", f"SELECT d.dept_name FROM students s {J_SD} {J_SM} GROUP BY d.dept_id HAVING AVG(m.score) > 60;", "hard"),
]
json.dump([{"q": q, "gold": g, "level": l} for q, g, l in rows], open("data/eval_set.json", "w"), indent=1)
print(len(rows), "questions written")
