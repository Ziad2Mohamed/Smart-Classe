import cv2
import dlib
import face_recognition
import numpy as np
import openpyxl
import glob
import os
import random
from datetime import datetime
from ultralytics import YOLO
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# 🧠 إعدادات كشف الوجه
students_images_path = r"C:\Users\DELL\PycharmProjects\capestone\.venv\image\Ziad"
known_face_encodings = []
known_face_names = []

for file in glob.glob(f"{students_images_path}/*.jpg"):
    image = face_recognition.load_image_file(file)
    encoding = face_recognition.face_encodings(image)[0]
    known_face_encodings.append(encoding)
    name = os.path.splitext(os.path.basename(file))[0]
    known_face_names.append(name)

# 🟢 الوضع الحالي: Lecture أو Exam
current_mode = "Lecture"
selected_student = None
student_scores = {name: 0 for name in known_face_names}  # تخزين درجات الطلاب

# 🧾 إنشاء أو فتح ملف الإكسل
attendance_list = set()

def send_email(subject, body):
    sender_email = "yoseifzeyad2@gmail.com"
    receiver_email = "yoseifzeyad2@gmail.com"
    app_password = "1346795953"

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, app_password)
            server.send_message(msg)
        print("📧 Email sent successfully!")
    except Exception as e:
        print("❌ Email failed:", e)

def create_or_open_workbook():
    current_date = datetime.now().strftime("%Y-%m-%d")
    file_name = f"{current_date}.xlsx"

    if not os.path.exists(file_name):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Event", "Student", "Time", "Details"])
        workbook.save(file_name)
    else:
        workbook = openpyxl.load_workbook(file_name)
        sheet = workbook.active

    return workbook, sheet

# 🤖 تحميل نموذج YOLO
model = YOLO(r"C:\Users\DELL\OneDrive\Desktop\yolov8s.pt")

# 📷 تشغيل الكاميرا
cap = cv2.VideoCapture(0)

workbook, sheet = create_or_open_workbook()

while True:
    ret, frame = cap.read()
    if not ret:
        break

    rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

    if current_mode == "Lecture":
        face_locations = face_recognition.face_locations(rgb_frame)
        face_encodings = face_recognition.face_encodings(rgb_frame, face_locations)

        for face_encoding, face_location in zip(face_encodings, face_locations):
            matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
            name = "Unknown"
            face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)

            if True in matches:
                best_match_index = np.argmin(face_distances)
                name = known_face_names[best_match_index]

            if name != "Unknown" and name not in attendance_list:
                attendance_list.add(name)
                current_time = datetime.now().strftime("%H:%M:%S")
                sheet.append(["Attendance", name, current_time, "Present"])
                workbook.save(f"{datetime.now().strftime('%Y-%m-%d')}.xlsx")

            top, right, bottom, left = face_location
            cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2)
            cv2.putText(frame, name, (left, top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 2)

        if cv2.waitKey(1) & 0xFF == ord('s'):
            if attendance_list:
                selected_student = random.choice(list(attendance_list))
                print(f"👨‍🎓 Selected student: {selected_student}")
                current_time = datetime.now().strftime("%H:%M:%S")
                sheet.append(["Question", selected_student, current_time, "Asked a question"])
                workbook.save(f"{datetime.now().strftime('%Y-%m-%d')}.xlsx")

        if selected_student:
            cv2.putText(frame, f"QUESTION: {selected_student}", (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 1.2, (255, 255, 0), 3)

        # إذا الدكتور ضغط a الطالب جاوب - ياخد درجة
        if cv2.waitKey(1) & 0xFF == ord('a') and selected_student:
            student_scores[selected_student] += 1
            current_time = datetime.now().strftime("%H:%M:%S")
            sheet.append(["Answer", selected_student, current_time, "Correct answer +1"])
            workbook.save(f"{datetime.now().strftime('%Y-%m-%d')}.xlsx")
            selected_student = None

        # إذا الدكتور ضغط w الطالب معرفش - يتم خصم درجة
        if cv2.waitKey(1) & 0xFF == ord('w') and selected_student:
            student_scores[selected_student] -= 1
            current_time = datetime.now().strftime("%H:%M:%S")
            sheet.append(["Answer", selected_student, current_time, "Wrong answer -1"])
            workbook.save(f"{datetime.now().strftime('%Y-%m-%d')}.xlsx")
            selected_student = None

    elif current_mode == "Exam":
        results = model(frame)
        for result in results:
            for detection in result.boxes.data:
                x, y, x2, y2, confidence, class_id = detection.tolist()
                if confidence > 0.6 and result.names[int(class_id)] in ["knife", "gun"]:
                    cv2.rectangle(frame, (int(x), int(y)), (int(x2), int(y2)), (0, 0, 255), 2)
                    cv2.putText(frame, f"ALERT: {result.names[int(class_id)]}", (int(x), int(y) - 10),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2)

                    current_time = datetime.now().strftime("%H:%M:%S")
                    sheet.append(["Dangerous Item", result.names[int(class_id)], current_time, "Detected"])
                    workbook.save(f"{datetime.now().strftime('%Y-%m-%d')}.xlsx")

                    send_email(f"Alert: {result.names[int(class_id)]} detected", f"Time: {current_time} - {result.names[int(class_id)]} detected!")

    cv2.putText(frame, f"MODE: {current_mode}", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (100, 255, 255), 2)
    cv2.imshow("Classroom Monitoring", frame)

    key = cv2.waitKey(1) & 0xFF
    if key == 27:
        break
    elif key == ord('m'):
        current_mode = "Exam" if current_mode == "Lecture" else "Lecture"

cap.release()
cv2.destroyAllWindows()
workbook.save(f"{datetime.now().strftime('%Y-%m-%d')}.xlsx")
