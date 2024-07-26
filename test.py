import cv2
import time
import os
import numpy as np
from datetime import datetime
from sklearn.neighbors import KNeighborsClassifier
import pickle
import csv
from win32com.client import Dispatch
def speak(message):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(message)
facedetect_path = 'data/haarcascade_frontalface_default.xml'
names_path = 'data/names.pkl'
faces_data_path = 'data/faces_data.pkl'
# Check if the required files exist
if not os.path.isfile(facedetect_path):
    print(f"Error: File {facedetect_path} not found.")
    exit()
if not os.path.isfile(names_path) or not os.path.isfile(faces_data_path):
    print(f"Error: File {names_path} or {faces_data_path} not found.")
    exit()
facedetect = cv2.CascadeClassifier(facedetect_path)
with open(names_path, 'rb') as w:
    LABELS = pickle.load(w)
with open(faces_data_path, 'rb') as f:
    FACES = pickle.load(f)
knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)
imgBackground = cv2.imread("background.png")
COL_NAMES = ['NAME', 'TIME', 'STATUS']
def process_frame(frame):
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)
    attendance = []
    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M-%S")
        cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x, y), (x+w, y+h), (50, 50, 255), 2)
        cv2.rectangle(frame, (x, y-40), (x+w, y), (50, 50, 255), -1)
        cv2.putText(frame, str(output[0]), (x, y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
        attendance = [str(output[0]), str(timestamp), "true"]
        break  # Stop after the first detected face
    return frame, attendance
def capture_and_process_frames():
    global attendance_data
    video = cv2.VideoCapture(0, cv2.CAP_DSHOW)  # Use DirectShow backend
    if not video.isOpened():
        print("Failed to open camera")
        return
    detected = False
    start_time = time.time()
    while time.time() - start_time < 5:  # Try for 5 seconds
        ret, frame = video.read()
        if not ret:
            print("Failed to capture frame")
            continue
        processed_frame, attendance = process_frame(frame)
        cv2.imshow("Frame", processed_frame)
        if attendance:
            attendance_data.append(attendance)
            detected = True
            break  # Stop processing after the first valid result
        if cv2.waitKey(1) == ord('q'):
            break
    video.release()
    cv2.destroyAllWindows()
    if not detected:
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M-%S")
        attendance_data.append(["Unknown", timestamp, "false"])
    save_attendance()
def save_attendance():
    global attendance_data
    if len(attendance_data) > 0:
        date = datetime.fromtimestamp(time.time()).strftime("%d-%m-%Y")
        exist = os.path.isfile(f"Attendance/Attendance_{date}.csv")
        if exist:
            with open(f"Attendance/Attendance_{date}.csv", "a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(attendance_data[0])  # Save only the first result
        else:
            with open(f"Attendance/Attendance_{date}.csv", "a") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(COL_NAMES)
                writer.writerow(attendance_data[0])  # Save only the first result
        attendance_data = []  # Clear the cached data after writing
capture_interval = 10  # Capture frame every 120 seconds (2 minutes) for testing
start_time = time.time()
attendance_data = []
# First capture immediately
capture_and_process_frames()
try:
    while True:
        if time.time() - start_time >= capture_interval:
            capture_and_process_frames()
            start_time = time.time()
        if cv2.waitKey(1) == ord('q'):
            break
finally:
    cv2.destroyAllWindows()
