import cv2
import pytesseract
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox

# Extract text from video frames using OCR
def extract_contacts_from_video(video_path):
    cap = cv2.VideoCapture(video_path)
    contacts = []
    frame_count = 0
    while True:
        ret, frame = cap.read()
        if not ret:
            break
        frame_count += 1
        if frame_count % 30 == 0:  # process every 30th frame
            text = pytesseract.image_to_string(frame)
            emails = re.findall(r"[\w\.-]+@[\w\.-]+", text)
            phones = re.findall(r"\+?\d[\d\s-]{7,}\d", text)
            names = re.findall(r"Name[:\s]*([A-Za-z ]+)", text)
            companies = re.findall(r"Company[:\s]*([A-Za-z0-9 &]+)", text)
            for i in range(max(len(names), len(companies), len(emails), len(phones))):
                contacts.append({
                    "Name": names[i] if i < len(names) else "",
                    "Company": companies[i] if i < len(companies) else "",
                    "Email": emails[i] if i < len(emails) else "",
                    "Phone": phones[i] if i < len(phones) else ""
                })
    cap.release()
    return contacts

# GUI for selecting video and saving results
def main():
    root = tk.Tk()
    root.withdraw()
    video_path = filedialog.askopenfilename(title="Select Video File", filetypes=[("Video Files", "*.mp4 *.avi *.mov")])
    if not video_path:
        messagebox.showerror("Error", "No video file selected!")
        return
    contacts = extract_contacts_from_video(video_path)
    if not contacts:
        messagebox.showinfo("Result", "No contacts found in the video.")
        return
    df = pd.DataFrame(contacts)
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel File", "*.xlsx"), ("CSV File", "*.csv")])
    if save_path.endswith(".csv"):
        df.to_csv(save_path, index=False)
    else:
        df.to_excel(save_path, index=False)
    messagebox.showinfo("Done", f"Contacts saved to {save_path}")

if __name__ == "__main__":
    main()
