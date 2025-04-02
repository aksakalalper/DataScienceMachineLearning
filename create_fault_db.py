import speech_recognition as sr
import time
import tkinter as tk
import threading
import os
from tkinter import messagebox

r = sr.Recognizer()

r.pause_threshold = 2.0  # 2 seconds of silence before stopping

# Saves text to file with timestamp
def save_text(text):
    dosya_yolu = "D:/BakimKayitlari/ariza_kayitlari.txt"
    os.makedirs(os.path.dirname(dosya_yolu), exist_ok=True)
    with open(dosya_yolu, "a", encoding="utf-8") as file:
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        file.write(f"{timestamp} - {text}\n")
    print(f"'{dosya_yolu}' dosyasına kayıt eklendi.")
    status_label.config(text=f"Kayıt eklendi.")

# Converts microphone speech to text
def convert_speech_to_text(timeout_value=5, phrase_limit=60):
    with sr.Microphone() as source:
        status_label.config(text="Mikrofon kontrolü yapılıyor...")
        window.update()
        r.adjust_for_ambient_noise(source, duration=3)
        status_label.config(text="Lütfen konuşun...")
        window.update()
        try:
            audio = r.listen(source, timeout=timeout_value, phrase_time_limit=phrase_limit)
            status_label.config(text="Konuşma algılandı, işleniyor...")
            window.update()
            text = r.recognize_google(audio, language="tr-TR")
            return text
        except sr.UnknownValueError:
            status_label.config(text="Konuşma anlaşılamadı.")
            return None
        except sr.RequestError:
            status_label.config(text="Bağlantı veya API hatası.")
            return None
        except sr.WaitTimeoutError:
            status_label.config(text="Belirtilen süre boyunca konuşma olmadı.")
            return None

# Starts recording in a thread
def start_recording():
    speak_button.config(state="disabled")
    result_entry.delete("1.0", tk.END)  # Clear the text field
    window.update()

    # Processes speech and updates UI
    def process_speech():
        text = convert_speech_to_text()
        if text:
            result_entry.config(state="normal")  # Enable editing
            result_entry.delete("1.0", tk.END)  # Clear previous text
            result_entry.insert("1.0", f"Konuşulan metin çıktısı: {text}")  # Insert new text
            # Unhide "Metin doğru mu?", "Evet" and "Hayır" labels
            question_label.grid(row=0, column=0, columnspan=2, pady=5)
            yes_button.grid(row=1, column=0, padx=15, pady=5)
            no_button.grid(row=1, column=1, padx=15, pady=5)
            yes_button.config(state="normal", command=lambda: approve_text())
            no_button.config(state="normal")
        else:
            status_label.config(text="Kayıt başarısız. Tekrar deneyin.")
            reset_interface()
        window.update()

    threading.Thread(target=process_speech, daemon=True).start()

# Approves and saves text
def approve_text():
    # Get the edited text from the Text widget, remove the prefix
    edited_text = result_entry.get("1.0", tk.END).strip().replace("Konuşulan metin çıktısı: ", "")
    if edited_text.strip():  # Ensure the text is not empty
        save_text(edited_text)
    reset_interface()

# Rejects text and restarts recording
def reject_text():
    status_label.config(text="Tekrar konuşmanız gerekiyor...")
    reset_interface()
    start_recording()

# Resets UI to initial state
def reset_interface():
    speak_button.config(state="normal")
    yes_button.config(state="disabled")
    no_button.config(state="disabled")
    # Hide "Metin doğru mu?", "Evet" and "Hayır" labels
    question_label.grid_forget()
    yes_button.grid_forget()
    no_button.grid_forget()
    result_entry.delete("1.0", tk.END)  # Clear the text field
    result_entry.config(state="disabled")  # Disable editing

# Create main window
window = tk.Tk()
window.title("Arıza Kayıt Programı")
window.geometry("450x450")
window.configure(bg="#f0f0f0")  

# Set window icon (optional)
welcome_label = tk.Label(window, text="Arıza Kayıt Programı\n\nHoşgeldiniz!", 
                         font=("Helvetica", 16, "bold"), bg="#f0f0f0", fg="#333333")
welcome_label.pack(pady=15)

status_label = tk.Label(window, text="Durum: Hazır", font=("Helvetica", 10), 
                        bg="#f0f0f0", fg="#555555")
status_label.pack(pady=5)

# Replace Entry with Text widget for editable text with increased height, initially disabled
result_entry = tk.Text(window, font=("Helvetica", 12), bg="#ffffff", 
                       fg="#000000", relief="sunken", borderwidth=2, 
                       state="disabled", height=7, wrap="word")
result_entry.pack(pady=10, fill="x", padx=20)

speak_button = tk.Button(window, text="Konuş", command=start_recording, 
                         font=("Helvetica", 12, "bold"), bg="#2196F3", fg="white", 
                         activebackground="#1e88e5", width=15)
speak_button.pack(pady=15)

# Create a frame for buttons
button_frame = tk.Frame(window, bg="#f0f0f0")
button_frame.pack(pady=30)

# Create labels and buttons for "Metin doğru mu?", "Evet" and "Hayır"
# Initially hidden
question_label = tk.Label(button_frame, text="Konuşulan metin çıktısı doğru mu?", 
                          font=("Helvetica", 12), bg="#f0f0f0", fg="#333333")

yes_button = tk.Button(button_frame, text="Evet", font=("Helvetica", 12), 
                       bg="#4CAF50", fg="white", activebackground="#45a049", 
                       width=10, state="disabled")

no_button = tk.Button(button_frame, text="Hayır", command=reject_text, 
                      font=("Helvetica", 12), bg="#f44336", fg="white", 
                      activebackground="#e53935", width=10, state="disabled")

# Place exit button in the top-right corner
exit_button = tk.Button(window, text="Çıkış", command=window.quit, 
                        font=("Helvetica", 12), bg="#757575", fg="white", 
                        activebackground="#616161", width=7)
exit_button.place(relx=1.0, rely=0.0, anchor="ne", x=-10, y=10)

window.mainloop()