import tkinter as tk
from tkinter import scrolledtext, filedialog
import threading
import speech_recognition as sr
import win32com.client
import pythoncom
import time

# ---------------- INITIALIZATION ----------------
pythoncom.CoInitialize()

speaker = win32com.client.Dispatch("SAPI.SpVoice")
recognizer = sr.Recognizer()

stt_running = False
tts_running = False
dark_mode = False

# ---------------- LANGUAGES ----------------
LANGUAGES = {
    "English": "en-IN",
    "Hindi": "hi-IN",
    "Tamil": "ta-IN"
}

# ---------------- VOICES ----------------
voices = speaker.GetVoices()
VOICE_MAP = {v.GetDescription(): v for v in voices}

def set_voice(voice_name):
    speaker.Voice = VOICE_MAP[voice_name]

# ---------------- SPEAK ----------------
def speak(text):
    speaker.Speak(text, 1)
    time.sleep(0.1)

# ---------------- SPEECH TO TEXT ----------------
def speech_to_text(gui_output, status_var, lang_var):
    global stt_running
    stt_running = True
    status_var.set("Status: Listening...")
    speak("Speech to text activated")

    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=1)

        while stt_running:
            try:
                audio = recognizer.listen(source, timeout=5)
                text = recognizer.recognize_google(
                    audio, language=LANGUAGES[lang_var.get()]
                )

                gui_output.insert(tk.END, f"📝 You said: {text}\n")
                gui_output.see(tk.END)

            except sr.WaitTimeoutError:
                continue
            except sr.UnknownValueError:
                speak("Sorry, I could not understand")

    status_var.set("Status: Ready")

# ---------------- TEXT TO SPEECH ----------------
def text_to_speech(gui_output, input_box, status_var):
    global tts_running
    tts_running = True
    status_var.set("Status: Text to Speech Active")
    speak("Text to speech activated")

    def speak_text(event=None):
        global tts_running
        if not tts_running:
            return

        text = input_box.get().strip()
        if text:
            gui_output.insert(tk.END, f"🔊 Speaking: {text}\n")
            gui_output.see(tk.END)
            speak(text)
            input_box.delete(0, tk.END)

    input_box.bind("<Return>", speak_text)

# ---------------- STOP FUNCTION ----------------
def stop_operation(status_var):
    global stt_running, tts_running
    stt_running = False
    tts_running = False
    status_var.set("Status: Operation Stopped")
    speak("Operation stopped")

# ---------------- SAVE CONVERSATION ----------------
def save_conversation(gui_output):
    file = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text Files", "*.txt")]
    )
    if file:
        with open(file, "w", encoding="utf-8") as f:
            f.write(gui_output.get("1.0", tk.END))
        speak("Conversation saved")

# ---------------- DARK MODE ----------------
def toggle_dark_mode(root, output_box, input_box):
    global dark_mode
    dark_mode = not dark_mode

    bg = "#1e1e1e" if dark_mode else "white"
    fg = "white" if dark_mode else "black"

    root.configure(bg=bg)
    output_box.configure(bg=bg, fg=fg, insertbackground=fg)
    input_box.configure(bg=bg, fg=fg, insertbackground=fg)

# ---------------- GUI ----------------
def start_gui():
    root = tk.Tk()
    root.title("EXPRESSA – Assistive Communication")
    root.geometry("760x620")
    root.resizable(False, False)

    status_var = tk.StringVar(value="Status: Ready")
    tk.Label(root, textvariable=status_var,
             font=("Arial", 12, "bold")).pack(pady=5)

    # Language selector
    lang_var = tk.StringVar(value="English")
    tk.Label(root, text="Select Language").pack()
    tk.OptionMenu(root, lang_var, *LANGUAGES.keys()).pack()

    # Voice selector
    voice_var = tk.StringVar(value=list(VOICE_MAP.keys())[0])
    tk.Label(root, text="Select Voice (Boy / Girl)").pack()
    tk.OptionMenu(root, voice_var, *VOICE_MAP.keys(),
                  command=set_voice).pack()
    set_voice(voice_var.get())

    # Output box
    output_box = scrolledtext.ScrolledText(
        root, width=90, height=20, font=("Consolas", 11)
    )
    output_box.pack(pady=10)

    # Input box
    input_box = tk.Entry(root, width=70, font=("Arial", 12))
    input_box.pack(pady=5)

    # Buttons
    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=10)

    tk.Button(btn_frame, text="Speech → Text",
              width=18, bg="#4CAF50", fg="white",
              command=lambda: threading.Thread(
                  target=speech_to_text,
                  args=(output_box, status_var, lang_var),
                  daemon=True).start()
              ).grid(row=0, column=0, padx=5)

    tk.Button(btn_frame, text="Text → Speech",
              width=18, bg="#2196F3", fg="white",
              command=lambda: threading.Thread(
                  target=text_to_speech,
                  args=(output_box, input_box, status_var),
                  daemon=True).start()
              ).grid(row=0, column=1, padx=5)

    tk.Button(btn_frame, text="Save",
              width=18, bg="#9C27B0", fg="white",
              command=lambda: save_conversation(output_box)
              ).grid(row=0, column=2, padx=5)

    tk.Button(btn_frame, text="Dark Mode",
              width=18, bg="#333", fg="white",
              command=lambda: toggle_dark_mode(
                  root, output_box, input_box)
              ).grid(row=0, column=3, padx=5)

    tk.Button(root, text="STOP",
              width=40, bg="red", fg="white",
              font=("Arial", 11, "bold"),
              command=lambda: stop_operation(status_var)
              ).pack(pady=10)

    speak("Welcome to Expressa. Ready for interaction.")
    root.mainloop()

# ---------------- MAIN ----------------
if __name__ == "__main__":
    start_gui()
