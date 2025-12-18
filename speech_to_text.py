import speech_recognition as sr
import keyboard
import time

recognizer = sr.Recognizer()

def listen_once():
    print("Hold SPACE to speak...")

    # Wait for SPACE press
    keyboard.wait("space")
    print("🎤 Recording started...")

    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source)

        start_time = time.time()
        audio_data = None

        # Loop until SPACE released or 15 seconds passed
        while True:
            # If SPACE released → stop immediately
            if not keyboard.is_pressed("space"):
                print("🛑 Recording stopped (SPACE released)")
                break

            # If max recording time exceeded → stop automatically
            if time.time() - start_time >= 15:
                print("🛑 Recording stopped (15 sec limit reached)")
                break

            # Listen in non-blocking chunks
            try:
                audio_data = recognizer.listen(source, timeout=0.1, phrase_time_limit=None)
            except sr.WaitTimeoutError:
                continue  # no audio in this small slice → keep looping

    if audio_data is None:
        print("⚠ No audio captured!")
        return ""

    # Convert audio to text
    try:
        text = recognizer.recognize_google(audio_data)
        print("You said:", text)
        return text.lower()
    except sr.UnknownValueError:
        print("⚠ Could not understand speech")
        return ""
    except sr.RequestError as e:
        print("⚠ Google STT error:", e)
        return ""
