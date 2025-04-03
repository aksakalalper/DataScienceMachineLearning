from PyQt5 import QtWidgets, QtCore
from MainWindow import Ui_MainWindow
import speech_recognition as sr
import threading
from openpyxl import Workbook, load_workbook  # Import openpyxl
import os  # Import os for file existence check

class FaultRecorderApp(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setup_connections()
        self.recognizer = sr.Recognizer()
        self.current_machine_name = ""  # Store the machine name
        self.current_fault_description = ""  # Store the fault description

    def setup_connections(self):
        # Remove the connection to update_selected_machine
        self.ui.comboBox.currentIndexChanged.connect(self.dummy_method)  # Replace with a dummy method
        self.ui.speakButton.clicked.connect(self.start_speech_recognition)
        self.ui.buttonBox.accepted.connect(self.handle_yes_response)
        self.ui.buttonBox.rejected.connect(self.handle_no_response)
        self.ui.saveButton.clicked.connect(self.save_manual_entry)
        self.ui.exitButton.clicked.connect(self.close_application)
        self.ui.manualEntryButton.clicked.connect(self.show_manual_entry_ui)  # Connect manual entry button

    def dummy_method(self):
        # Do nothing when an item is selected
        pass

    def start_speech_recognition(self):
        # Check if a microphone is available
        self.ui.resultOutput.setText("Mikrofon kontrolleri yapılıyor...")
        if not self.is_microphone_available():
            self.ui.resultOutput.setText("Mikrofon bulunamadı. Lütfen bir mikrofon bağlayın.")
            return

        # Hide UI elements and start speech recognition for machine name
        
        self.ui.manualEntryText.hide()
        self.ui.machineName.hide()
        self.ui.faultDefinition.hide()
        self.ui.comboBox.hide()
        self.ui.faultDefinitionManualEntry.hide()
        self.ui.saveButton.hide()
        threading.Thread(target=self.capture_machine_name, daemon=True).start()

    def is_microphone_available(self):
        """Check if a microphone is available."""
        try:
            mic_list = sr.Microphone.list_microphone_names()
            return len(mic_list) > 0
        except OSError:
            return False

    def capture_machine_name(self):
        try:
            with sr.Microphone() as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=2)
                QtCore.QMetaObject.invokeMethod(self.ui.resultOutput, "setText", QtCore.Qt.QueuedConnection, QtCore.Q_ARG(str, "Lütfen makine adını söyleyin..."))
                audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=10)
                machine_name = self.recognizer.recognize_google(audio, language="tr-TR")
                self.current_machine_name = machine_name
                QtCore.QMetaObject.invokeMethod(self.ui.speechOutput, "setText", QtCore.Qt.QueuedConnection, QtCore.Q_ARG(str, f"Makine adı: {machine_name}"))
                QtCore.QMetaObject.invokeMethod(self.ui.checkSpeechOutput, "setText", QtCore.Qt.QueuedConnection, QtCore.Q_ARG(str, "Makine adı doğru mu?"))
                QtCore.QMetaObject.invokeMethod(self.ui.checkSpeechOutput, "show", QtCore.Qt.QueuedConnection)
                QtCore.QMetaObject.invokeMethod(self.ui.buttonBox, "show", QtCore.Qt.QueuedConnection)
        except sr.UnknownValueError:
            QtCore.QMetaObject.invokeMethod(self.ui.resultOutput, "setText", QtCore.Qt.QueuedConnection, QtCore.Q_ARG(str, "Makine adı anlaşılamadı."))
        except sr.RequestError:
            QtCore.QMetaObject.invokeMethod(self.ui.resultOutput, "setText", QtCore.Qt.QueuedConnection, QtCore.Q_ARG(str, "Bağlantı veya API hatası."))
        except sr.WaitTimeoutError:
            QtCore.QMetaObject.invokeMethod(self.ui.resultOutput, "setText", QtCore.Qt.QueuedConnection, QtCore.Q_ARG(str, "Belirtilen süre boyunca konuşma olmadı."))

    def handle_yes_response(self):
        if not self.current_fault_description:
            # Proceed to capture fault description
            self.ui.resultOutput.setText("Mikrofon kontrolleri yapılıyor...")
            threading.Thread(target=self.capture_fault_description, daemon=True).start()
        else:
            # Save the data if fault description is already captured
            self.save_to_excel(self.current_machine_name, self.current_fault_description)
            self.ui.resultOutput.setText("Kayıt başarıyla eklendi!")
            self.reset_ui()

    def handle_no_response(self):
        """Handle the 'No' response from the user."""
        if self.current_fault_description:
            # If fault description is being confirmed, restart fault description capture
            self.ui.resultOutput.setText("Lütfen arıza tanımını tekrar söyleyin.")
            self.current_fault_description = ""  # Clear the current fault description
            self.ui.speechOutput.clear()  # Clear the speech output
            self.ui.checkSpeechOutput.hide()  # Hide the confirmation label
            self.ui.buttonBox.hide()  # Hide the Yes/No buttons
            threading.Thread(target=self.capture_fault_description, daemon=True).start()
        else:
            # If machine name is being confirmed, restart machine name capture
            self.ui.resultOutput.setText("Lütfen makine adını tekrar söyleyin.")
            self.current_machine_name = ""  # Clear the current machine name
            self.ui.speechOutput.clear()  # Clear the speech output
            self.ui.checkSpeechOutput.hide()  # Hide the confirmation label
            self.ui.buttonBox.hide()  # Hide the Yes/No buttons
            threading.Thread(target=self.capture_machine_name, daemon=True).start()

    def capture_fault_description(self):
        try:
            with sr.Microphone() as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=2)
                QtCore.QMetaObject.invokeMethod(self.ui.resultOutput, "setText", QtCore.Qt.QueuedConnection, QtCore.Q_ARG(str, "Lütfen arıza tanımını söyleyin..."))
                audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=30)
                fault_description = self.recognizer.recognize_google(audio, language="tr-TR")
                self.current_fault_description = fault_description
                QtCore.QMetaObject.invokeMethod(self.ui.speechOutput, "setText", QtCore.Qt.QueuedConnection, QtCore.Q_ARG(str, f"Arıza tanımı: {fault_description}"))
                QtCore.QMetaObject.invokeMethod(self.ui.checkSpeechOutput, "setText", QtCore.Qt.QueuedConnection, QtCore.Q_ARG(str, "Arıza tanımı doğru mu?"))
                QtCore.QMetaObject.invokeMethod(self.ui.checkSpeechOutput, "show", QtCore.Qt.QueuedConnection)
                QtCore.QMetaObject.invokeMethod(self.ui.buttonBox, "show", QtCore.Qt.QueuedConnection)
        except sr.UnknownValueError:
            QtCore.QMetaObject.invokeMethod(self.ui.resultOutput, "setText", QtCore.Qt.QueuedConnection, QtCore.Q_ARG(str, "Arıza tanımı anlaşılamadı."))
        except sr.RequestError:
            QtCore.QMetaObject.invokeMethod(self.ui.resultOutput, "setText", QtCore.Qt.QueuedConnection, QtCore.Q_ARG(str, "Bağlantı veya API hatası."))
        except sr.WaitTimeoutError:
            QtCore.QMetaObject.invokeMethod(self.ui.resultOutput, "setText", QtCore.Qt.QueuedConnection, QtCore.Q_ARG(str, "Belirtilen süre boyunca konuşma olmadı."))

    def approve_transcription(self):
        text = self.ui.speechOutput.toPlainText().replace("Konuşulan metin çıktısı: ", "").strip()
        if text:
            try:
                self.save_to_excel("Speech Transcription", text)
                self.ui.resultOutput.setText("Kayıt başarıyla eklendi!")
            except Exception as e:
                self.ui.resultOutput.setText(f"Excel dosyasına yazma hatası: {e}")
        else:
            self.ui.resultOutput.setText("Kayıt için geçerli bir metin bulunamadı.")
        self.reset_ui()

    def reject_transcription(self):
        self.ui.resultOutput.setText("Lütfen bilgileri elle giriniz...")
        self.reset_ui()
        self.show_manual_entry_ui()  # Show the manual entry UI
        self.ui.resultOutput.setText("Kayıt eklenemedi.")

    def save_manual_entry(self):
        machine_name = self.ui.comboBox.currentText()
        fault_description = self.ui.faultDefinitionManualEntry.toPlainText().strip()  # Use toPlainText() for QTextEdit
        if machine_name and fault_description:
            try:
                self.save_to_excel(machine_name, fault_description)
                self.ui.resultOutput.setText("Manuel kayıt eklendi.")
                self.ui.faultDefinitionManualEntry.clear()  # Clear the input field after saving
            except Exception as e:
                self.ui.resultOutput.setText(f"Excel dosyasına yazma hatası: {e}")
        else:
            self.ui.resultOutput.setText("Lütfen tüm alanları doldurun.")

    def save_to_excel(self, machine_name, fault_description):
        # Update the file path to use the C: drive
        directory = "C:/BakimKayitlari"
        file_path = f"{directory}/ariza_kayitlari.xlsx"

        # Ensure the directory exists
        try:
            if not os.path.exists(directory):
                os.makedirs(directory, exist_ok=True)  # Create the directory if it doesn't exist
        except OSError as e:
            self.ui.resultOutput.setText(f"Dosya dizini oluşturulamadı: {e}")
            return

        # Check if the file exists and create or append accordingly
        try:
            if not os.path.exists(file_path):
                # Create a new workbook if the file doesn't exist
                workbook = Workbook()
                sheet = workbook.active
                sheet.title = "Fault Records"
                sheet.append(["Timestamp", "Machine Name", "Fault Description"])  # Add headers
            else:
                # Load the existing workbook
                workbook = load_workbook(file_path)
                sheet = workbook.active

            # Append the new record
            timestamp = QtCore.QDateTime.currentDateTime().toString("yyyy-MM-dd HH:mm:ss")
            sheet.append([timestamp, machine_name, fault_description])

            # Save the workbook
            workbook.save(file_path)
            self.ui.resultOutput.setText("Kayıt başarıyla kaydedildi.")
        except PermissionError:
            self.ui.resultOutput.setText("Dosya yazma izni yok. Lütfen dosyanın açık olmadığından emin olun.")
        except Exception as e:
            self.ui.resultOutput.setText(f"Excel dosyasına yazma hatası: {e}")

    def reset_ui(self):
        self.ui.speechOutput.clear()
        self.ui.checkSpeechOutput.setText("")
        self.ui.checkSpeechOutput.hide()  # Hide the label
        self.ui.buttonBox.hide()  # Hide the buttons
        self.ui.resultOutput.setText("Kayıt başarıyla eklendi!")
        self.ui.manualEntryText.hide()  # Hide manual entry text
        self.ui.machineName.hide()  # Hide machine name label
        self.ui.faultDefinition.hide()  # Hide fault definition label
        self.ui.comboBox.hide()  # Hide combo box
        self.ui.faultDefinitionManualEntry.hide()  # Hide fault definition manual entry
        self.ui.saveButton.hide()  # Hide save button

    def show_manual_entry_ui(self):
        """Toggle the visibility of the manual entry UI elements."""
        if self.ui.manualEntryText.isVisible():
            # Hide the manual entry UI elements
            self.ui.resultOutput.setText("Manuel giriş ekranı kapatıldı.")
            self.ui.manualEntryText.hide()
            self.ui.machineName.hide()
            self.ui.faultDefinition.hide()
            self.ui.comboBox.hide()
            self.ui.faultDefinitionManualEntry.hide()
            self.ui.saveButton.hide()
        else:
            # Show the manual entry UI elements
            self.ui.resultOutput.setText("Manuel giriş ekranı açıldı.")
            self.ui.manualEntryText.show()
            self.ui.machineName.show()
            self.ui.faultDefinition.show()
            self.ui.comboBox.show()
            self.ui.faultDefinitionManualEntry.show()
            self.ui.saveButton.show()

    def close_application(self):
        self.close()

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    window = FaultRecorderApp()
    window.show()
    sys.exit(app.exec_())
