import os
import sys
import cv2
import moviepy.editor as mpy
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QPushButton,
    QLabel, QFileDialog, QProgressBar, QComboBox, QWidget, 
    QMessageBox, QHBoxLayout, QFrame
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize, QPropertyAnimation, QEasingCurve
from PyQt5.QtGui import QIcon, QFont, QPixmap, QPalette, QColor, QLinearGradient, QBrush

# Suppress warnings
import warnings
warnings.filterwarnings("ignore")

class AnimatedProgressBar(QProgressBar):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._animation = QPropertyAnimation(self, b"value")
        self._animation.setEasingCurve(QEasingCurve.OutQuad)
        self._animation.setDuration(800)
        
    def setValueAnimated(self, value):
        self._animation.stop()
        self._animation.setStartValue(self.value())
        self._animation.setEndValue(value)
        self._animation.start()

class ConverterThread(QThread):
    progress_updated = pyqtSignal(int)
    conversion_finished = pyqtSignal(bool, str)

    def __init__(self, input_file, output_file):
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file

    def run(self):
        try:
            # Step 1: Read video with OpenCV (works with WMV files)
            cap = cv2.VideoCapture(self.input_file)
            if not cap.isOpened():
                raise ValueError("Could not open video file. Make sure Windows Media Player codecs are installed.")

            fps = cap.get(cv2.CAP_PROP_FPS)
            width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
            height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
            total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))

            # Step 2: Extract audio with MoviePy
            try:
                audio_clip = mpy.AudioFileClip(self.input_file)
            except:
                audio_clip = None  # Some WMV files might not have extractable audio
            
            # Create temporary video file path
            temp_output = os.path.splitext(self.output_file)[0] + "_temp.mp4"
            
            # Step 3: Write video frames
            fourcc = cv2.VideoWriter_fourcc(*'mp4v')
            out = cv2.VideoWriter(temp_output, fourcc, fps, (width, height))

            frame_count = 0
            while cap.isOpened():
                ret, frame = cap.read()
                if not ret:
                    break

                # Convert BGR to RGB
                frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                out.write(frame)
                frame_count += 1
                progress = int((frame_count / total_frames) * 100)
                self.progress_updated.emit(min(progress, 95))

            cap.release()
            out.release()

            # Step 4: Combine video and audio if audio exists
            video_clip = mpy.VideoFileClip(temp_output)
            if audio_clip:
                final_clip = video_clip.set_audio(audio_clip)
            else:
                final_clip = video_clip
                
            final_clip.write_videofile(
                self.output_file,
                codec='libx264',
                audio_codec='aac' if audio_clip else None,
                fps=fps,
                threads=4,
                logger=None,
                preset='slow',
                bitrate='5000k'
            )
            
            # Cleanup
            video_clip.close()
            if audio_clip:
                audio_clip.close()
            os.remove(temp_output)
            
            self.progress_updated.emit(100)
            self.conversion_finished.emit(True, self.output_file)

        except Exception as e:
            self.conversion_finished.emit(False, str(e))

class MediaConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.setup_styles()
        self.input_file = ""
        self.converter_thread = None

    def setup_ui(self):
        self.setWindowTitle("Python Media Converter")
        self.setMinimumSize(750, 550)
        self.setAcceptDrops(True)

        # Main container with vibrant background
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(25)

        # Header with icon
        header = QLabel()
        header_layout = QHBoxLayout()
        
        icon_label = QLabel()
        icon_label.setPixmap(QIcon.fromTheme("video-x-generic").pixmap(64, 64))
        
        title = QLabel("PYTHON MEDIA CONVERTER")
        title.setStyleSheet("""
            font-size: 32px;
            font-weight: bold;
            color: #FFFFFF;
            letter-spacing: 1px;
        """)
        
        header_layout.addWidget(icon_label)
        header_layout.addWidget(title)
        header_layout.addStretch()
        header.setLayout(header_layout)

        # File selection card
        file_card = QFrame()
        file_card.setFrameShape(QFrame.StyledPanel)
        file_card.setStyleSheet("""
            QFrame {
                background-color: #FFFFFF;
                border-radius: 15px;
                padding: 25px;
                border: 2px solid #E0E0E0;
            }
        """)
        
        file_layout = QVBoxLayout()
        file_layout.setSpacing(20)

        self.file_label = QLabel("Drag & drop a media file or click below")
        self.file_label.setAlignment(Qt.AlignCenter)
        self.file_label.setStyleSheet("""
            font-size: 18px;
            color: #333333;
            margin-bottom: 15px;
        """)

        self.browse_btn = QPushButton("BROWSE FILES")
        self.browse_btn.setIcon(QIcon.fromTheme("document-open"))
        self.browse_btn.setStyleSheet(self.get_button_style())
        self.browse_btn.setMinimumHeight(60)
        self.browse_btn.clicked.connect(self.select_file)

        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.browse_btn)
        file_card.setLayout(file_layout)

        # Settings card
        settings_card = QFrame()
        settings_card.setFrameShape(QFrame.StyledPanel)
        settings_card.setStyleSheet("""
            QFrame {
                background-color: #FFFFFF;
                border-radius: 15px;
                padding: 25px;
                border: 2px solid #E0E0E0;
            }
        """)
        
        settings_layout = QVBoxLayout()
        settings_layout.setSpacing(20)

        format_label = QLabel("OUTPUT FORMAT:")
        format_label.setStyleSheet("""
            font-size: 18px;
            color: #333333;
            font-weight: bold;
        """)

        self.format_combo = QComboBox()
        self.format_combo.addItems(["MP4 (Recommended)", "MOV", "AVI", "MKV"])
        self.format_combo.setStyleSheet("""
            QComboBox {
                background-color: #F5F5F5;
                color: #333333;
                border: 2px solid #CCCCCC;
                border-radius: 8px;
                padding: 12px;
                min-height: 50px;
                font-size: 16px;
                font-weight: bold;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 50px;
                border-left-width: 2px;
                border-left-color: #CCCCCC;
                border-left-style: solid;
                border-top-right-radius: 8px;
                border-bottom-right-radius: 8px;
            }
            QComboBox QAbstractItemView {
                background-color: #FFFFFF;
                color: #333333;
                selection-background-color: #FF5722;
                border: 2px solid #CCCCCC;
                font-size: 16px;
                padding: 10px;
            }
        """)

        self.convert_btn = QPushButton("START CONVERSION")
        self.convert_btn.setIcon(QIcon.fromTheme("media-playback-start"))
        self.convert_btn.setStyleSheet(self.get_button_style(primary=True))
        self.convert_btn.setMinimumHeight(60)
        self.convert_btn.setEnabled(False)
        self.convert_btn.clicked.connect(self.start_conversion)

        settings_layout.addWidget(format_label)
        settings_layout.addWidget(self.format_combo)
        settings_layout.addSpacing(15)
        settings_layout.addWidget(self.convert_btn)
        settings_card.setLayout(settings_layout)

        # Progress area
        progress_card = QFrame()
        progress_card.setFrameShape(QFrame.StyledPanel)
        progress_card.setStyleSheet("""
            QFrame {
                background-color: #FFFFFF;
                border-radius: 15px;
                padding: 25px;
                border: 2px solid #E0E0E0;
            }
        """)
        
        progress_layout = QVBoxLayout()
        progress_layout.setSpacing(20)

        self.progress = AnimatedProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setStyleSheet("""
            QProgressBar {
                border: 2px solid #CCCCCC;
                border-radius: 10px;
                background-color: #F5F5F5;
                height: 35px;
            }
            QProgressBar::chunk {
                border-radius: 10px;
                background: qlineargradient(
                    spread:pad, x1:0, y1:0.5, x2:1, y2:0.5, 
                    stop:0 #FF5722, stop:1 #FF9800
                );
            }
        """)
        self.progress.setVisible(False)

        self.progress_label = QLabel()
        self.progress_label.setAlignment(Qt.AlignCenter)
        self.progress_label.setStyleSheet("""
            font-size: 16px;
            color: #333333;
            font-weight: bold;
        """)

        progress_layout.addWidget(self.progress)
        progress_layout.addWidget(self.progress_label)
        progress_card.setLayout(progress_layout)

        # Add widgets to main layout
        main_layout.addWidget(header)
        main_layout.addWidget(file_card)
        main_layout.addWidget(settings_card)
        main_layout.addWidget(progress_card)
        main_layout.addStretch()

        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

    def setup_styles(self):
        # Set vibrant gradient background
        self.setAutoFillBackground(True)
        palette = self.palette()
        gradient = QLinearGradient(0, 0, self.width(), self.height())
        gradient.setColorAt(0, QColor(33, 150, 243))  # Bright blue
        gradient.setColorAt(1, QColor(244, 67, 54))   # Vibrant red
        palette.setBrush(QPalette.Window, QBrush(gradient))
        self.setPalette(palette)

        # Set application font
        font = QFont()
        font.setFamily("Segoe UI")
        font.setPointSize(12)
        font.setWeight(QFont.Bold)
        self.setFont(font)

    def get_button_style(self, primary=False):
        if primary:
            return """
                QPushButton {
                    background: qlineargradient(
                        spread:pad, x1:0, y1:0, x2:1, y2:0, 
                        stop:0 #FF5722, stop:1 #FF9800
                    );
                    color: #FFFFFF;
                    border: none;
                    border-radius: 10px;
                    padding: 15px 30px;
                    font-weight: bold;
                    font-size: 18px;
                    text-transform: uppercase;
                }
                QPushButton:hover {
                    background: qlineargradient(
                        spread:pad, x1:0, y1:0, x2:1, y2:0, 
                        stop:0 #E64A19, stop:1 #F57C00
                    );
                }
                QPushButton:pressed {
                    background: qlineargradient(
                        spread:pad, x1:0, y1:0, x2:1, y2:0, 
                        stop:0 #D84315, stop:1 #EF6C00
                    );
                }
                QPushButton:disabled {
                    background: #BDBDBD;
                    color: #757575;
                }
            """
        else:
            return """
                QPushButton {
                    background-color: #FFFFFF;
                    color: #FF5722;
                    border: 2px solid #FF5722;
                    border-radius: 10px;
                    padding: 15px 30px;
                    font-size: 18px;
                    font-weight: bold;
                    text-transform: uppercase;
                }
                QPushButton:hover {
                    background-color: #FF5722;
                    color: #FFFFFF;
                }
                QPushButton:pressed {
                    background-color: #E64A19;
                    color: #FFFFFF;
                }
            """

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        supported_formats = ['.wmv', '.asf', '.mp4', '.mov', '.avi', '.mkv']
        if files and any(files[0].lower().endswith(ext) for ext in supported_formats):
            self.input_file = files[0]
            self.file_label.setText(f"SELECTED: {os.path.basename(self.input_file).upper()}")
            self.convert_btn.setEnabled(True)
            self.progress_label.setText("READY TO CONVERT")
            self.progress_label.setStyleSheet("color: #4CAF50; font-weight: bold; font-size: 16px;")
        else:
            self.progress_label.setText("❌ INVALID FILE TYPE - USE .WMV, .ASF, .MP4, .MOV, .AVI OR .MKV")
            self.progress_label.setStyleSheet("color: #F44336; font-weight: bold; font-size: 16px;")

    def select_file(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "Select Media File", "", 
            "Windows Media Files (*.wmv *.asf);;Video Files (*.mp4 *.mov *.avi *.mkv);;All Files (*)"
        )
        if file:
            self.input_file = file
            self.file_label.setText(f"SELECTED: {os.path.basename(file).upper()}")
            self.convert_btn.setEnabled(True)
            self.progress_label.setText("READY TO CONVERT")
            self.progress_label.setStyleSheet("color: #4CAF50; font-weight: bold; font-size: 16px;")

    def start_conversion(self):
        if not self.input_file or not os.path.exists(self.input_file):
            QMessageBox.critical(self, "Error", "Input file does not exist!")
            return
            
        # Get format without the "(Recommended)" text if present
        output_format = self.format_combo.currentText().split()[0].lower()
        output_file = os.path.splitext(self.input_file)[0] + '.' + output_format
        
        self.progress.setVisible(True)
        self.progress.setValueAnimated(0)
        self.progress_label.setText("CONVERTING... 0%")
        self.progress_label.setStyleSheet("color: #FF9800; font-weight: bold; font-size: 16px;")
        self.convert_btn.setEnabled(False)
        
        self.converter_thread = ConverterThread(self.input_file, output_file)
        self.converter_thread.progress_updated.connect(self.update_progress)
        self.converter_thread.conversion_finished.connect(self.conversion_complete)
        self.converter_thread.start()

    def update_progress(self, value):
        self.progress.setValueAnimated(value)
        self.progress_label.setText(f"CONVERTING... {value}%")
        if value < 30:
            self.progress_label.setStyleSheet("color: #FF9800; font-weight: bold; font-size: 16px;")
        elif value < 70:
            self.progress_label.setStyleSheet("color: #FF5722; font-weight: bold; font-size: 16px;")
        else:
            self.progress_label.setStyleSheet("color: #4CAF50; font-weight: bold; font-size: 16px;")

    def conversion_complete(self, success, message):
        self.progress.setVisible(False)
        self.convert_btn.setEnabled(True)
        
        if success:
            self.progress_label.setText("✅ CONVERSION SUCCESSFUL!")
            self.progress_label.setStyleSheet("color: #4CAF50; font-weight: bold; font-size: 16px;")
            QMessageBox.information(
                self, 
                "SUCCESS", 
                f"File converted successfully:\n{message}",
                QMessageBox.Ok
            )
        else:
            self.progress_label.setText(f"❌ CONVERSION FAILED: {message}")
            self.progress_label.setStyleSheet("color: #F44336; font-weight: bold; font-size: 16px;")
            QMessageBox.critical(
                self, 
                "ERROR", 
                f"Conversion failed:\n{message}",
                QMessageBox.Ok
            )

if __name__ == "__main__":
    # Check for required Python packages only
    try:
        import cv2
        import moviepy.editor as mpy
        from PyQt5.QtWidgets import QApplication
    except ImportError as e:
        app = QApplication(sys.argv)
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setWindowTitle("Missing Dependencies")
        msg.setText("Required Python packages not installed")
        msg.setInformativeText(
            "Please install these packages first:\n\n"
            "pip install opencv-python moviepy numpy PyQt5\n\n"
            f"Error: {str(e)}"
        )
        msg.setStyleSheet("""
            QMessageBox {
                background-color: #FFFFFF;
                font-weight: bold;
            }
            QLabel {
                color: #333333;
                font-size: 14px;
            }
            QPushButton {
                background-color: #FF5722;
                color: white;
                border: none;
                padding: 8px 16px;
                min-width: 100px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #E64A19;
            }
        """)
        msg.exec_()
        sys.exit(1)
        
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # Set application-wide font
    font = QFont()
    font.setFamily("Segoe UI")
    font.setPointSize(12)
    font.setWeight(QFont.Bold)
    app.setFont(font)
    
    window = MediaConverterApp()
    window.show()
    sys.exit(app.exec_())