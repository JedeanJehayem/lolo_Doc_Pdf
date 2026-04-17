import sys
import os
import random
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QPixmap, QGuiApplication
from PyQt6.QtWidgets import QApplication, QWidget, QLabel


class WallpaperLoading(QWidget):
    def __init__(self, image_path=None, duration_ms=60000):
        super().__init__()

        self.image_path = self.resolve_image(image_path)
        self.duration_ms = duration_ms

        self.vel_x = random.randint(2, 3)
        self.vel_y = random.randint(2, 3)

        self.img_w = 400
        self.img_h = 400

        self.setup_window()
        self.setup_image()
        self.start_animation()
        self.start_auto_close()

    def resolve_image(self, image_path):
        if image_path and os.path.exists(image_path):
            return image_path

        pasta = os.getcwd()
        imagens = [
            f for f in os.listdir(pasta)
            if f.lower().endswith((".png", ".jpg", ".jpeg"))
        ]

        if imagens:
            return os.path.join(pasta, random.choice(imagens))

        raise FileNotFoundError(
            "Nenhuma imagem encontrada. Passe um caminho válido ou coloque uma imagem na pasta."
        )

    def setup_window(self):
        screen = QGuiApplication.primaryScreen()
        geometry = screen.geometry()

        self.setGeometry(geometry)

        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.Tool
            | Qt.WindowType.WindowStaysOnTopHint
        )

        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, True)

        self.showFullScreen()

    def setup_image(self):
        self.label = QLabel(self)
        self.label.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)

        pixmap = QPixmap(self.image_path)

        if pixmap.isNull():
            raise ValueError(f"Erro ao carregar imagem: {self.image_path}")

        pixmap = pixmap.scaled(
            self.img_w,
            self.img_h,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation
        )

        self.label.setPixmap(pixmap)
        self.label.resize(pixmap.width(), pixmap.height())

        self.img_w = pixmap.width()
        self.img_h = pixmap.height()

        self.pos_x = max(0, (self.width() - self.img_w) // 2)
        self.pos_y = max(0, (self.height() - self.img_h) // 2)

        self.label.move(self.pos_x, self.pos_y)
        self.label.show()

    def start_animation(self):
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.animate)
        self.timer.start(16)

    def animate(self):
        self.pos_x += self.vel_x
        self.pos_y += self.vel_y

        if self.pos_x <= 0:
            self.pos_x = 0
            self.vel_x *= -1
        elif self.pos_x + self.img_w >= self.width():
            self.pos_x = self.width() - self.img_w
            self.vel_x *= -1

        if self.pos_y <= 0:
            self.pos_y = 0
            self.vel_y *= -1
        elif self.pos_y + self.img_h >= self.height():
            self.pos_y = self.height() - self.img_h
            self.vel_y *= -1

        self.label.move(self.pos_x, self.pos_y)

    def start_auto_close(self):
        self.close_timer = QTimer(self)
        self.close_timer.setSingleShot(True)
        self.close_timer.timeout.connect(self.close)
        self.close_timer.start(self.duration_ms)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.close()

    def closeEvent(self, event):
        self.timer.stop()
        event.accept()


def main():
    app = QApplication(sys.argv)

    image_path = None
    duration_ms = 60000

    if len(sys.argv) > 1:
        image_path = sys.argv[1]

    if len(sys.argv) > 2:
        try:
            duration_ms = int(sys.argv[2])
        except ValueError:
            pass

    window = WallpaperLoading(image_path=image_path, duration_ms=duration_ms)
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()