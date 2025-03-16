import os
import sys
import time
import cv2
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import win32com.client

# Constants
SUPPORTED_VIDEO_FORMATS = (".mp4", ".avi", ".mov")
SUPPORTED_PPT_FORMATS = (".ppt", ".pptx")


class MediaManager:
    """Handles file management for media playback."""
    
    def __init__(self, folder):
        self.folder = folder
        self.files = []
        self.load_files()

    def load_files(self):
        """Loads video and PowerPoint files from the folder."""
        self.files = [f for f in os.listdir(self.folder) if f.endswith(SUPPORTED_VIDEO_FORMATS + SUPPORTED_PPT_FORMATS)]
    
    def get_files(self):
        """Returns the list of files."""
        return self.files


class VideoPlayer:
    """Handles video playback using OpenCV."""
    
    def __init__(self):
        self.playing = False
        self.cap = None
        self.stop_flag = False

    def play(self, file_path):
        """Plays a video in full-screen mode."""
        self.stop_flag = False
        self.playing = True
        self.cap = cv2.VideoCapture(file_path)
        cv2.namedWindow("Video", cv2.WND_PROP_FULLSCREEN)
        cv2.setWindowProperty("Video", cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)

        while self.playing and self.cap.isOpened():
            if self.stop_flag:
                break

            ret, frame = self.cap.read()
            if not ret:
                break

            cv2.imshow("Video", frame)
            if cv2.waitKey(25) & 0xFF == ord("q"):
                break

        self.stop()

    def stop(self):
        """Stops video playback."""
        self.playing = False
        self.stop_flag = True
        if self.cap is not None:
            self.cap.release()
            self.cap = None
        cv2.destroyAllWindows()

    def close(self):
            """Clean up resources."""
            self.stop()
            cv2.destroyAllWindows()

class PowerPointPlayer:
    """Handles PowerPoint file playback."""
    
    def __init__(self):
        try:
            self.app = win32com.client.DispatchEx("PowerPoint.Application")  # Create a new instance
            self.presentation = None
            print("PowerPoint initialized successfully.")
        except Exception as e:
            print(f"Failed to initialize PowerPoint: {e}")

    def play(self, file_path):
        """Opens and plays a PowerPoint file in full-screen mode."""
        ppt = win32com.client.DispatchEx("PowerPoint.Application")
        ppt.Visible = True

        try:
            print(f"Opening file: {file_path}")
            presentation = ppt.Presentations.Open(file_path, ReadOnly=1, WithWindow=True)

            print("Configuring slideshow settings...")
            presentation.SlideShowSettings.StartingSlide = 1
            presentation.SlideShowSettings.EndingSlide = presentation.Slides.Count
            presentation.SlideShowSettings.AdvanceMode = 2  # Auto Advance
            presentation.SlideShowSettings.LoopUntilStopped = True
            presentation.SlideShowSettings.ShowWithNarration = False
            presentation.SlideShowSettings.ShowWithAnimation = True
            presentation.SlideShowSettings.ShowType = 3  # Full-screen

            # Set slide transition timings (5 seconds per slide)
            for slide in presentation.Slides:
                slide.SlideShowTransition.AdvanceOnTime = True
                slide.SlideShowTransition.AdvanceTime = 5  # Adjust seconds per slide

            print("Starting slideshow...")
            presentation.SlideShowSettings.Run()

            time.sleep(2)  # Give time for the slideshow to start

            # Wait for the slideshow to finish
            while True:
                # If the slideshow window is open, check the current slide
                if presentation.SlideShowWindow is not None:
                    current_slide = presentation.SlideShowWindow.View.Slide
                    total_slides = presentation.Slides.Count

                    # Check if we're on the last slide
                    if current_slide.SlideIndex == total_slides:
                        print("Last slide finished playing.")
                        break
                    
                time.sleep(1)  # Wait a second before checking again

            # Automatically close the presentation and quit PowerPoint
            presentation.Close()
            ppt.Quit()

            print("PowerPoint closed successfully.")

        except Exception as e:
            print(f"Error: {e}")

    def stop(self):
        """Stops PowerPoint playback."""
        try:
            if self.presentation:
                self.presentation.Close()
                self.presentation = None
                print("PowerPoint closed successfully.")
        except Exception as e:
            print(f"Error closing PowerPoint: {e}")

    def close(self):
        """Clean up resources."""
        self.stop()
        self.app.Quit()


class FileWatcher(FileSystemEventHandler):
    """Monitors the folder for file changes."""
    
    def __init__(self, media_manager, ui_callback):
        self.media_manager = media_manager
        self.ui_callback = ui_callback

    def on_modified(self, event):
        """Triggers when files are added or removed."""
        self.media_manager.load_files()
        self.ui_callback()


class MediaApp:
    """Main GUI application."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Media Player")
        self.root.geometry("500x400")
        if getattr(sys, 'frozen', False):
            # Running as a bundled executable
            icon_path = os.path.join(sys._MEIPASS, "Toad.ico")
        else:
            # Running from the script
            icon_path = "Toad.ico"

        self.root.iconbitmap(icon_path)

        self.media_manager = None
        self.video_player = VideoPlayer()
        self.ppt_player = PowerPointPlayer()
        self.current_index = 0
        self.playing = False

        self.create_ui()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def create_ui(self):
        """Creates the GUI elements."""
        button_width = 20
        button_height = 2

        self.listbox = tk.Listbox(self.root, width=50, height=10)
        self.listbox.pack(pady=10)

        self.select_folder_button = tk.Button(self.root, text="Select Folder", command=self.select_folder, width=button_width, height=button_height)
        self.select_folder_button.pack(pady=5)

        self.play_button = tk.Button(self.root, text="Start", command=self.start_playback, width=button_width, height=button_height)
        self.play_button.pack(pady=5)

        self.stop_button = tk.Button(self.root, text="Stop", command=self.stop_playback, width=button_width, height=button_height)
        self.stop_button.pack(pady=5)

        self.refresh_button = tk.Button(self.root, text="Refresh Files", command=self.refresh_files, width=button_width, height=button_height)
        self.refresh_button.pack(pady=5)

        self.refresh_files()

    def select_folder(self):
        """Opens a dialog for the user to select a folder."""
        folder_path = filedialog.askdirectory(title="Select Media Folder")
        if folder_path:
            self.media_manager = MediaManager(folder_path)
            self.start_file_watcher()
            self.refresh_files()
            self.start_file_watcher()

    def refresh_files(self):
        """Refreshes the file list in the UI."""
        if self.media_manager is None:
            return
        
        self.listbox.delete(0, tk.END)
        for file in self.media_manager.get_files():
            self.listbox.insert(tk.END, file)

    def start_playback(self):
        """Starts playing files in sequence."""
        if not self.media_manager.get_files():
            messagebox.showerror("Error", "No media files found!")
            return

        self.playing = True
        threading.Thread(target=self.play_files, daemon=True).start()

    def stop_playback(self):
        """Stops playback."""
        self.playing = False
        self.video_player.stop()
        self.ppt_player.stop()

    def play_files(self):
        """Plays each file in a loop."""
        while self.playing:
            files = self.media_manager.get_files()
            if not files:
                break

            file_path = os.path.join(self.media_manager.folder, files[self.current_index])

            if file_path.endswith(SUPPORTED_VIDEO_FORMATS):
                self.video_player.play(file_path)
            elif file_path.endswith(SUPPORTED_PPT_FORMATS):
                self.ppt_player.play(file_path)

            self.current_index = (self.current_index + 1) % len(files)

    def start_file_watcher(self):
        """Starts a file watcher thread."""
        event_handler = FileWatcher(self.media_manager, self.refresh_files)
        observer = Observer()
        observer.schedule(event_handler, self.media_manager.folder, recursive=False)
        observer.start()

    def on_close(self):
        """Handles cleanup when the window is closed."""
        print("Closing application...")

        # Stop playback and cleanup resources
        self.stop_playback()

        # Ensure all resources are released
        self.video_player.close()
        self.ppt_player.close()

        # Destroy the window
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = MediaApp(root)
    root.mainloop()