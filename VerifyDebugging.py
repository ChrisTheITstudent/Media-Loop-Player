import win32com.client
import time

ppt = win32com.client.DispatchEx("PowerPoint.Application")
ppt.Visible = True

file_path = r"C:\MediaFiles\Test File.pptx"

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
