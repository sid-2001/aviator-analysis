




import pytesseract
from PIL import ImageGrab
import openpyxl
import tkinter as tk

class RegionSelector(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Select Region")
        self.attributes('-fullscreen', True)
        self.attributes('-alpha', 0.5)  # Make the window semi-transparent
        # self.configure(bg='gray')  # Background color to help with transparency
        self.canvas = tk.Canvas(self, cursor="cross", bg='gray')
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.start_x = None
        self.start_y = None
        self.rect = None

        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)

    def on_button_press(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red', width=2)

    def on_mouse_drag(self, event):
        cur_x, cur_y = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)

    def on_button_release(self, event):
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)
        self.region = (self.start_x, self.start_y, end_x, end_y)
        self.destroy()

def get_screen_region():
    app = RegionSelector()
    app.mainloop()
    return app.region

if __name__ == "__main__":
    region = get_screen_region()
    print(f"Selected region: {region}")

    # Capture the specified region of the screen
    screenshot = ImageGrab.grab(bbox=region)

    # Perform OCR on the captured image
    text = pytesseract.image_to_string(screenshot)

    # Create a new Excel workbook and select the active worksheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Write the extracted text to the first cell
    sheet['A1'] = text

    # Save the workbook to a file
    workbook.save('extracted_text.xlsx')

    print('Text extracted and saved to extracted_text.xlsx')
