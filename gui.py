from tkinter import *
import subprocess
from PIL import Image
from PIL import ImageTk
import time

class SimpleApp(object):
    def __init__(self, master, filename, **kwargs):
        self.master = master
        self.filename = filename
        self.canvas = Canvas(master, width=500, height=500)
        self.canvas.pack()

        self.update = self.draw().__next__
        master.after(100, self.update)

    def draw(self):
        image = Image.open(self.filename)
        angle = 0
        while True:
            tkimage = ImageTk.PhotoImage(image.rotate(angle))
            canvas_obj = self.canvas.create_image(
                250, 250, image=tkimage)
            self.master.after_idle(self.update)
            yield
            self.canvas.delete(canvas_obj)
            angle += 1
            angle %= 360

root = Tk()
root.title('Alex - The Virtual Assistant')
root.iconbitmap(r'./icon.ico')
subprocess.Popen(['python', './engine.py'], shell = True)
app = SimpleApp(root, 'main.png')
root.mainloop()
