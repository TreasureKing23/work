import customtkinter as ctk
import tkinter as tk

# Setup appearance
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")


class CustomWindow(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Remove default title bar
        self.overrideredirect(False)
        self.geometry("600x400")
        self.configure(bg="silver")  # Silver border effect

        # Create a black interior body
        self.main_frame = ctk.CTkFrame(self, fg_color="black", corner_radius=0)
        self.main_frame.pack(expand=True, fill="both", padx=2, pady=2)

        # Custom Title Bar
        self.title_bar = tk.Frame(self.main_frame, bg="black", height=30)
        self.title_bar.pack(fill="x", side="top")

        # Add the 3 mac-style dots
        self.create_title_buttons()

        # Allow dragging the window
        self.title_bar.bind("<Button-1>", self.get_pos)
        self.title_bar.bind("<B1-Motion>", self.move_window)

    def create_title_buttons(self):
        # Button settings
        btn_size = 14
        pad_x = 6



        # Red (Close)
        self.red_btn = tk.Canvas(self.title_bar, width=btn_size, height=btn_size, bg="black", highlightthickness=0)
        self.red_btn.create_oval(0, 0, btn_size, btn_size, fill="#ff5f56", outline="")
        self.red_btn.pack(side="right", padx=pad_x)
        self.red_btn.bind("<Button-1>", lambda e: self.destroy())

        
        # Green (Maximize)
        self.green_btn = tk.Canvas(self.title_bar, width=btn_size, height=btn_size, bg="black", highlightthickness=0)
        self.green_btn.create_oval(0, 0, btn_size, btn_size, fill="#00ff00", outline="")
        self.green_btn.pack(side="right", padx=pad_x)
        self.green_btn.bind("<Button-1>", lambda e: self.toggle_fullscreen())

       
        

         # Yellow (Minimize)
        self.yellow_btn = tk.Canvas(self.title_bar, width=btn_size, height=btn_size, bg="black", highlightthickness=0)
        self.yellow_btn.create_oval(0, 0, btn_size, btn_size, fill="#ffff00", outline="")
        self.yellow_btn.pack(side="right", padx=pad_x)
        self.yellow_btn.bind("<Button-1>", lambda e: self.minimize_window())


    def minimize_window(self):
        self.overrideredirect(False)
        self.iconify()
        self.bind("<Map>", lambda e: self.overrideredirect(True))

    def get_pos(self, event):
        self.xwin = event.x
        self.ywin = event.y

    def move_window(self, event):
        self.geometry(f'+{event.x_root - self.xwin}+{event.y_root - self.ywin}')

    def toggle_fullscreen(self):
        self.attributes("-fullscreen", not self.attributes("-fullscreen"))


if __name__ == "__main__":
    app = CustomWindow()
    app.mainloop()