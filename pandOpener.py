import pathlib
import tkinter as tk
from tkinter import *
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog
from tkinter import font
from tkinter import colorchooser
import os, sys
import win32print
import win32api

class Notepad(tk.Tk):
    """Minimalist notepad app"""
    def __init__(self):
        super().__init__()
        self.title("Panda Document Editor")
        self.menubar = tk.Menu(self, tearoff=False)
        self['menu'] = self.menubar
        self.menu_file = tk.Menu(self.menubar, tearoff=False)
        self.menu_edit = tk.Menu(self.menubar, tearoff=False)
        self.menu_preferences = tk.Menu(self.menubar, tearoff=False)
        self.menu_tools = tk.Menu(self.menubar, tearoff=False)
        self.menu_help = tk.Menu(self.menubar, tearoff=False)
        self.menubar.add_cascade(menu=self.menu_file, label='File')
        self.menubar.add_cascade(menu=self.menu_edit, label='Edit')
        self.menubar.add_cascade(menu=self.menu_tools, label='Tools')
        self.menubar.add_cascade(menu=self.menu_help, label='Help')

        #File Menu Commands
        self.menu_file.add_command(label='New', accelerator='Ctrl+N', command=self.new_file)
        self.menu_file.add_command(label='Open', accelerator='Ctrl+O', command=self.open_file)
        self.menu_file.add_command(label='Save', accelerator='Ctrl+S', command=self.save_file)
        self.menu_file.add_command(label='Save As', command=self.save_file_as)
        self.menu_file.add_separator()
        self.menu_file.add_command(label='Exit', command=self.destroy)

        #Edit Menu Commands
        self.menu_edit.add_command(label='Cut', accelerator='Ctrl+X', command=self.cut_text)
        self.menu_edit.add_command(label='Copy', accelerator='Ctrl+C', command=self.copy_text)
        self.menu_edit.add_command(label='Paste', accelerator='Ctrl+V', command=self.paste_text)
        self.menu_edit.add_separator()
        self.menu_edit.add_command(label='Bold', accelerator='Ctrl+B', command=self.bold_text)
        self.menu_edit.add_command(label='Italicize', accelerator='Ctrl+I', command=self.italicize_text)
        self.menu_edit.add_command(label='Underline', accelerator='Ctrl+I', command=self.underline_text)
        self.menu_edit.add_separator()

        #Preferences Menu Commands
        self.dark_mode=IntVar()
        self.dark_mode.set(0)
        self.menu_edit.add_cascade(menu=self.menu_preferences, label='Preferences')
        self.menu_preferences.add_checkbutton(label='Dark Mode', variable=self.dark_mode, accelerator='Ctrl+D', command=self.dark_mode)

        #Tool Menu Commands
        self.menu_tools.add_command(label='Word Count', command=self.word_count)
        
        #About Menu Commands
        self.menu_help.add_command(label='About', command=self.about_me)

        self.info_var = tk.StringVar()
        self.info_var.set('>  New File  <')
        self.bar_color = tk.StringVar()
        self.bar_color.set('#ff0000')
        self.info_bar = tk.Label(self, textvariable=self.info_var, bg=bar_color, fg='ffffff')
        self.info_bar.configure(anchor=tk.W, font='-size -14 -weight bold', padx=5, pady=5)
        self.text = ScrolledText(self, font='-size -16')
        self.info_bar.pack(side=tk.TOP, fill=tk.X)
        self.text.pack(fill=tk.BOTH, expand=tk.YES)
        self.file = None

        #File Menu Keybinds
        self.bind("<Control-n>", self.new_file)
        self.bind("<Control-s>", self.save_file)
        self.bind("<Control-o>", self.open_file)

        #Edit Menu Keybinds
        self.bind("<Control-x>", self.cut_text)
        self.bind("<Control-c>", self.copy_text)
        self.bind("<Control-v>", self.paste_text)

        self.bind("<Control-b>", self.bold_text)
        self.bind("<Control-i>", self.italicize_text)
        self.bind("<Control-u>", self.underline_text)
        
        self.bind("<Control-d>", self.dark_mode)

    def open_file(self, event=None):
        """Open file and update infobar"""
        file = filedialog.askopenfilename(title='Open', defaultextension=".pand", filetypes=(('Panda Documents', '*.pand'),('All Files', '*.*')))
        if file:
            self.file = pathlib.Path(file)
            self.text.delete('1.0', tk.END)
            self.text.insert(tk.END, self.file.read_text())
            self.info_var.set(self.file.absolute())

    def new_file(self, event=None):
        """Reset body and clear variables"""
        self.file = None
        self.text.delete('1.0', tk.END)
        self.info_var.set('>  New File  <')

    def save_file(self, event=None):
        """Save file instantly, otherwise use Save As method"""
        if self.file:
            text = self.text.get('1.0', tk.END)
            self.file.write_text(text)
        else:
            self.save_file_as()

    def save_file_as(self):
        """Save new file or existing file to new name or location"""
        file = filedialog.asksaveasfilename(title="Save", defaultextension=".pand", filetypes=(('Panda Documents', '*.pand'),('All Files', '*.*')))
        if file:
            self.file = pathlib.Path(file)
            text = self.text.get('1.0', tk.END)
            self.file.write_text(text)
            self.info_var.set(self.file.absolute())

    #Edit Menu
    def cut_text(self):
        global selected
        if self.text.selection_get():
            selected = self.text.selection_get()
            self.text.delete("sel.first", "sel.last")
            self.clipboard_clear()
            self.clipboard_append(selected)

    def copy_text(self):
        global selected
        if self.text.selection_get():
            selected = self.text.selection_get()
            self.clipboard_clear()
            self.clipboard_append(selected)

    def paste_text(self):
        if selected:
            position = self.text.index(INSERT)
            self.text.insert(position, selected)

    def bold_text(self):
        bold_font = font.Font(self.text, self.text.cget("font"))
        bold_font.configure(weight="bold")
        self.text.tag_configure("bold", font=bold_font)
        current_tags = self.text.tag_names("sel.first")

        if "bold" in current_tags:
            self.text.tag_remove("bold", "sel.first", "sel.last")
        else:
            self.text.tag_add("bold", "sel.first", "sel.last")

    def italicize_text(self):
        italic_font = font.Font(self.text, self.text.cget("font"))
        italic_font.configure(slant="italic")
        self.text.tag_configure("italic", font=italic_font)
        current_tags = self.text.tag_names("sel.first")

        if "italic" in current_tags:
            self.text.tag_remove("italic", "sel.first", "sel.last")
        else:
            self.text.tag_add("italic", "sel.first", "sel.last")

    def underline_text(self):
        underline_font = font.Font(self.text, self.text.cget("font"))
        underline_font.configure(underline=True)
        self.text.tag_configure("underline", font=underline_font)
        current_tags = self.text.tag_names("sel.first")

        if "underline" in current_tags:
            self.text.tag_remove("underline", "sel.first", "sel.last")
        else:
            self.text.tag_add("underline", "sel.first", "sel.last")

    def dark_mode(self):
        global text_color
        global bar_color
        global bg_color

        if(dark_toggle == 0):
            dark_toggle = 1

            bar_color = "#ff0000"
            bg_color = "#1c1c1c"
            text_color = "ffffff"

            self.bg.config(bg=bg_color)
        else:
            dark_toggle = 0

            bar_color = "#ff0000"
            bg_color = "ffffff"
            text_color = "000000"

            self.config(bg=bg_color)

    #Tool Menu
    def word_count(self):
        """Display estimated word count"""
        words = list(self.text.get('1.0', tk.END).split(' '))
        word_count = len(words)
        messagebox.showinfo(title='Word Count', message=f'Word Count: {word_count:,d}')

    #About Menu
    def about_me(self):
        """Short pithy quote"""
        text = 'Panda Documents, extension .pand, store notes like .txt files, but with a password protected twist!'
        messagebox.showinfo(title="About .pand", message=text)

if __name__ == '__main__':
    app = Notepad()
    app.mainloop()