from docx import Document
import tkinter as tk
from tkinter import filedialog
import os

def add_symbols_to_superscript_subscript(doc):
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if run.font.subscript:
                run.text = f'__{run.text}'  # Add your desired symbol for subscript here
            elif run.font.superscript:
                run.text = f'^^{run.text}'  # Add your desired symbol for superscript here

def process_document(file_path):
    try:
        # Load the Word document
        doc = Document(file_path)

        # Add symbols to superscript and subscript
        add_symbols_to_superscript_subscript(doc)

        # Create a modified file path
        base_path, file_name = os.path.split(file_path)
        modified_file_name = 'modified_' + file_name
        modified_file_path = os.path.join(base_path, modified_file_name)

        # Save the modified document
        doc.save(modified_file_path)

        print(f"Modified document saved to: {modified_file_path}")

    except Exception as e:
        print(f"An error occurred: {e}")

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        process_document(file_path)

# Create the main window
window = tk.Tk()
window.title("Word Document Modifier")

# Set window size and color
window.geometry("400x200")
window.configure(bg="#f0f0f0")  # Set background color

# Create and configure widgets
label = tk.Label(window, text="Select a Word document:", bg="#f0f0f0", fg="#333333")
label.pack(pady=10)

browse_button = tk.Button(window, text="Browse", command=browse_file, bg="#4CAF50", fg="white")
browse_button.pack(pady=10)
window.resizable(False, False)
# Start the main loop
window.mainloop()
