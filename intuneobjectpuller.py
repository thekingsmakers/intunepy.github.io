import pandas as pd
import logging
from tkinter import Tk, filedialog, Button, Label, Entry, scrolledtext, messagebox, PhotoImage, Frame, Text, END
import os
import webbrowser

# Set up logging
logging.basicConfig(filename='compare_and_export.log', level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def browse_file(entry_widget, file_type):
    """Open a file dialog to select a file."""
    file_path = filedialog.askopenfilename(title=f"Select {file_type} file", filetypes=[("Excel files", "*.xlsx")])
    entry_widget.delete(0, END)  # Clear the entry widget
    entry_widget.insert(0, file_path)  # Insert the selected file path
    logging.info(f'{file_type} file uploaded: {file_path}')

def select_output_folder():
    """Open a folder dialog to select a folder."""
    folder_path = filedialog.askdirectory(title="Select folder to save output file")
    logging.info(f'Output folder selected: {folder_path}')
    return folder_path

def display_logs(log_text_widget):
    """Display logs in the Text widget."""
    with open('compare_and_export.log', 'r') as log_file:
        log_text_widget.config(state="normal")
        log_text_widget.delete(1.0, END)  # Clear the Text widget
        log_text_widget.insert(END, log_file.read())  # Insert log file content
        log_text_widget.config(state="disabled")

def open_twitter():
    """Open Twitter link."""
    webbrowser.open_new("https://x.com/thekingsmakers")

def show_readme(readme_text_widget, log_text_widget):
    """Show README and hide logs."""
    log_text_widget.grid_remove()
    readme_text_widget.grid()

def hide_readme(readme_text_widget, log_text_widget):
    """Hide README and show logs."""
    readme_text_widget.grid_remove()
    log_text_widget.grid()

def compare_and_export(file_export, file_compare, output_folder, log_text_widget, readme_text_widget):
    hide_readme(readme_text_widget, log_text_widget)
    logging.info('Starting comparison process.')

    try:
        # Load the Excel files
        df_export = pd.read_excel(file_export)
        df_compare = pd.read_excel(file_compare)

        # Normalize DeviceName to ensure matching
        df_export['DeviceName'] = df_export['DeviceName'].str.strip().str.lower()
        df_compare['DeviceName'] = df_compare['DeviceName'].str.strip().str.lower()

        # Convert 'approximateLastSignInDateTime' to datetime and sort
        df_export['approximateLastSignInDateTime'] = pd.to_datetime(df_export['approximateLastSignInDateTime'])
        df_export.sort_values(by='approximateLastSignInDateTime', ascending=False, inplace=True)

        # Drop duplicates, keeping the last sign-in date
        df_export = df_export.drop_duplicates(subset='DeviceName', keep='first')

        # Find matching DeviceNames
        matching_devices = df_export[df_export['DeviceName'].isin(df_compare['DeviceName'])]

        if matching_devices.empty:
            logging.warning("No matching devices found.")
            messagebox.showwarning("Warning", "No matching devices found.")
            return

        # Define the output file path
        output_file = os.path.join(output_folder, 'output.xlsx')

        # Write the output to a new Excel file
        matching_devices.to_excel(output_file, index=False)
        logging.info(f'Comparison completed. Results saved at: {output_file}')
        messagebox.showinfo("Success", f"Comparison completed and file saved at: {output_file}")
    except Exception as e:
        logging.error(f"An error occurred: {e}")
        messagebox.showerror("Error", f"An error occurred: {e}")
    
    display_logs(log_text_widget)  # Refresh logs after operation

# Create the main application window
root = Tk()
root.title("Intune Compare and Export Tool")

# Create a frame for the logo and credits
top_frame = Frame(root)
top_frame.grid(row=0, column=0, columnspan=3, pady=10)

# Add Intune logo
try:
    logo = PhotoImage(file='intune_logo.png')
    logo_label = Label(top_frame, image=logo)
    logo_label.grid(row=0, column=1, padx=10)
except Exception as e:
    logging.error(f"Logo not found: {e}")
    logo_label = Label(top_frame, text="Intune Tool", font=("Arial", 20))
    logo_label.grid(row=0, column=1, padx=10)

# Add credits and Twitter link
credits_frame = Frame(top_frame)
credits_frame.grid(row=0, column=0, sticky="e")

credits_label = Label(credits_frame, text="Credits: Omar Osman", font=("Arial", 10))
credits_label.grid(row=0, column=0, sticky="w")

twitter_label = Label(credits_frame, text="Follow on Twitter", font=("Arial", 10, "underline"), fg="blue", cursor="hand2")
twitter_label.grid(row=1, column=0, sticky="w")
twitter_label.bind("<Button-1>", lambda e: open_twitter())

# Add a README section
readme_text = Text(root, width=80, height=20, state="normal")
readme_text.grid(row=6, column=0, columnspan=3, padx=5, pady=5)
readme_text.insert(END, "Welcome to the Intune Compare and Export Tool!\n\n"
                       "This tool allows Intune admins to compare device lists and export matching results.\n\n"
                       "Instructions:\n"
                       "1. Select the export file.\n"
                       "2. Select the compare file.\n"
                       "3. Choose an output folder.\n"
                       "4. Click 'Select Output Folder' to start the comparison.\n\n"
                       "Credits: Omar Osman\n"
                       "Twitter: https://x.com/thekingsmakers\n")
readme_text.config(state="disabled")

# Create GUI elements
Label(root, text="Export File:").grid(row=2, column=0, padx=5, pady=5)
export_entry = Entry(root, width=50)
export_entry.grid(row=2, column=1, padx=5, pady=5)
Button(root, text="Browse", command=lambda: browse_file(export_entry, "Export")).grid(row=2, column=2, padx=5, pady=5)

Label(root, text="Compare File:").grid(row=3, column=0, padx=5, pady=5)
compare_entry = Entry(root, width=50)
compare_entry.grid(row=3, column=1, padx=5, pady=5)
Button(root, text="Browse", command=lambda: browse_file(compare_entry, "Compare")).grid(row=3, column=2, padx=5, pady=5)

Button(root, text="Select Output Folder", command=lambda: compare_and_export(
    export_entry.get(), compare_entry.get(), select_output_folder(), log_text, readme_text)).grid(row=4, column=1, pady=10)

Label(root, text="Logs:").grid(row=5, column=0, padx=5, pady=5)
log_text = scrolledtext.ScrolledText(root, width=80, height=20, state="normal")
log_text.grid(row=6, column=0, columnspan=3, padx=5, pady=5)
log_text.grid_remove()  # Initially hide logs

Button(root, text="Show Logs", command=lambda: hide_readme(readme_text, log_text)).grid(row=7, column=0, pady=10)
Button(root, text="Refresh Logs", command=lambda: display_logs(log_text)).grid(row=7, column=1, pady=10)
Button(root, text="Show README", command=lambda: show_readme(readme_text, log_text)).grid(row=7, column=2, pady=10)

# Start the Tkinter main loop
root.mainloop()
