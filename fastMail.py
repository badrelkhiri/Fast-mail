# Import required modules
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import smtplib
from email.message import EmailMessage
import time
import random
import os
import xml.etree.ElementTree as ET
from PIL import Image
import html  # For escaping HTML characters
import threading
import queue  # For inter-thread communication

# ------------------------------------------------------------------------
# Constants and Configuration
# ------------------------------------------------------------------------

# Constants for XML settings file
SETTINGS_FILE = os.path.join(os.getenv('APPDATA'), "FastMails", "settings.xml")  # Path to your XML config file

# Ensure the directory exists
os.makedirs(os.path.dirname(SETTINGS_FILE), exist_ok=True)

# Default settings
DEFAULT_SETTINGS = {
  "sender_email": "",
  "app_password": "",
  "smtp_server": "smtp.gmail.com",
  "emails_per_batch": "70",
  "batch_delay_min": "181",
  "batch_delay_max": "230",
  "email_delay_min": "3",
  "email_delay_max": "5"
}

# List of randomized closing phrases
CLOSING_PHRASES = [
  "Cordialement,",
  "Sincèrement,",
  "Salutations,",
  "Merci,",
  "Chaleureusement,",
  "Bien à vous,",
  "Tous mes vœux,",
  "Amitiés,",
  "Respectueusement,"
]

# ------------------------------------------------------------------------
# Functions for Loading and Saving Settings
# ------------------------------------------------------------------------

def load_settings():
  """
  Load settings from the XML settings file.
  If the file doesn't exist or is corrupted, default settings are loaded.
  """
  settings = DEFAULT_SETTINGS.copy()
  if os.path.exists(SETTINGS_FILE):
      try:
          tree = ET.parse(SETTINGS_FILE)
          root = tree.getroot()
          for key in settings.keys():
              element = root.find(key)
              if element is not None:
                  settings[key] = element.text
      except ET.ParseError:
          messagebox.showerror("Error", "Settings file is corrupted. Loading default settings.")
  return settings

def save_settings(settings):
  """
  Save current settings to the XML settings file.
  """
  root = ET.Element("settings")
  for key, value in settings.items():
      elem = ET.SubElement(root, key)
      elem.text = value
  tree = ET.ElementTree(root)
  tree.write(SETTINGS_FILE)
  messagebox.showinfo("Success", "Settings saved successfully!")

def save_current_settings():
  """
  Collect current settings from the settings frame inputs,
  perform basic validation, save them, and update the global settings.
  """
  current_settings = {
      "sender_email": sender_email_entry.get(),
      "app_password": app_password_entry.get(),
      "smtp_server": smtp_server_entry.get(),
      "emails_per_batch": emails_per_batch_entry.get(),
      "batch_delay_min": batch_delay_min_entry.get(),
      "batch_delay_max": batch_delay_max_entry.get(),
      "email_delay_min": email_delay_min_entry.get(),
      "email_delay_max": email_delay_max_entry.get()
  }

  # Check if email and password are not empty
  if not current_settings["sender_email"].strip():
      messagebox.showerror("Error", "Sender email cannot be empty.")
      return

  if not current_settings["app_password"].strip():
      messagebox.showerror("Error", "App password cannot be empty.")
      return

  # Basic validation
  try:
      int(current_settings["emails_per_batch"])
      int(current_settings["batch_delay_min"])
      int(current_settings["batch_delay_max"])
      int(current_settings["email_delay_min"])
      int(current_settings["email_delay_max"])
  except ValueError:
      messagebox.showerror("Error", "Please enter valid integer values for delays and batch sizes.")
      return

  save_settings(current_settings)

  # Update the global settings
  global settings
  settings = current_settings

# Load settings at startup
settings = load_settings()

# ------------------------------------------------------------------------
# Main Application Setup
# ------------------------------------------------------------------------

# Initialize the main window
ctk.set_appearance_mode("System")  # Options: "System", "Light", "Dark"
app = ctk.CTk()
app.title("Fast mail")

# # Set the desired width and height of the form
# form_width = 800
# form_height = 700

# # Get screen dimensions
# screen_width = app.winfo_screenwidth()
# screen_height = app.winfo_screenheight()

# # Calculate x and y coordinates for the center position
# x = (screen_width // 2) - (form_width // 2)
# y = (screen_height // 2) - (form_height // 2)

# # Set the size and position of the form
# app.geometry(f"{form_width}x{form_height}+{x}+{y}")

app.resizable(False, False)  # False for width, False for height


# Set the window icon using an .ico file
app.iconbitmap("logo.ico")

# Configure grid layout
app.grid_rowconfigure(0, weight=1)
app.grid_columnconfigure(1, weight=1)

# ------------------------------------------------------------------------
# Left Menu Frame
# ------------------------------------------------------------------------

menu_frame = ctk.CTkFrame(app, width=200)
menu_frame.grid(row=0, column=0, sticky="ns")
menu_frame.grid_rowconfigure(2, weight=0)

# Load and create the logo image
logo_path = "logo.png"
if os.path.exists(logo_path):
  original_image = Image.open(logo_path).resize((170, 170))
  logo_image = ctk.CTkImage(original_image, size=(170, 170))
  logo_label = ctk.CTkLabel(menu_frame, image=logo_image, text="")
else:
  logo_label = ctk.CTkLabel(menu_frame, text="Logo Here", font=("Arial", 24))
logo_label.grid(row=0, column=0, padx=10, pady=(20, 10), sticky="n")

# Navigation Buttons
send_emails_button = ctk.CTkButton(menu_frame, text="Send Emails", command=lambda: show_frame(send_emails_frame))
send_emails_button.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

settings_button = ctk.CTkButton(menu_frame, text="Settings", command=lambda: show_frame(settings_frame))
settings_button.grid(row=2, column=0, padx=20, pady=10, sticky="ew")

# Function to switch between frames
def show_frame(frame):
  """
  Hide all frames and display the selected frame.
  """
  send_emails_frame.grid_remove()
  settings_frame.grid_remove()
  frame.grid()

# ------------------------------------------------------------------------
# Send Emails Frame
# ------------------------------------------------------------------------

send_emails_frame = ctk.CTkFrame(app)
send_emails_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
send_emails_frame.grid_rowconfigure(14, weight=1)
send_emails_frame.grid_columnconfigure(0, weight=1)
send_emails_frame.configure(fg_color="transparent")  # Set frame background to transparent

# Subject Frame
subject_frame = ctk.CTkFrame(send_emails_frame)
subject_frame.grid(row=0, column=0, sticky="w", pady=(0, 10), padx=(0, 10))
subject_frame.configure(fg_color="transparent")  # Set frame background to transparent

# Subject Label and Entry
subject_label = ctk.CTkLabel(subject_frame, text="Email Subject:")
subject_label.grid(row=0, column=0, sticky="w", pady=(0, 5), padx=(0,10))
subject_entry = ctk.CTkEntry(subject_frame, width=500)
subject_entry.grid(row=0, column=1, sticky="w", pady=(0, 10))

# Email Body Label
email_body_label = ctk.CTkLabel(send_emails_frame, text="Email Body:")
email_body_label.grid(row=2, column=0, sticky="w", pady=(0, 5))

# Formatting Buttons Frame
formatting_frame = ctk.CTkFrame(send_emails_frame)
formatting_frame.grid(row=3, column=0, sticky="w")
formatting_frame.configure(fg_color="transparent")  # Set frame background to transparent

# ------------------------------------------------------------------------
# Email Body Text Widget and Formatting Functions
# ------------------------------------------------------------------------

# Text widget for composing the email body
email_body_text = tk.Text(send_emails_frame, wrap="word", height=10)
email_body_text.grid(row=4, column=0, sticky="nsew", pady=(10, 10))

# Configure text widget tags for formatting
email_body_text.tag_configure("bold", font=("TkDefaultFont", 10, "bold"))
email_body_text.tag_configure("italic", font=("TkDefaultFont", 10, "italic"))
email_body_text.tag_configure("underline", font=("TkDefaultFont", 10, "underline"))
email_body_text.tag_configure("size_12", font=("TkDefaultFont", 12))
email_body_text.tag_configure("size_14", font=("TkDefaultFont", 14))
email_body_text.tag_configure("size_16", font=("TkDefaultFont", 16))

# Available fonts and tag mapping
AVAILABLE_FONTS = [
  "Sans_Serif", "Serif", "Fixed_Width", "Wide", "Narrow",
  "Comic_Sans_MS", "Garamond", "Georgia", "Tahoma", "Trebuchet_MS", "Verdana"
]

font_tag_mapping = {
  "Sans_Serif": ("Sans Serif", "TkDefaultFont"),
  "Serif": ("Serif", "Times New Roman"),
  "Fixed_Width": ("Fixed Width", "Courier New"),
  "Wide": ("Wide", "Arial"),
  "Narrow": ("Narrow", "Arial Narrow"),
  "Comic_Sans_MS": ("Comic Sans MS", "Comic Sans MS"),
  "Garamond": ("Garamond", "Garamond"),
  "Georgia": ("Georgia", "Georgia"),
  "Tahoma": ("Tahoma", "Tahoma"),
  "Trebuchet_MS": ("Trebuchet MS", "Trebuchet MS"),
  "Verdana": ("Verdana", "Verdana")
}

def apply_tag(tag_name):
  """
  Apply a formatting tag to the selected text in the email body text widget.
  Handles font and size tags as well as bold, italic, underline.
  """
  try:
      # Get selection range
      start = email_body_text.index("sel.first")
      end = email_body_text.index("sel.last")
      
      # Get all current tags at the selection
      current_tags = set()
      for index in range(len(email_body_text.tag_names())):
          tags_at_index = email_body_text.tag_names(f"{start}+{index}c")
          current_tags.update(tags_at_index)

      # Handle font tags separately
      font_tags = [tag for tag in current_tags if tag.startswith("font_")]
      size_tags = [tag for tag in current_tags if tag.startswith("size_")]
      format_tags = [tag for tag in current_tags if tag in ["bold", "italic", "underline"]]
      
      # If applying a font tag
      if tag_name.startswith("font_"):
          # Store current formatting
          current_size = None
          for size_tag in size_tags:
              current_size = size_tag
              
          # Remove old font tag
          for font_tag in font_tags:
              email_body_text.tag_remove(font_tag, start, end)
          
          # Apply new font tag
          email_body_text.tag_add(tag_name, start, end)
          
          # Reapply formatting tags
          for format_tag in format_tags:
              email_body_text.tag_add(format_tag, start, end)
          
          # Reapply size if it existed
          if current_size:
              email_body_text.tag_add(current_size, start, end)
              
      # If applying a size tag
      elif tag_name.startswith("size_"):
          # Remove old size tags
          for size_tag in size_tags:
              email_body_text.tag_remove(size_tag, start, end)
          
          # Apply new size tag
          email_body_text.tag_add(tag_name, start, end)
              
      # If applying formatting tag (bold, italic, underline)
      else:
          # Toggle the formatting tag
          if tag_name in current_tags:
              email_body_text.tag_remove(tag_name, start, end)
          else:
              email_body_text.tag_add(tag_name, start, end)
              
  except tk.TclError:
      pass  # No text selected

def configure_font_tags():
  """
  Configure font tags for different fonts and styles.
  """
  base_size = 10
  for tag_key, (display_name, font_name) in font_tag_mapping.items():
      # Configure regular font
      email_body_text.tag_configure(f"font_{tag_key}", font=(font_name, base_size))

      # Configure combinations with bold
      email_body_text.tag_configure(f"font_{tag_key}+bold", 
                                  font=(font_name, base_size, "bold"))

      # Configure combinations with italic
      email_body_text.tag_configure(f"font_{tag_key}+italic", 
                                  font=(font_name, base_size, "italic"))

      # Configure combinations with both bold and italic
      email_body_text.tag_configure(f"font_{tag_key}+bold+italic", 
                                  font=(font_name, base_size, "bold italic"))

# Configure font tags after creating the text widget
configure_font_tags()

# ------------------------------------------------------------------------
# Formatting Buttons
# ------------------------------------------------------------------------

# Bold Button
bold_button = ctk.CTkButton(formatting_frame, text="B", width=30, command=lambda: apply_tag("bold"))
bold_button.pack(side="left", padx=5)

# Italic Button
italic_button = ctk.CTkButton(formatting_frame, text="I", width=30, command=lambda: apply_tag("italic"))
italic_button.pack(side="left", padx=5)

# Underline Button
underline_button = ctk.CTkButton(formatting_frame, text="U", width=30, command=lambda: apply_tag("underline"))
underline_button.pack(side="left", padx=5)

# Size Buttons
size_frame = ctk.CTkFrame(formatting_frame)
size_frame.pack(side="left", padx=5)

size_12_button = ctk.CTkButton(size_frame, text="12", width=30, command=lambda: apply_tag("size_12"))
size_12_button.grid(row=0, column=0, padx=2)
size_14_button = ctk.CTkButton(size_frame, text="14", width=30, command=lambda: apply_tag("size_14"))
size_14_button.grid(row=0, column=1, padx=2)
size_16_button = ctk.CTkButton(size_frame, text="16", width=30, command=lambda: apply_tag("size_16"))
size_16_button.grid(row=0, column=2, padx=2)

# Font Selection
font_frame = ctk.CTkFrame(formatting_frame)
font_frame.pack(side="left", padx=5)
font_frame.configure(fg_color="transparent")  # Set frame background to transparent

# Create a StringVar to hold the selected font
selected_font = tk.StringVar(value=AVAILABLE_FONTS[0])

# Font Combobox
font_combobox = ctk.CTkComboBox(font_frame, values=AVAILABLE_FONTS, variable=selected_font)
font_combobox.pack(side="left")

def apply_font():
  """
  Apply the selected font to the selected text in the email body.
  """
  font_key = selected_font.get()
  tag_name = f"font_{font_key}"
  apply_tag(tag_name)

font_button = ctk.CTkButton(font_frame, text="Apply Font", command=apply_font)
font_button.pack(side="left", padx=5)

# ------------------------------------------------------------------------
# Recipient Emails Section
# ------------------------------------------------------------------------

# Recipient Emails Label
recipient_emails_label = ctk.CTkLabel(send_emails_frame, text="Recipient Emails (comma-separated):")
recipient_emails_label.grid(row=6, column=0, sticky="w", pady=(10, 5))

# Recipient Emails Text Widget
recipient_emails_text = tk.Text(send_emails_frame, height=5)
recipient_emails_text.grid(row=7, column=0, sticky="nsew", pady=(0, 10))

# Load Emails Button
load_emails_button = ctk.CTkButton(send_emails_frame, text="Load Emails from File", command=lambda: load_emails(recipient_emails_text))
load_emails_button.grid(row=8, column=0, sticky="w", pady=(0, 10))

# Function to load emails from file
def load_emails(text_widget):
  """
  Open a file dialog to select a file containing emails, and load the emails into the text widget.
  """
  file_path = filedialog.askopenfilename(title="Select Email List File", filetypes=[("Text Files", "*.txt")])
  if file_path:
      try:
          with open(file_path, 'r') as file:
              emails = file.read()
              text_widget.delete("1.0", "end")
              text_widget.insert("1.0", emails)
      except Exception as e:
          messagebox.showerror("Error", f"Failed to load emails: {e}")

def process_email_list(email_string):
  """
  Process the email string by splitting into a list, handling commas and newlines.
  Returns a list of email addresses.
  """
  emails = [email.strip() for email in email_string.replace('\n', ',').split(',') if email.strip()]
  return emails

# ------------------------------------------------------------------------
# PDF Attachment Section
# ------------------------------------------------------------------------

# PDF Attachment Frame
pdf_frame = ctk.CTkFrame(send_emails_frame)
pdf_frame.grid(row=9, column=0, sticky="w", pady=(0, 10), padx=(0, 10))
pdf_frame.configure(fg_color="transparent")  # Set frame background to transparent

# PDF File Path Variable
pdf_file_path = tk.StringVar()

def choose_pdf_file(label):
  """
  Open a file dialog to choose a PDF file and update the label.
  """
  file_path = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])
  if file_path:
      pdf_file_path.set(file_path)
      label.configure(text=os.path.basename(file_path))

# Button to Attach PDF File
choose_file_button = ctk.CTkButton(pdf_frame, text="Attach PDF File", command=lambda: choose_pdf_file(pdf_file_label))
choose_file_button.grid(row=0, column=0, sticky="w", pady=(0,10), padx=(0,5))

# Label to Display Selected PDF File
pdf_file_label = ctk.CTkLabel(pdf_frame, text="No PDF file selected")
pdf_file_label.grid(row=0, column=1, sticky="w", pady=(0,5))

# ------------------------------------------------------------------------
# Progress and Control Buttons
# ------------------------------------------------------------------------

# Progress Labels
emails_sent_label = ctk.CTkLabel(send_emails_frame, text="Emails Sent: 0 / 0")
emails_sent_label.grid(row=10, column=0, sticky="w", pady=(0, 0))

batch_delay_label = ctk.CTkLabel(send_emails_frame, text="")
batch_delay_label.grid(row=11, column=0, sticky="w", pady=(5, 5))

# Send Emails Button
send_button = ctk.CTkButton(send_emails_frame, text="Send Emails", command=lambda: send_emails())
send_button.grid(row=12, column=0, sticky="nsew", pady=(10,0))


# ------------------------------------------------------------------------
# Functions for Sending Emails
# ------------------------------------------------------------------------

def send_email(subject, body, to_email, pdf_file, smtp_session, sender_email):
  """
  Send a single email using the provided SMTP session.
  Optionally attach a PDF file.
  """
  msg = EmailMessage()
  msg.set_content(body, subtype='html')
  msg['Subject'] = subject
  msg['To'] = to_email
  msg['From'] = sender_email

  # Attach PDF if it exists
  if pdf_file and os.path.exists(pdf_file):
      with open(pdf_file, 'rb') as f:
          pdf_data = f.read()
          msg.add_attachment(pdf_data,
                             maintype='application',
                             subtype='pdf',
                             filename=os.path.basename(pdf_file))

  smtp_session.send_message(msg)

# Queue for inter-thread communication
progress_queue = queue.Queue()

def send_bulk_emails(sender_email, app_password, smtp_server, subject, email_body, pdf_file_path, to_list, settings):
  """
  Send bulk emails to the list of recipients.
  Handles batch delays and email delays.
  Updates progress via a queue for inter-thread communication.
  """
  EMAILS_PER_BATCH = int(settings["emails_per_batch"])
  BATCH_DELAY_MIN = int(settings["batch_delay_min"])
  BATCH_DELAY_MAX = int(settings["batch_delay_max"])
  EMAIL_DELAY_MIN = int(settings["email_delay_min"])
  EMAIL_DELAY_MAX = int(settings["email_delay_max"])

  smtp_session = smtplib.SMTP(smtp_server, 587)
  smtp_session.starttls()
  smtp_session.login(sender_email, app_password)

  # Send total emails count to the queue
  progress_queue.put({'total_emails': len(to_list)})

  for i, email in enumerate(to_list):
      # Handle batch delays
      if i > 0 and i % EMAILS_PER_BATCH == 0:
          smtp_session.quit()
          batch_delay = random.randint(BATCH_DELAY_MIN, BATCH_DELAY_MAX)
          for remaining in range(batch_delay, 0, -1):
              progress_queue.put({'batch_delay': remaining})
              time.sleep(1)
          smtp_session = smtplib.SMTP(smtp_server, 587)
          smtp_session.starttls()
          smtp_session.login(sender_email, app_password)
          # Clear the batch delay label after the delay is over
          progress_queue.put({'batch_delay': 0})

      # Select a random closing phrase
      closing_phrase = random.choice(CLOSING_PHRASES)
      # Create a unique identifier (e.g., email index)
      unique_id = f" [{i+1}]"

      # Append the closing phrase and unique ID to the email body
      unique_email_body = f"{email_body}<br><br>{closing_phrase}<br>{sender_email}{unique_id}"

      # Send the email
      send_email(subject, unique_email_body, email, pdf_file_path, smtp_session, sender_email)
      # Update progress
      progress_queue.put({'emails_sent': i + 1})

      if i < len(to_list) - 1:
          email_delay = random.randint(EMAIL_DELAY_MIN, EMAIL_DELAY_MAX)
          time.sleep(email_delay)

  smtp_session.quit()
  # Indicate that sending is done
  progress_queue.put({'status': 'done'})

def monitor_progress():
  """
  Monitor the progress of email sending using the progress_queue,
  and update the GUI accordingly.
  """
  try:
      message = progress_queue.get_nowait()
      if 'total_emails' in message:
          total_emails = message['total_emails']
          emails_sent_label.configure(text=f"Emails Sent: 0 / {total_emails}")

      if 'emails_sent' in message:
          emails_sent = message['emails_sent']
          total_text = emails_sent_label.cget("text")
          total_emails = total_text.split('/')[-1].strip()
          emails_sent_label.configure(text=f"Emails Sent: {emails_sent} / {total_emails}")

      if 'batch_delay' in message:
          remaining = message['batch_delay']
          if remaining > 0:
              batch_delay_label.configure(text=f"Waiting for {remaining} seconds")
          else:
              batch_delay_label.configure(text="")
      else:
          batch_delay_label.configure(text="")

      if 'status' in message and message['status'] == 'done':
          batch_delay_label.configure(text="Emails sent successfully!")
          messagebox.showinfo("Success", "Emails sent successfully!")
          return  # Stop monitoring
  except queue.Empty:
      pass
  app.after(100, monitor_progress)

def send_emails():
  """
  Gather input data from the GUI and start the email sending process in a new thread.
  """
  sender_email = sender_email_entry.get()
  app_password = app_password_entry.get()
  smtp_server = smtp_server_entry.get()
  subject = subject_entry.get()
  # Get formatted email body from the text widget
  email_body = get_formatted_email_body()
  recipients = recipient_emails_text.get("1.0", "end")
  pdf_path = pdf_file_path.get()

  # Process recipient emails
  to_list = process_email_list(recipients)

  if not all([sender_email, app_password, smtp_server, subject, email_body, pdf_path, to_list]):
      messagebox.showerror("Error", "Please fill in all fields and attach a PDF file.")
      return

  # Confirm before sending
  if not messagebox.askyesno("Confirmation", f"Are you sure you want to send emails to {len(to_list)} recipients?"):
      return

  try:
      # Start the email sending in a separate thread
      threading.Thread(target=send_bulk_emails, args=(sender_email, app_password, smtp_server, subject, email_body, pdf_path, to_list, settings), daemon=True).start()
      # Start the progress monitoring
      monitor_progress()
  except Exception as e:
      messagebox.showerror("Error", f"An error occurred: {e}")

# ------------------------------------------------------------------------
# Functions to Convert Text with Tags to HTML
# ------------------------------------------------------------------------

def get_formatted_email_body():
  """
  Converts the content of the Text widget, along with its tags, into an HTML-formatted string.
  Handles multiple overlapping tags (e.g., bold, italic, underline) correctly.
  """
  # Initialize variables
  html_output = ""
  index = "1.0"
  prev_tags = []

  while True:
      # Get the character at the current index
      char = email_body_text.get(index)

      # Break if we've reached the end
      if char == "":
          break

      # Get current tags at this index
      current_tags = email_body_text.tag_names(index)

      # Determine tags to close
      closing_tags = [tag for tag in prev_tags if tag not in current_tags]
      # Determine tags to open
      opening_tags = [tag for tag in current_tags if tag not in prev_tags]

      # Close tags in reverse order to maintain proper nesting
      for tag in reversed(closing_tags):
          html_output += get_html_closing_tag(tag)

      # Open new tags
      for tag in opening_tags:
          html_output += get_html_opening_tag(tag)

      # Add the current character, escaped for HTML
      if char == "\n":
          html_output += "<br>"
      else:
          html_output += escape_html(char)

      # Update previous tags
      prev_tags = current_tags

      # Move to next character
      index = email_body_text.index(f"{index}+1c")

  # Close any remaining open tags
  for tag in reversed(prev_tags):
      html_output += get_html_closing_tag(tag)

  return html_output

def get_html_opening_tag(tag):
  """
  Returns the appropriate HTML opening tag based on the tkinter tag.
  """
  if tag == "bold":
      return "<b>"
  elif tag == "italic":
      return "<i>"
  elif tag == "underline":
      return "<u>"
  elif tag.startswith("size_"):
      size = tag.split("_")[1]
      return f'<span style="font-size:{size}px">'
  elif tag.startswith("font_"):
      font_key = tag.split("_", 1)[1]
      # Map back the tag to actual font names with spaces
      font_display_mapping = {
          "Sans_Serif": "Sans Serif",
          "Serif": "Serif",
          "Fixed_Width": "Fixed Width",
          "Wide": "Wide",
          "Narrow": "Narrow",
          "Comic_Sans_MS": "Comic Sans MS",
          "Garamond": "Garamond",
          "Georgia": "Georgia",
          "Tahoma": "Tahoma",
          "Trebuchet_MS": "Trebuchet MS",
          "Verdana": "Verdana"
      }
      font_name = font_display_mapping.get(font_key, font_key)
      return f'<span style="font-family:\'{font_name}\';">'
  elif tag in ["left", "center", "right"]:
      return f'<div style="text-align:{tag};">'
  else:
      return ""

def get_html_closing_tag(tag):
  """
  Returns the appropriate HTML closing tag based on the tkinter tag.
  """
  if tag == "bold":
      return "</b>"
  elif tag == "italic":
      return "</i>"
  elif tag == "underline":
      return "</u>"
  elif tag.startswith("size_") or tag.startswith("font_"):
      return "</span>"
  elif tag in ["left", "center", "right"]:
      return "</div>"
  else:
      return ""

def escape_html(text):
  """
  Escape HTML special characters in text.
  """
  return html.escape(text)

# ------------------------------------------------------------------------
# Settings Frame
# ------------------------------------------------------------------------

settings_frame = ctk.CTkFrame(app)
settings_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
settings_frame.grid_columnconfigure(1, weight=1)
settings_frame.configure(fg_color="transparent")  # Set frame background to transparent

# Sender Email
sender_email_label = ctk.CTkLabel(settings_frame, text="Sender Email:")
sender_email_label.grid(row=0, column=0, sticky="w", pady=(0,5))
sender_email_entry = ctk.CTkEntry(settings_frame, width=400, placeholder_text="someone@example.com")
sender_email_entry.grid(row=0, column=1, sticky="w", pady=(0,5))
sender_email_entry.insert(0, settings.get("sender_email", ""))

# App Password
app_password_label = ctk.CTkLabel(settings_frame, text="App Password:")
app_password_label.grid(row=1, column=0, sticky="w", pady=(0,5))
app_password_entry = ctk.CTkEntry(settings_frame, width=400, show="*", placeholder_text="Enter your application password")
app_password_entry.grid(row=1, column=1, sticky="w", pady=(0,5))
app_password_entry.insert(0, settings.get("app_password", ""))

# SMTP Server
smtp_server_label = ctk.CTkLabel(settings_frame, text="SMTP Server:")
smtp_server_label.grid(row=2, column=0, sticky="w", pady=(0,5))
smtp_server_entry = ctk.CTkEntry(settings_frame, width=400)
smtp_server_entry.grid(row=2, column=1, sticky="w", pady=(0,5))
smtp_server_entry.insert(0, settings.get("smtp_server", "smtp.gmail.com"))

# Emails per Batch
emails_per_batch_label = ctk.CTkLabel(settings_frame, text="Emails Per Batch:")
emails_per_batch_label.grid(row=3, column=0, sticky="w", pady=(0,5))
emails_per_batch_entry = ctk.CTkEntry(settings_frame, width=400)
emails_per_batch_entry.grid(row=3, column=1, sticky="w", pady=(0,5))
emails_per_batch_entry.insert(0, settings.get("emails_per_batch", "70"))

# Batch Delay Min
batch_delay_min_label = ctk.CTkLabel(settings_frame, text="Batch Delay Min (seconds):")
batch_delay_min_label.grid(row=4, column=0, sticky="w", pady=(0,5))
batch_delay_min_entry = ctk.CTkEntry(settings_frame, width=400)
batch_delay_min_entry.grid(row=4, column=1, sticky="w", pady=(0,5))
batch_delay_min_entry.insert(0, settings.get("batch_delay_min", "181"))

# Batch Delay Max
batch_delay_max_label = ctk.CTkLabel(settings_frame, text="Batch Delay Max (seconds):")
batch_delay_max_label.grid(row=5, column=0, sticky="w", pady=(0,5))
batch_delay_max_entry = ctk.CTkEntry(settings_frame, width=400)
batch_delay_max_entry.grid(row=5, column=1, sticky="w", pady=(0,5))
batch_delay_max_entry.insert(0, settings.get("batch_delay_max", "230"))

# Email Delay Min
email_delay_min_label = ctk.CTkLabel(settings_frame, text="Email Delay Min (seconds):")
email_delay_min_label.grid(row=6, column=0, sticky="w", pady=(0,5))
email_delay_min_entry = ctk.CTkEntry(settings_frame, width=400)
email_delay_min_entry.grid(row=6, column=1, sticky="w", pady=(0,5))
email_delay_min_entry.insert(0, settings.get("email_delay_min", "3"))

# Email Delay Max
email_delay_max_label = ctk.CTkLabel(settings_frame, text="Email Delay Max (seconds):")
email_delay_max_label.grid(row=7, column=0, sticky="w", pady=(0,5))
email_delay_max_entry = ctk.CTkEntry(settings_frame, width=400)
email_delay_max_entry.grid(row=7, column=1, sticky="w", pady=(0,5))
email_delay_max_entry.insert(0, settings.get("email_delay_max", "5"))

# Save Settings Button
save_settings_button = ctk.CTkButton(settings_frame, text="Save Settings", command=lambda: save_current_settings())
save_settings_button.grid(row=8, column=1, sticky="e", pady=(10,0))

# ------------------------------------------------------------------------
# Start the Application
# ------------------------------------------------------------------------

# Initially show the Send Emails frame
show_frame(send_emails_frame)

# Run the main application loop
app.mainloop()