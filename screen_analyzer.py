import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
import os
import io
import google.generativeai as genai
from PIL import Image, ImageGrab
import speech_recognition as sr
import sqlite3
from datetime import datetime
import threading
import time
import pyttsx3
import queue
import pyautogui
import json
import re
import cv2
import numpy as np
from tkinter import messagebox, Label, Entry, Button, Text, Scrollbar, Frame
try:
    import pygetwindow as gw
except ImportError:
    print("Warning: pygetwindow not installed. App context detection will be limited.")
    gw = None
DATABASE_FILE = "conversation_log.db"
HARDCODED_API_KEY = "AIzaSyBBalgg7yFEO90hxDhQd1El30HTuMMzg"
LISTENING_KEYWORD = "anish"
tts_engine = pyttsx3.init()
request_queue = queue.Queue()
def configure_genai():
    """Configures the Gemini client using a hardcoded API key."""
    api_key = "AIzaSyBBalgg7yFEO90hxDhQd1El30HTL3Mzg"
    if not api_key or api_key == "YOUR_ACTUAL_GOOGLE_API_KEY":
        print("ERROR: Google API Key is not set in the script.")
        print("Please replace 'YOUR_ACTUAL_GOOGLE_API_KEY' in the script with your actual key.")
        return False
    try:
        genai.configure(api_key=api_key)
        print("Google AI SDK Configured using hardcoded key.")
        return True
    except Exception as e:
        print(f"ERROR: Failed to configure Google AI SDK: {e}")
        return False
def setup_database():
    """Creates the conversation log database and table if they don't exist."""
    log_manager._initialize_db()
    print(f"Database '{DATABASE_FILE}' setup complete.")
class ConversationLogManager:
    """Manages conversation and command logs in the database."""
    def __init__(self, db_file=DATABASE_FILE):
        self.db_file = db_file
        self._initialize_db()
    def _initialize_db(self):
        """Ensures database exists with all required tables."""
        conn = None
        try:
            conn = sqlite3.connect('conversation_log.db')
            cursor = conn.cursor()
            
            # First check if the database has any tables
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = [table[0] for table in cursor.fetchall()]
            
            # Create conversations table if it doesn't exist
            if 'conversations' not in tables:
                print("Creating conversations table...")
                cursor.execute("""
                    CREATE TABLE conversations (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                        user_query TEXT,
                        ai_response TEXT,
                        has_screenshot INTEGER DEFAULT 0,
                        command_type TEXT,
                        target TEXT,
                        success INTEGER DEFAULT 1,
                        details TEXT
                    )
                """)
            else:
                # If table exists, check for schema updates
                cursor.execute("PRAGMA table_info(conversations)")
                columns = [column[1] for column in cursor.fetchall()]
                if "command_type" not in columns:
                    print("Detected outdated database schema. Updating...")
                    cursor.execute("ALTER TABLE conversations RENAME TO conversations_old")
                    cursor.execute("""
                        CREATE TABLE conversations (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                            user_query TEXT,
                            ai_response TEXT,
                            has_screenshot INTEGER DEFAULT 0,
                            command_type TEXT,
                            target TEXT,
                            success INTEGER DEFAULT 1,
                            details TEXT
                        )
                    """)
                    cursor.execute("""
                        INSERT INTO conversations (id, timestamp, user_query, ai_response, has_screenshot)
                        SELECT id, timestamp, user_query, ai_response, has_screenshot FROM conversations_old
                    """)
                    cursor.execute("DROP TABLE conversations_old")
                    print("Database schema successfully updated.")
            
            # Create coordinates table if it doesn't exist
            if 'coordinates' not in tables:
                print("Creating coordinates table...")
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS coordinates (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        app_name TEXT NOT NULL DEFAULT 'unknown_app',
                        element_description TEXT NOT NULL,
                        x INTEGER NOT NULL,
                        y INTEGER NOT NULL,
                        timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                        success_count INTEGER DEFAULT 1,
                        fail_count INTEGER DEFAULT 0,
                        UNIQUE(app_name, element_description) ON CONFLICT REPLACE
                    )
                """)
            
            conn.commit()
            print(f"Database initialization complete: {self.db_file}")
        except sqlite3.Error as e:
            print(f"Database Error during initialization: {e}")
            if conn:
                try:
                    conn.rollback()
                except:
                    pass
        finally:
            if conn:
                conn.close()
    def log_conversation(self, user_query, ai_response, has_screenshot=False):
        """Logs a user query and AI response to the database."""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO conversations (user_query, ai_response, has_screenshot)
                VALUES (?, ?, ?)
            """, (user_query, ai_response, 1 if has_screenshot else 0))
            conn.commit()
            conn.close()
            return True
        except sqlite3.Error as e:
            print(f"Database Error: Failed to log conversation - {e}")
            return False
    def log_command(self, command_type, command_text, target=None, success=True, details=None):
        """Logs a command execution to the database using the conversations table.
        Args:
            command_type: String type of command (click, type, shortcut, etc.)
            command_text: The full text of the command
            target: The specific target (button name, text to type, etc.)
            success: Whether the command executed successfully
            details: Any additional details (coordinates, error messages, etc.)
        """
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO conversations (user_query, command_type, target, success, details)
                VALUES (?, ?, ?, ?, ?)
            """, (command_text, command_type, target, 1 if success else 0, details))
            conn.commit()
            conn.close()
            print(f"Command logged: {command_type} - {target}")
            return True
        except sqlite3.Error as e:
            print(f"Database Error: Failed to log command - {e}")
            return False
    def get_recent_commands(self, limit=10):
        """Retrieves recent commands from the database."""
        try:
            conn = sqlite3.connect(self.db_file)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("""
                SELECT * FROM conversations
                WHERE command_type IS NOT NULL
                ORDER BY timestamp DESC
                LIMIT ?
            """, (limit,))
            results = [dict(row) for row in cursor.fetchall()]
            conn.close()
            return results
        except sqlite3.Error as e:
            print(f"Database Error: Failed to retrieve commands - {e}")
            return []
    def get_command_stats(self):
        """Gets statistics about command usage."""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            cursor.execute("""
                SELECT command_type, COUNT(*) as count,
                       SUM(success) as success_count,
                       COUNT(*) - SUM(success) as fail_count
                FROM conversations
                WHERE command_type IS NOT NULL
                GROUP BY command_type
                ORDER BY count DESC
            """)
            type_stats = cursor.fetchall()
            cursor.execute("""
                SELECT COUNT(*) as total,
                       SUM(success) as total_success,
                       COUNT(*) - SUM(success) as total_fail
                FROM conversations
                WHERE command_type IS NOT NULL
            """)
            totals = cursor.fetchone()
            conn.close()
            stats = {
                "totals": {
                    "total": totals[0] if totals else 0,
                    "success": totals[1] if totals else 0,
                    "fail": totals[2] if totals else 0
                },
                "by_type": {row[0]: {"count": row[1], "success": row[2], "fail": row[3]} for row in type_stats}
            }
            return stats
        except sqlite3.Error as e:
            print(f"Database Error: Failed to get command stats - {e}")
            return {"totals": {"total": 0, "success": 0, "fail": 0}, "by_type": {}}
    def find_similar_commands(self, search_text, limit=5):
        """Finds similar commands to the provided text."""
        try:
            conn = sqlite3.connect(self.db_file)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("""
                SELECT * FROM conversations
                WHERE command_type IS NOT NULL
                  AND (user_query LIKE ? OR target LIKE ?)
                ORDER BY timestamp DESC
                LIMIT ?
            """, (f"%{search_text}%", f"%{search_text}%", limit))
            results = [dict(row) for row in cursor.fetchall()]
            conn.close()
            return results
        except sqlite3.Error as e:
            print(f"Database Error: Failed to find similar commands - {e}")
            return []
    def get_all_conversations(self, limit=100):
        """Get both conversations and commands."""
        try:
            conn = sqlite3.connect(self.db_file)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("""
                SELECT * FROM conversations
                ORDER BY timestamp DESC
                LIMIT ?
            """, (limit,))
            results = [dict(row) for row in cursor.fetchall()]
            conn.close()
            return results
        except sqlite3.Error as e:
            print(f"Database Error: Failed to retrieve conversations - {e}")
            return []
    def get_saved_coordinates(self, limit=None):
        """Get saved UI element coordinates. Use limit=None to get all entries."""
        try:
            conn = sqlite3.connect(self.db_file)
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            if limit is None:
                cursor.execute("""
                    SELECT * FROM coordinates
                    ORDER BY timestamp DESC
                """)
            else:
                cursor.execute("""
                    SELECT * FROM coordinates
                    ORDER BY timestamp DESC
                    LIMIT ?
                """, (limit,))
            results = [dict(row) for row in cursor.fetchall()]
            conn.close()
            return results
        except sqlite3.Error as e:
            print(f"Database Error: Failed to retrieve coordinates - {e}")
            return []
    def update_coordinates(self, coord_id, element_description=None, x=None, y=None, app_name=None):
        """Update saved coordinates by ID."""
        try:
            update_parts = []
            params = []
            if element_description is not None:
                update_parts.append("element_description = ?")
                params.append(element_description)
            if x is not None:
                update_parts.append("x = ?")
                params.append(int(x))
            if y is not None:
                update_parts.append("y = ?")
                params.append(int(y))
            if app_name is not None:
                update_parts.append("app_name = ?")
                params.append(app_name)
            update_parts.append("timestamp = CURRENT_TIMESTAMP")
            if not update_parts:
                print("No parameters provided for update")
                return False
            query = f"UPDATE coordinates SET {', '.join(update_parts)} WHERE id = ?"
            params.append(coord_id)
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            cursor.execute(query, params)
            conn.commit()
            success = cursor.rowcount > 0
            conn.close()
            if success:
                print(f"Updated coordinates id={coord_id}")
            else:
                print(f"No coordinates found with id={coord_id}")
            return success
        except sqlite3.Error as e:
            print(f"Database Error: Failed to update coordinates - {e}")
            return False
    def delete_coordinates(self, coord_id):
        """Delete saved coordinates by ID."""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM coordinates WHERE id = ?", (coord_id,))
            conn.commit()
            success = cursor.rowcount > 0
            conn.close()
            if success:
                print(f"Deleted coordinates id={coord_id}")
            else:
                print(f"No coordinates found with id={coord_id}")
            return success
        except sqlite3.Error as e:
            print(f"Database Error: Failed to delete coordinates - {e}")
            return False
log_manager = ConversationLogManager()
def log_conversation(user_query, ai_response, has_screenshot):
    """Logs a user query and AI response to the database."""
    return log_manager.log_conversation(user_query, ai_response, has_screenshot)
def process_command_conversation(command_type, command_text, target=None, success=True, details=None):
    """Process and log a command to the conversation log."""
    log_manager.log_command(command_type, command_text, target, success, details)
    return True
def speak_text(text):
    """Uses pyttsx3 to speak the provided text aloud."""
    if not text:
        return
    try:
        print(f"Speaking: {text}")
        tts_engine.say(text)
        tts_engine.runAndWait()
    except Exception as e:
        print(f"Error during text-to-speech: {e}")
def capture_screen_to_pil():
    """Captures the primary screen and returns it as a PIL Image object."""
    try:
        time.sleep(0.2)
        screenshot = ImageGrab.grab()
        print(f"Screenshot captured (size: {screenshot.size})")
        return screenshot
    except Exception as e:
        print(f"Error capturing screen: {e}")
        return None
def get_ai_response(user_query, pil_image=None):
    """Sends the user query and optional image to Gemini for a response."""
    try:
        model = genai.GenerativeModel('gemini-2.0-flash')
        prompt_parts = []
        base_prompt = (
            "You are a friendly and helpful AI assistant designed for voice interaction. "
            "Analyze the user's voice query and the attached screenshot (if provided). "
            "Respond naturally and conversationally, as if you were talking to the user. "
            "Keep your answers concise and to the point (usually 1-2 sentences), focusing on being helpful. "
            "If asked a question, answer directly. If shown the screen, comment briefly on what you see or offer a relevant suggestion."
        )
        prompt_parts.append(base_prompt)
        if user_query:
            prompt_parts.append(f'\nUser Query: "{user_query}"')
        else:
            prompt_parts.append("\nUser did not provide a specific voice query, focus on the screen.")
        if pil_image:
            prompt_parts.append("\nAnalyze the attached screenshot in conjunction with the query:")
            prompt_parts.append(pil_image)
            has_screenshot = True
        else:
            prompt_parts.append("\n(No screenshot provided for this query)")
            has_screenshot = False
        print("Sending request to Gemini...")
        response = model.generate_content(prompt_parts)
        if response and response.text:
            ai_text = response.text.strip()
            print("Received response from Gemini.")
            log_conversation(user_query or "(Screen Analysis Only)", ai_text, has_screenshot)
            return ai_text
        else:
            print("Gemini returned an empty response or failed.")
            try:
                print(f"Gemini response details: {response.prompt_feedback}")
            except Exception:
                pass
            return "Error: Failed to get analysis from AI. Check API key and model access."
    except Exception as e:
        print(f"Error interacting with Gemini API: {e}")
        return f"Error during AI analysis: {e}"
def save_coordinates(element_description, x, y, success=True, app_name=None):
    """Save coordinates for an element to the database for future use."""
    try:
        if not app_name:
            app_info = get_active_app_info()
            app_name = app_info["app_name"]
        conn = sqlite3.connect(DATABASE_FILE)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, success_count, fail_count FROM coordinates 
            WHERE app_name = ? AND element_description = ?
        """, (app_name.lower(), element_description.lower()))
        existing = cursor.fetchone()
        if existing:
            entry_id, success_count, fail_count = existing
            if success:
                success_count += 1
            else:
                fail_count += 1
            cursor.execute("""
                UPDATE coordinates 
                SET x = ?, y = ?, timestamp = CURRENT_TIMESTAMP, 
                    success_count = ?, fail_count = ?
                WHERE id = ?
            """, (x, y, success_count, fail_count, entry_id))
            print(f"Updated coordinates for '{element_description}' in app '{app_name}': ({x}, {y})")
        else:
            cursor.execute("""
                INSERT INTO coordinates (app_name, element_description, x, y, 
                                        success_count, fail_count)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (app_name.lower(), element_description.lower(), x, y, 
                 1 if success else 0, 0 if success else 1))
            print(f"Saved new coordinates for '{element_description}' in app '{app_name}': ({x}, {y})")
        conn.commit()
        conn.close()
        return True
    except sqlite3.Error as e:
        print(f"Database Error: Failed to save coordinates - {e}")
        return False
def find_coordinates_in_db(element_description, app_name=None):
    """Look for previously saved coordinates for an element description using flexible matching."""
    try:
        if not app_name:
            app_info = get_active_app_info()
            app_name = app_info["app_name"]
        conn = sqlite3.connect(DATABASE_FILE)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT x, y FROM coordinates 
            WHERE app_name = ? AND LOWER(element_description) = ? 
            AND success_count > fail_count
            ORDER BY timestamp DESC
            LIMIT 1
        """, (app_name.lower(), element_description.lower()))
        result = cursor.fetchone()
        if not result:
            print(f"  DB Search: No exact match for '{element_description}' in '{app_name}'. Trying LIKE...")
            cursor.execute("""
                SELECT x, y FROM coordinates 
                WHERE app_name = ? AND LOWER(element_description) LIKE ?
                AND success_count > fail_count
                ORDER BY timestamp DESC
                LIMIT 1
            """, (app_name.lower(), f"%{element_description.lower()}%"))
            result = cursor.fetchone()
        if not result:
            print(f"  DB Search: No LIKE match for '{element_description}' in '{app_name}'. Trying reverse LIKE...")
            cursor.execute("""
                SELECT x, y FROM coordinates 
                WHERE app_name = ? AND ? LIKE '%' || LOWER(element_description) || '%'
                AND success_count > fail_count
                ORDER BY LENGTH(element_description) DESC, timestamp DESC
                LIMIT 1
            """, (app_name.lower(), element_description.lower()))
            result = cursor.fetchone()
        if not result:
            print(f"  DB Search: No match in '{app_name}'. Trying LIKE in any app...")
            cursor.execute("""
                SELECT x, y FROM coordinates 
                WHERE (LOWER(element_description) LIKE ? OR ? LIKE '%' || LOWER(element_description) || '%')
                AND success_count > fail_count
                ORDER BY 
                    CASE WHEN app_name = ? THEN 1 ELSE 2 END, -- Prioritize current app
                    LENGTH(element_description) DESC, -- Prioritize longer (more specific) matches
                    timestamp DESC
                LIMIT 1
            """, (f"%{element_description.lower()}%", element_description.lower(), app_name.lower())) # Corrected parameters
            result = cursor.fetchone()
        conn.close()
        if result:
            x, y = result
            print(f"Found saved coordinates ({x}, {y}) for '{element_description}' in app '{app_name}' (using flexible matching)")
            return (x, y)
        print(f"No coordinates found for '{element_description}' in app '{app_name}' (after flexible matching)")
        return None
    except sqlite3.Error as e:
        print(f"Database Error: Failed to find coordinates - {e}")
        return None
def get_coordinates_from_ai(element_description, pil_image, app_name=None):
    """Sends image and element description to AI specifically asking for coordinates."""
    if not app_name:
        app_info = get_active_app_info()
        app_name = app_info["app_name"]
    saved_coords = find_coordinates_in_db(element_description, app_name)
    if saved_coords:
        print(f"Using saved coordinates {saved_coords} for '{element_description}' in app '{app_name}'")
        return json.dumps({"x": saved_coords[0], "y": saved_coords[1]})
    if not pil_image:
        return '{"error": "Screenshot required for coordinate detection."}'
    try:
        model = genai.GenerativeModel('gemini-2.0-flash') 
        prompt = (
            "You are an expert UI element locator, simulating human visual identification. "
            "Your primary goal is to find the *center* coordinates (x, y) of a specific UI element based on a description and a screenshot.\n\n" 
            f'CONTEXT: The user wants to interact with the UI element described as: "{element_description}" within the application "{app_name}".'
            ' *Note: The user\'s description might be slightly misspelled, abbreviated, or incomplete (e.g., "save" instead of "save button"). Use visual context to find the BEST match.*\n\n'
            "INSTRUCTIONS:\n"
            "1. Carefully examine the screenshot, focusing on the area likely belonging to the application mentioned.\n"
            "2. Look for visual cues (text labels, icons, shapes, placement) that BEST match the description: "
            f"'{element_description}'. Prioritize the most likely intended target based on the description, even if imperfect.\n"
            "3. Consider common UI patterns (menus top, buttons bottom/right, close top-right) but prioritize the specific description.\n"
            "4. Calculate the exact pixel coordinates of the *center* of the *most likely* match.\n\n"
            "RESPONSE FORMAT:\n"
            "- Accuracy is critical. If confident, respond *only* with the JSON: {\"x\": <center_x_coordinate>, \"y\": <center_y_coordinate>}\n"
            "- If unsure or the element isn't clearly visible/identifiable, respond *only* with: {\"error\": \"Element not found or ambiguous\"}\n"
            "ABSOLUTELY NO other text, comments, or explanations."
        )
        print(f"Sending request to Gemini for coordinates of '{element_description}' in app '{app_name}' (AI fallback)...")
        response = model.generate_content([prompt, pil_image])
        if response and response.text:
            ai_text = response.text.strip()
            print(f"Received coordinate response from Gemini: {ai_text}")
            log_conversation(f"Click Request: {element_description} in {app_name}", ai_text, True)
            return ai_text
        else:
            print("Gemini coordinate request returned empty response.")
            return '{"error": "AI did not respond for coordinates"}'
    except Exception as e:
        print(f"Error interacting with Gemini API for coordinates: {e}")
        return f'{{"error": "API error during coordinate request: {e}"}}'
def click_at_coordinates(x, y):
    """Moves mouse to (x, y) and performs a left click."""
    try:
        target_x, target_y = int(x), int(y)
        print(f"Attempting to click at ({target_x}, {target_y})...")
        pyautogui.moveTo(target_x, target_y, duration=0.2)
        time.sleep(0.1)
        pyautogui.click(target_x, target_y)
        print(f"Clicked at ({target_x}, {target_y}).")
        return True
    except Exception as e:
        print(f"Error performing click at ({x}, {y}): {e}")
        return False
def type_text(text_to_type):
    """Uses pyautogui to type the given text."""
    try:
        print(f"Attempting to type: '{text_to_type}'")
        pyautogui.write(text_to_type, interval=0.05)
        print("Typing complete.")
        return True
    except Exception as e:
        print(f"Error performing typing: {e}")
        return False
def get_active_app_info():
    """Get information about the currently active application window."""
    result = {"app_name": "Unknown", "window_title": "Unknown"}
    try:
        if 'gw' in globals():
            active_window = gw.getActiveWindow()
            if active_window:
                result["window_title"] = active_window.title
                app_parts = result["window_title"].split(" - ")
                if len(app_parts) > 1:
                    result["app_name"] = app_parts[-1].strip()
                else:
                    result["app_name"] = result["window_title"]
    except Exception as e:
        print(f"Error getting active window info: {e}")
    if result["app_name"] == "Unknown" or not result["app_name"]:
        result["app_name"] = "Unknown"
    result["app_name"] = result["app_name"][:50]
    return result
class ChatScreenApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.waiting_for_calibration_confirmation = None
        self.listening_active = False
        self.listening_timeout = 0
        self.LISTENING_TIMEOUT_SECONDS = 30
        self.training_mode_active = False
        self.training_elements_db = {}
        if gw:
            try:
                current_window = gw.getActiveWindow()
                if current_window:
                    print(f"Active window detection working: {current_window.title}")
                else:
                    print("No active window detected, but pygetwindow is available")
            except Exception as e:
                print(f"Warning: Active window detection may not work correctly: {e}")
        else:
            print("Warning: pygetwindow is not installed. Run: pip install pygetwindow")
            print("Window-specific coordinate memory will be limited.")
        self.title("AI Screen & Voice Assistant")
        self.geometry("700x600")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=0)
        self.chat_display = ctk.CTkTextbox(self, wrap=tk.WORD, state="disabled", fg_color="transparent")
        self.chat_display.grid(row=0, column=0, columnspan=2, padx=20, pady=(20, 10), sticky="nsew")
        self.input_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.input_frame.grid(row=1, column=0, columnspan=2, padx=20, pady=(5, 5), sticky="ew")
        self.input_frame.grid_columnconfigure(0, weight=1)
        self.input_entry = ctk.CTkEntry(self.input_frame, placeholder_text="Type your message here...")
        self.input_entry.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        self.input_entry.bind("<Return>", self.handle_send_message)
        self.send_button = ctk.CTkButton(self.input_frame, text="Send", width=80, command=self.handle_send_message)
        self.send_button.grid(row=0, column=1, sticky="e")
        self.status_label = ctk.CTkLabel(self, text="Status: Initializing...", anchor="w")
        self.status_label.grid(row=2, column=0, padx=20, pady=(5, 10), sticky="ew")
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.grid(row=2, column=1, padx=20, pady=(5, 10), sticky="e")
        self.history_button = ctk.CTkButton(self.button_frame, text="Command History", width=120, 
                                          command=self.show_command_history)
        self.history_button.grid(row=0, column=0, padx=(0, 10), sticky="e")
        self.analyze_screen_button = ctk.CTkButton(self.button_frame, text="Analyze Screen Only", 
                                                command=self.handle_analyze_screen_only)
        self.analyze_screen_button.grid(row=0, column=1, sticky="e")
        self.update_status("Starting background listener...")
        listener_thread = threading.Thread(target=self.continuous_listener, args=(self.update_status_safe,), daemon=True)
        listener_thread.start()
        self.check_request_queue()
    def update_status_safe(self, message):
        """Safely update status label from any thread using after()."""
        self.after(0, self.update_status, message)
    def process_request_safe(self, command_tuple):
        """Safely triggers processing based on command from listener queue."""
        self.after(0, self.route_command, command_tuple)
    def update_status(self, message):
        self.status_label.configure(text=f"Status: {message}")
    def display_message(self, sender, message):
        self.chat_display.configure(state="normal")
        self.chat_display.insert(tk.END, f"{sender}: {message}\n\n")
        self.chat_display.configure(state="disabled")
        self.chat_display.see(tk.END)
    def run_in_thread(self, target_func, *args):
        """Runs a function in a separate thread to avoid blocking the GUI."""
        thread = threading.Thread(target=target_func, args=args, daemon=True)
        thread.start()
    def check_request_queue(self):
        """Periodically check the queue for commands from the listener thread."""
        try:
            command_tuple = request_queue.get_nowait()
            self.route_command(command_tuple)
        except queue.Empty:
            pass
        finally:
            self.after(100, self.check_request_queue)
    def route_command(self, command_tuple):
        """Routes the command from the queue to the appropriate handler."""
        command_verb, data = command_tuple
        print(f"[GUI] Routing command: {command_verb}, Data: {data}")
        process_command_conversation("route", f"{command_verb}: {data}", command_verb)
        if command_verb == "query":
            self.process_user_request(data)
        elif command_verb == "click":
            self.process_click_request(data)
        elif command_verb == "type":
            self.process_type_request(data)
        elif command_verb == "shortcut":
            self.process_shortcut_command(data)
        elif command_verb == "system":
            self.execute_system_command(data)
        elif command_verb == "action":
            self.execute_action(data)
        elif command_verb == "history":
            self.show_command_history()
        elif command_verb == "analyze":
            self.handle_analyze_screen_only()
        elif command_verb == "confirm_calibration_yes":
            self.process_start_calibration(data)
        elif command_verb == "confirm_calibration_no":
            self.display_message("System", "OK, calibration cancelled.")
            self.update_status(f"Listener: Ready (Say '{LISTENING_KEYWORD}' + command)")
        else:
            print(f"[GUI] Unknown command received: {command_verb}")
            self.display_message("System", f"Sorry, I don't know the command '{command_verb}'.")
            self.update_status(f"Listener: Ready (Say '{LISTENING_KEYWORD}' + command)")
    def process_user_request(self, user_query):
        """Handles screen capture, AI call, display, and TTS for a general query."""
        self._disable_inputs()
        if user_query:
            self.display_message("You", user_query)
        else:
            self.update_status("Keyword detected, analyzing screen context.")
        self.update_status_safe("Capturing screen...")
        screenshot = capture_screen_to_pil()
        if not screenshot:
            self.update_status_safe("Screen capture failed.")
            self.display_message("System", "Error: Could not capture screen.")
            self._enable_inputs()
            return
        self.update_status_safe("Sending to AI...")
        self.run_in_thread(self._query_ai_and_speak_thread, user_query, screenshot)
    def _query_ai_and_speak_thread(self, user_query, screenshot):
        """Runs general AI query and TTS in a thread."""
        ai_response = get_ai_response(user_query, screenshot)
        self.after(0, self.display_message, "AI", ai_response)
        speak_text(ai_response)
        if not self.waiting_for_calibration_confirmation:
            self.after(0, self.update_status, f"Listener: Ready (Say '{LISTENING_KEYWORD}' + command)")
        self.after(0, self._enable_inputs())
    def process_click_request(self, element_description):
        """Handles screen capture, coordinate finding, parsing, and clicking."""
        self._disable_inputs()
        self.display_message("Action", f"Attempting to click: {element_description}")
        self.update_status_safe("Capturing screen for click target...")
        process_command_conversation("click_request", f"Click request: {element_description}", element_description)
        screenshot = capture_screen_to_pil()
        if not screenshot:
            self.update_status_safe("Screen capture failed.")
            self.display_message("System", "Error: Could not capture screen for click.")
            process_command_conversation("click_request", f"Click request failed: {element_description}", 
                                     element_description, False, "Screen capture failed")
            self._enable_inputs()
            return
        self.update_status_safe(f"Asking AI for coordinates of '{element_description}'...")
        self.run_in_thread(self._find_and_click_thread, element_description, screenshot)
    def _find_and_click_thread(self, element_description, screenshot):
        """Gets coordinates, parses, and performs click in a thread."""
        app_info = get_active_app_info()
        app_name = app_info["app_name"]
        full_title = app_info["window_title"]
        self.after(0, self.update_status, f"Finding '{element_description}' in '{app_name}'...")
        coord_response_raw = get_coordinates_from_ai(element_description, screenshot, app_name)
        coords = None
        error_message = None
        try:
            coord_response = coord_response_raw.strip()
            if coord_response.startswith("```json"):
                coord_response = coord_response[7:].strip()
            elif coord_response.startswith("```"):
                coord_response = coord_response[3:].strip()
            if coord_response.endswith("```"):
                coord_response = coord_response[:-3].strip()
            data = json.loads(coord_response) 
            if "error" in data:
                error_message = f"AI could not find '{element_description}' in '{app_name}': {data['error']}"
            elif "x" in data and "y" in data:
                coords = (data['x'], data['y'])
            else:
                error_message = "AI response did not contain valid x, y coordinates."
        except json.JSONDecodeError:
            error_message = f"Failed to parse AI coordinate response: {coord_response[:100]}..."
        except Exception as e:
            error_message = f"Error processing coordinate response: {e}"
        if coords:
            self.after(0, self.update_status, f"Coordinates found in '{app_name}': {coords}. Clicking...")
            click_success = click_at_coordinates(coords[0], coords[1])
            if click_success:
                self.after(0, self.display_message, "System", f"Clicked at {coords} for '{element_description}' in '{app_name}'.")
                save_coordinates(element_description, coords[0], coords[1], True, app_name)
                process_command_conversation("click_execute", f"Clicked {element_description} in {app_name}", 
                                         element_description, True, f"Coords: {coords}")
                self.after(0, self.update_status, f"Listener: Ready (Say '{LISTENING_KEYWORD}' + command)")
                self.after(0, self._enable_inputs())
            else:
                self.after(0, self.display_message, "System", f"Click failed at {coords} for '{element_description}' in '{app_name}'.")
                process_command_conversation("click_execute", f"Click failed for {element_description} in {app_name}", 
                                         element_description, False, f"Coords: {coords}")
                self.after(0, self.update_status, "Click action failed.")
                self.after(0, self.update_status, f"Listener: Ready (Say '{LISTENING_KEYWORD}' + command)")
                self.after(0, self._enable_inputs())
        else:
            self.after(0, self.display_message, "Error", error_message or "Could not determine coordinates.")
            self.after(0, self.update_status, "Failed to find coordinates.")
            process_command_conversation("click_execute", f"Failed to find {element_description} in {app_name}", 
                                     element_description, False, error_message)
            speak_text(f"Sorry, I could not find '{element_description}' in {app_name}. Try being more specific.") 
            self.after(0, self.update_status, f"Listener: Ready (Say '{LISTENING_KEYWORD}' + command)")
            self.after(0, self._enable_inputs())
    def ask_for_calibration(self, app_name, element_description, failed_coords):
        """Asks the user if they want to manually calibrate the failed click."""
        self.waiting_for_calibration_confirmation = (app_name, element_description, failed_coords)
        calib_question = f"Click failed for '{element_description}' in '{app_name}'. Would you like to manually calibrate its position? (Say 'yes' or 'no')"
        self.display_message("System", calib_question)
        self.update_status("Waiting for calibration confirmation (yes/no)...")
        speak_text("Would you like to manually calibrate its position?")
        self.after(0, self._enable_inputs())
    def process_type_request(self, text_to_type):
        """Handles the request to type text."""
        self._disable_inputs()
        self.display_message("Action", f"Attempting to type: {text_to_type}")
        self.update_status_safe("Typing text...")
        process_command_conversation("type_request", f"Type request: {text_to_type}", text_to_type)
        self.run_in_thread(self._do_typing_thread, text_to_type)
    def _do_typing_thread(self, text_to_type):
        """Performs the typing action in a thread."""
        type_success = type_text(text_to_type)
        if type_success:
            self.after(0, self.display_message, "System", "Typing complete.")
            process_command_conversation("type_execute", f"Typed: {text_to_type}", text_to_type, True)
            self.after(0, self.update_status, f"Listener: Ready (Say '{LISTENING_KEYWORD}' + command)")
        else:
            self.after(0, self.display_message, "System", "Typing failed.")
            process_command_conversation("type_execute", f"Typing failed: {text_to_type}", text_to_type, False)
            self.after(0, self.update_status, "Typing action failed.")
        self.after(0, self._enable_inputs())
    def handle_analyze_screen_only(self):
        """Captures and analyzes the screen without a query."""
        self._disable_inputs()
        self.update_status("Starting screen analysis...")
        process_command_conversation("analyze_screen", "Screen analysis requested", "screen_only")
        self.run_in_thread(self._analyze_screen_only_thread)
    def _analyze_screen_only_thread(self):
        """Thread to handle screen capture and analysis."""
        self.update_status_safe("Capturing screen...")
        screenshot = capture_screen_to_pil()
        if not screenshot:
            self.update_status_safe("Screen capture failed.")
            self.display_message("System", "Error: Could not capture screen.")
            process_command_conversation("analyze_screen", "Screen analysis failed", "screen_only", False, "Screen capture failed")
            self._enable_inputs()
            return
        self.update_status_safe("Sending to AI...")
        self.run_in_thread(self._query_ai_and_speak_thread, None, screenshot)
    def process_shortcut_command(self, command_name):
        """Executes the keyboard shortcut associated with the command name."""
        mappings = self.get_shortcut_mappings()
        if command_name in mappings:
            keys_to_press = mappings[command_name]
            self.display_message("Action", f"Executing shortcut: {command_name.capitalize()} ({' + '.join(keys_to_press)})")
            self.update_status_safe(f"Executing shortcut: {command_name.capitalize()}")
            process_command_conversation("shortcut_request", f"Shortcut request: {command_name}", command_name)
            self.run_in_thread(self._do_shortcut_thread, keys_to_press, command_name)
        else:
            print(f"Error: Unknown shortcut command '{command_name}'")
            self.display_message("System", f"Error: Shortcut '{command_name}' not recognized.")
            process_command_conversation("shortcut_request", f"Unknown shortcut: {command_name}", 
                                      command_name, False, "Shortcut not recognized")
            self._enable_inputs()
    def execute_system_command(self, command_text):
        """Executes a system command using the command prompt."""
        self._disable_inputs()
        self.display_message("Action", f"Executing system command: {command_text}")
        self.update_status_safe("Running system command...")
        process_command_conversation("system_command", f"System command: {command_text}", command_text)
        self.run_in_thread(self._do_system_command_thread, command_text)
    def _do_system_command_thread(self, command_text):
        """Executes the system command in a thread."""
        try:
            import subprocess
            if "powershell -c" in command_text and "MinimizeAll" in command_text:
                self.after(0, self.display_message, "System", "Minimizing all windows...")
                result = subprocess.run(command_text, capture_output=True, text=True, shell=True)
                if result.returncode == 0:
                    self.after(0, self.display_message, "System", "Successfully minimized all windows")
                    process_command_conversation("system_command", "Minimize all windows", "desktop", True, "Success")
                else:
                    error = result.stderr.strip() or f"Command failed with exit code {result.returncode}"
                    self.after(0, self.display_message, "System", f"Failed to minimize windows: {error[:100]}")
                    process_command_conversation("system_command", "Minimize all windows", "desktop", False, error[:100])
                self.after(0, self._enable_inputs)
                return
            forbidden_keywords = ["format", "deltree", "rmdir /s", "rd /s", "rm -rf"]
            if any(keyword in command_text.lower() for keyword in forbidden_keywords):
                self.after(0, self.display_message, "System", "This command is not allowed for safety reasons.")
                process_command_conversation("system_command", f"Blocked command: {command_text}", 
                                         command_text, False, "Command blocked for safety")
                self.after(0, self._enable_inputs)
                return
            result = subprocess.run(f"cmd /c {command_text}", 
                                  capture_output=True, 
                                  text=True, 
                                  shell=True)
            if result.returncode == 0:
                output = result.stdout.strip() or "Command executed successfully (no output)."
                self.after(0, self.display_message, "System", f"Command succeeded:\n{output[:300]}")
                process_command_conversation("system_command", f"Executed: {command_text}", 
                                         command_text, True, output[:100])
            else:
                error = result.stderr.strip() or f"Command failed with exit code {result.returncode}"
                self.after(0, self.display_message, "System", f"Command failed:\n{error[:300]}")
                process_command_conversation("system_command", f"Failed: {command_text}", 
                                         command_text, False, error[:100])
        except Exception as e:
            self.after(0, self.display_message, "System", f"Error executing command: {e}")
            process_command_conversation("system_command", f"Error: {command_text}", 
                                     command_text, False, str(e))
        finally:
            self.after(0, self._enable_inputs)
    def execute_action(self, action_name):
        """Execute a direct system action like 'minimize all' or 'maximize window'."""
        self._disable_inputs()
        self.display_message("Action", f"Executing: {action_name}")
        process_command_conversation("direct_action", f"Action: {action_name}", action_name)
        action_lower = action_name.lower().strip()
        try:
            import subprocess
            if action_lower in ["minimize all", "minimise all", "show desktop"]:
                self.display_message("System", "Minimizing all windows...")
                result = subprocess.run("powershell -c \"(New-Object -ComObject Shell.Application).MinimizeAll()\"", 
                                     shell=True, capture_output=True, text=True)
                if result.returncode == 0:
                    self.display_message("System", "All windows minimized successfully")
                    process_command_conversation("direct_action", "Minimize all windows", "desktop", True, "Success")
                else:
                    self.display_message("System", "Using keyboard shortcut instead...")
                    pyautogui.hotkey('win', 'd')
                    self.display_message("System", "All windows minimized using keyboard shortcut")
                    process_command_conversation("direct_action", "Minimize all windows", "desktop", True, "Used shortcut")
            elif action_lower in ["maximize all", "maximise all", "restore all", "restore windows"]:
                self.display_message("System", "Restoring all windows...")
                result = subprocess.run("powershell -c \"(New-Object -ComObject Shell.Application).UndoMinimizeALL()\"", 
                                     shell=True, capture_output=True, text=True)
                if result.returncode == 0:
                    self.display_message("System", "All windows restored successfully")
                else:
                    pyautogui.hotkey('win', 'd')
                    self.display_message("System", "All windows restored using keyboard shortcut")
            elif action_lower in ["maximize window", "maximise window", "maximize current window"]:
                self.display_message("System", "Maximizing current window...")
                pyautogui.hotkey('alt', 'space')
                time.sleep(0.1)
                pyautogui.press('x')
                self.display_message("System", "Current window maximized")
            elif action_lower in ["lock computer", "lock screen", "lock pc", "lock windows"]:
                self.display_message("System", "Locking your computer...")
                subprocess.run("rundll32.exe user32.dll,LockWorkStation", shell=True)
                self.display_message("System", "Computer locked")
            elif action_lower in ["sleep", "sleep mode", "put computer to sleep"]:
                self.display_message("System", "Putting computer to sleep...")
                subprocess.run("rundll32.exe powrprof.dll,SetSuspendState 0,1,0", shell=True)
                self.display_message("System", "Sleep command sent")
            elif action_lower in ["take screenshot", "screenshot", "capture screen"]:
                self.display_message("System", "Taking screenshot...")
                os.makedirs("screenshots", exist_ok=True)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"screenshots/screenshot_{timestamp}.png"
                screenshot = pyautogui.screenshot()
                screenshot.save(filename)
                self.display_message("System", f"Screenshot saved to {filename}")
            elif action_lower in ["open settings", "windows settings"]:
                self.display_message("System", "Opening Windows Settings...")
                subprocess.Popen("start ms-settings:", shell=True)
                self.display_message("System", "Windows Settings opened")
            elif action_lower in ["open file explorer", "explorer", "files"]:
                self.display_message("System", "Opening File Explorer...")
                subprocess.Popen("explorer", shell=True)
                self.display_message("System", "File Explorer opened")
            elif action_lower in ["task manager", "open task manager"]:
                self.display_message("System", "Opening Task Manager...")
                pyautogui.hotkey('ctrl', 'shift', 'esc')
                self.display_message("System", "Task Manager opened")
            elif action_lower in ["run", "open run dialog"]:
                self.display_message("System", "Opening Run dialog...")
                pyautogui.hotkey('win', 'r')
                self.display_message("System", "Run dialog opened")
            elif "volume" in action_lower:
                if "up" in action_lower or "increase" in action_lower:
                    self.display_message("System", "Increasing volume...")
                    for _ in range(5):
                        pyautogui.press('volumeup')
                    self.display_message("System", "Volume increased")
                elif "down" in action_lower or "decrease" in action_lower:
                    self.display_message("System", "Decreasing volume...")
                    for _ in range(5):
                        pyautogui.press('volumedown')
                    self.display_message("System", "Volume decreased")
                elif "mute" in action_lower or "unmute" in action_lower:
                    self.display_message("System", "Toggling mute...")
                    pyautogui.press('volumemute')
                    self.display_message("System", "Mute toggled")
            elif action_lower in ["clipboard history", "show clipboard"]:
                self.display_message("System", "Opening clipboard history...")
                pyautogui.hotkey('win', 'v')
                self.display_message("System", "Clipboard history opened")
            else:
                self.display_message("System", f"Unknown action: {action_name}")
                process_command_conversation("direct_action", f"Unknown action: {action_name}", action_name, False, "Not implemented")
        except Exception as e:
            self.display_message("System", f"Error executing action: {str(e)}")
            process_command_conversation("direct_action", f"Action failed: {action_name}", action_name, False, str(e))
        self._enable_inputs()
    def handle_send_message(self, event=None):
        """Handles sending a typed message to the AI."""
        user_message = self.input_entry.get()
        if not user_message.strip():
            return
        self.display_message("You", user_message)
        self.input_entry.delete(0, tk.END)
        self._disable_inputs()
        print(f"Text input received: '{user_message}'")
        is_command = False
        commands = re.split(r'\s+and\s+|\s*,\s*', user_message)
        if len(commands) > 1:
            print(f"Split text input into {len(commands)} commands: {commands}")
            for cmd_index, cmd_text in enumerate(commands):
                cmd_text = cmd_text.strip()
                if not cmd_text:
                    continue
                print(f"Processing text command #{cmd_index+1}: '{cmd_text}'")
                is_cmd_processed = self._process_command_text(cmd_text, add_to_queue=True, command_index=cmd_index)
                if is_cmd_processed:
                    is_command = True
        else:
            is_command = self._process_command_text(user_message, command_index=0)
        if not is_command:
            self.update_status_safe("Sending message to AI...")
            self.run_in_thread(self._query_ai_and_speak_thread, user_message, None)
    def _process_command_text(self, command_text, add_to_queue=False, command_index=0):
        """Process a single command text and identify if it's a click, type, or shortcut command.
        Returns True if recognized as a command, False otherwise."""
        command_text = command_text.strip()
        history_patterns = ["show history", "command history", "show command history", "view history", "display history"]
        if any(pattern in command_text.lower() for pattern in history_patterns):
            print(f"Command history request detected in: '{command_text}'")
            if add_to_queue:
                request_queue.put(("history", None))
            else:
                self.show_command_history()
            return True
        click_keywords = ["find", "click", "tap", "press", "select", "where is"]
        is_click_keyword_present = any(keyword in command_text.lower() for keyword in click_keywords)
        if is_click_keyword_present or (command_index > 0 and not any(kw in command_text.lower() for kw in ["type", "enter", "input", "write", "show history", "command history"])):
            print(f"Processing as click command: '{command_text}'")
            if command_index > 0 and not is_click_keyword_present:
                element_description = command_text
                print(f"Implicit click target: '{element_description}'")
                self.update_status_safe(f"Processing click: {element_description}")
                if add_to_queue:
                    request_queue.put(("click", element_description))
                else:
                    self.process_click_request(element_description)
                return True
            for keyword in click_keywords:
                pattern = r'\b' + re.escape(keyword) + r'\b'
                match = re.search(pattern, command_text, re.IGNORECASE)
                if match:
                    keyword_pos = match.start()
                    element_description = command_text[keyword_pos + len(keyword):].strip()
                    if element_description:
                        print(f"Click target: '{element_description}'")
                        self.update_status_safe(f"Processing click: {element_description}")
                        if add_to_queue:
                            request_queue.put(("click", element_description))
                        else:
                            self.process_click_request(element_description)
                            return True
            return False
        elif any(keyword in command_text.lower() for keyword in ["type", "enter", "input", "write"]):
            type_keywords = ["type", "enter", "input", "write"]
            for keyword in type_keywords:
                pattern = r'\b' + re.escape(keyword) + r'\b'
                match = re.search(pattern, command_text, re.IGNORECASE)
                if match:
                    keyword_pos = match.start()
                    text_to_type = command_text[keyword_pos + len(keyword):].strip()
                    if text_to_type.lower() == "key" and keyword.lower() == "enter":
                        print(f"Processing enter key shortcut from text input")
                        if add_to_queue:
                            request_queue.put(("shortcut", "enter key"))
                        else:
                            self.process_shortcut_command("enter key")
                    elif text_to_type:
                        print(f"Processing typing from text input: '{text_to_type}'")
                        if add_to_queue:
                            request_queue.put(("type", text_to_type))
                        else:
                            self.process_type_request(text_to_type)
                    return True
        elif self.is_shortcut_command(command_text):
            shortcut_command = self.extract_shortcut_command(command_text)
        if shortcut_command:
            print(f"Processing shortcut from text input: '{shortcut_command}'")
            if add_to_queue:
                request_queue.put(("shortcut", shortcut_command))
            else:
                self.process_shortcut_command(shortcut_command)
                return True
        return False
    def _disable_inputs(self):
        """Disable input entry, send button, and analyze button."""
        self.input_entry.configure(state="disabled")
        self.send_button.configure(state="disabled")
        self.analyze_screen_button.configure(state="disabled")
    def _enable_inputs(self):
        """Re-enable input entry, send button, and analyze button."""
        self.input_entry.configure(state="normal")
        self.send_button.configure(state="normal")
        self.analyze_screen_button.configure(state="normal")
    def continuous_listener(self, update_status_callback):
        """Continuously listens for voice commands."""
        r = sr.Recognizer()
        while True:
            with sr.Microphone() as source:
                r.adjust_for_ambient_noise(source)
                current_time = time.time()
                if self.listening_active and current_time > self.listening_timeout:
                    self.listening_active = False
                    print("Active listening timeout reached")
                    update_status_callback("Listening mode timed out. Say keyword to reactivate.")
                try:
                    update_status_callback(f"Listener: {'Active - listening for commands' if self.listening_active else f'Ready (Say \'{LISTENING_KEYWORD}\' to activate)'}")
                    print("Listening for voice commands...")
                    audio = r.listen(source, timeout=5, phrase_time_limit=8)
                    try:
                        text = r.recognize_google(audio).lower()
                        print(f"Recognized: {text}")
                        if self.waiting_for_calibration_confirmation:
                            app_name, element_desc, coords = self.waiting_for_calibration_confirmation
                            if LISTENING_KEYWORD in text and "yes" in text:
                                request_queue.put(("confirm_calibration_yes", element_desc))
                                continue
                            elif LISTENING_KEYWORD in text and "no" in text:
                                request_queue.put(("confirm_calibration_no", None))
                                continue
                        if LISTENING_KEYWORD in text or self.listening_active:
                            if LISTENING_KEYWORD in text:
                                text = text.replace(LISTENING_KEYWORD, "").strip()
                                self.listening_active = True
                                self.listening_timeout = time.time() + self.LISTENING_TIMEOUT_SECONDS
                                print(f"Active listening mode engaged for {self.LISTENING_TIMEOUT_SECONDS} seconds")
                            if self.listening_active:
                                self.listening_timeout = time.time() + self.LISTENING_TIMEOUT_SECONDS
                            if text:
                                text_lower = text.lower().strip()
                                if any(term in text_lower for term in ["minimize all", "minimise all", "show desktop"]):
                                    print(f"DIRECT ACTION: Minimize All Windows")
                                    request_queue.put(("action", "minimize all"))
                                    continue
                                elif any(term in text_lower for term in ["maximize all", "maximise all", "restore all", "restore windows"]):
                                    print(f"DIRECT ACTION: Maximize/Restore All Windows")
                                    request_queue.put(("action", "maximize all"))
                                    continue
                                elif any(term in text_lower for term in ["maximize window", "maximise window", "maximize current"]):
                                    print(f"DIRECT ACTION: Maximize Current Window")
                                    request_queue.put(("action", "maximize window"))
                                    continue
                                elif any(term in text_lower for term in ["lock computer", "lock screen", "lock pc", "lock windows"]):
                                    print(f"DIRECT ACTION: Lock Computer")
                                    request_queue.put(("action", "lock computer"))
                                    continue
                                elif any(term in text_lower for term in ["sleep", "sleep mode", "put computer to sleep"]):
                                    print(f"DIRECT ACTION: Sleep Computer")
                                    request_queue.put(("action", "sleep"))
                                    continue
                                elif any(term in text_lower for term in ["take screenshot", "screenshot", "capture screen"]):
                                    print(f"DIRECT ACTION: Take Screenshot")
                                    request_queue.put(("action", "take screenshot"))
                                    continue
                                elif any(term in text_lower for term in ["open settings", "windows settings"]):
                                    print(f"DIRECT ACTION: Open Settings")
                                    request_queue.put(("action", "open settings"))
                                    continue
                                elif any(term in text_lower for term in ["open file explorer", "explorer", "files"]):
                                    print(f"DIRECT ACTION: Open File Explorer")
                                    request_queue.put(("action", "open file explorer"))
                                    continue
                                elif any(term in text_lower for term in ["task manager", "open task manager"]):
                                    print(f"DIRECT ACTION: Open Task Manager")
                                    request_queue.put(("action", "task manager"))
                                    continue
                                elif any(term in text_lower for term in ["run", "open run dialog"]):
                                    print(f"DIRECT ACTION: Open Run Dialog")
                                    request_queue.put(("action", "run"))
                                    continue
                                elif "volume up" in text_lower or "increase volume" in text_lower:
                                    print(f"DIRECT ACTION: Volume Up")
                                    request_queue.put(("action", "volume up"))
                                    continue
                                elif "volume down" in text_lower or "decrease volume" in text_lower:
                                    print(f"DIRECT ACTION: Volume Down")
                                    request_queue.put(("action", "volume down"))
                                    continue
                                elif "mute" in text_lower or "unmute" in text_lower:
                                    print(f"DIRECT ACTION: Toggle Mute")
                                    request_queue.put(("action", "volume mute"))
                                    continue
                                elif "clipboard" in text_lower:
                                    print(f"DIRECT ACTION: Show Clipboard")
                                    request_queue.put(("action", "clipboard history"))
                                    continue
                                print(f"Full command text after keyword removal: '{text}'")
                                commands = re.split(r'\s+and\s+|\s*,\s*', text)
                                print(f"Split into {len(commands)} commands: {commands}")
                                for command_index, command_text in enumerate(commands):
                                    command_text = command_text.strip()
                                    if not command_text: continue
                                    print(f"Processing command #{command_index+1}: '{command_text}'") # Enhanced debug
                                    if ("training mode" in command_text or "train mode" in command_text) and ("activate" in command_text or "start" in command_text or "on" in command_text):
                                        self.activate_training_mode()
                                        continue
                                    elif ("training mode" in command_text or "train mode" in command_text) and ("deactivate" in command_text or "stop" in command_text or "off" in command_text):
                                        self.deactivate_training_mode()
                                        continue
                                    elif self.training_mode_active and ("it's a" in command_text or "this is a" in command_text or "this is the" in command_text):
                                        self.process_training_command(command_text)
                                        continue
                                    click_keywords = ["find", "click", "tap", "press", "select", "where is"]
                                    is_click_keyword_present = any(keyword in command_text for keyword in click_keywords)
                                    if is_click_keyword_present or (command_index > 0 and not any(kw in command_text for kw in ["type", "enter", "input", "write", "show history", "command history"])):
                                        print(f"  ↪ Processing as click command: '{command_text}'")
                                        if command_index > 0 and not is_click_keyword_present:
                                            element_description = command_text
                                            print(f"  ↪ Implicit click target: '{element_description}'")
                                            update_status_callback(f"Processing click: {element_description}")
                                            request_queue.put(("click", element_description))
                                            continue
                                        keyword_found = False
                                        for keyword in click_keywords:
                                            pattern = r'\b' + re.escape(keyword) + r'\b'
                                            match = re.search(pattern, command_text, re.IGNORECASE)
                                            if match:
                                                keyword_found = True
                                                keyword_pos = match.start()
                                                element_description = command_text[keyword_pos + len(keyword):].strip()
                                                if element_description:
                                                    print(f"  ↪ Click target: '{element_description}'")
                                                    update_status_callback(f"Processing click: {element_description}")
                                                    request_queue.put(("click", element_description))
                                                    break
                                                else:
                                                    print(f"  ↪ Warning: Found keyword '{keyword}' but no element description follows")
                                        if not keyword_found and is_click_keyword_present:
                                            print(f"  ↪ Warning: Click-type command detected but couldn't identify specific keyword")
                                        continue
                                    elif self.is_shortcut_command(command_text):
                                        shortcut_command = self.extract_shortcut_command(command_text)
                                        if shortcut_command:
                                            print(f"Shortcut command detected: '{shortcut_command}'")
                                            update_status_callback(f"Processing shortcut: {shortcut_command}")
                                            request_queue.put(("shortcut", shortcut_command))
                                        continue
                                    elif any(keyword in command_text for keyword in ["type", "enter", "input", "write"]):
                                        print(f"  ↪ Detected type command: '{command_text}'")
                                        type_keywords = ["type", "enter", "input", "write"]
                                        keyword_found = False
                                        for keyword in type_keywords:
                                            pattern = r'\b' + re.escape(keyword) + r'\b'
                                            match = re.search(pattern, command_text, re.IGNORECASE)
                                            if match:
                                                keyword_found = True
                                                keyword_pos = match.start()
                                                text_to_type = command_text[keyword_pos + len(keyword):].strip()
                                                if text_to_type.lower() == "key" and keyword.lower() == "enter":
                                                    print(f"  ↪ Mapping 'enter key' to shortcut.")
                                                    request_queue.put(("shortcut", "enter key"))
                                                elif text_to_type:
                                                    print(f"  ↪ Text to type: '{text_to_type}'")
                                                    update_status_callback(f"Processing typing: {text_to_type}")
                                                    request_queue.put(("type", text_to_type))
                                                else:
                                                    print(f"  ↪ Warning: Found keyword '{keyword}' but no text to type follows")
                                                break
                                        if not keyword_found:
                                            print(f"  ↪ Warning: Type command detected but couldn't identify specific keyword")
                                        continue
                                    elif any(pattern in command_text.lower() for pattern in ["show history", "command history", "show command history", "view history", "display history"]):
                                        print(f"  ↪ History command detected: '{command_text}'")
                                        update_status_callback("Processing history request...")
                                        request_queue.put(("history", None))
                                        continue
                                    elif "minimize all" in command_text.lower() or "show desktop" in command_text.lower() or "m minimise all" in command_text.lower():
                                        print(f"  ↪ Special system action: Minimize All Windows")
                                        update_status_callback("Minimizing all windows...")
                                        request_queue.put(("system", "powershell -c \"(New-Object -ComObject Shell.Application).MinimizeAll()\""))
                                        continue
                                    elif ("run" in command_text.lower() and "command" in command_text.lower()) or ("execute" in command_text.lower()) or command_text.lower().startswith("cmd "):
                                        cmd_match = None
                                        if "run command" in command_text.lower():
                                            cmd_match = re.search(r'run command\s+(.*)', command_text, re.IGNORECASE)
                                        elif "execute" in command_text.lower():
                                            cmd_match = re.search(r'execute\s+(.*)', command_text, re.IGNORECASE)
                                        elif command_text.lower().startswith("cmd "):
                                            cmd_match = re.search(r'cmd\s+(.*)', command_text, re.IGNORECASE)
                                        if cmd_match:
                                            system_cmd = cmd_match.group(1).strip()
                                            print(f"  ↪ System command detected: '{system_cmd}'")
                                            update_status_callback(f"Running system command: {system_cmd}")
                                            request_queue.put(("system", system_cmd))
                                        else:
                                            print(f"  ↪ Detected system command keyword but no command found")
                                            update_status_callback("No command specified")
                                        continue
                                    elif len(commands) == 1:
                                        print(f"Sending general query: '{command_text}'")
                                        update_status_callback(f"Processing query: {command_text}")
                                        request_queue.put(("query", command_text))
                                    else:
                                        print(f"Ignoring non-action segment in multi-command sequence: '{command_text}'")
                    except sr.UnknownValueError:
                        pass
                    except sr.RequestError as e:
                        print(f"Could not request results from Google Speech Recognition service; {e}")
                    except Exception as e:
                        print(f"Error during voice recognition: {e}")
                except sr.WaitTimeoutError:
                    pass
                except Exception as e:
                    print(f"Error during listening: {e}")
            time.sleep(0.1)
    def process_start_calibration(self, element_description):
        """Initiates the manual coordinate capture process."""
        if not self.waiting_for_calibration_confirmation:
            app_info = get_active_app_info()
            app_name = app_info["app_name"]
        else:
            app_name, element_description, _ = self.waiting_for_calibration_confirmation
        self.analyze_screen_button.configure(state="disabled")
        self.display_message("System", f"Starting calibration for: '{element_description}' in '{app_name}'")
        self.update_status_safe("Prepare to position mouse...")
        self.run_in_thread(self._interactive_manual_position_thread, app_name, element_description)
        self.waiting_for_calibration_confirmation = None
    def _interactive_manual_position_thread(self, app_name, element_description):
        """Guides user to position mouse and captures coordinates."""
        try:
            self.after(0, self.display_message, "System", 
                       "Position your mouse cursor precisely over "
                       f"'{element_description}' in '{app_name}' within 5 seconds.")
            self.after(0, speak_text, "Position mouse over the element.")
            for i in range(5, 0, -1):
                self.after(0, self.update_status, f"Capturing position in {i}...")
                time.sleep(1)
            self.after(0, self.update_status, "Capturing...")
            x, y = pyautogui.position()
            new_coords = (int(x), int(y))
            save_success = save_coordinates(element_description, new_coords[0], new_coords[1], True, app_name)
            if save_success:
                 self.after(0, self.display_message, "System", 
                            f"Calibration successful! New coordinates {new_coords} saved for "
                            f"'{element_description}' in '{app_name}'.")
                 self.after(0, self.update_status, f"Listener: Ready (Say '{LISTENING_KEYWORD}' + command)")
            else:
                 self.after(0, self.display_message, "System", 
                            f"Calibration failed: Could not save coordinates for '{element_description}' in '{app_name}'.")
                 self.after(0, self.update_status, "Calibration failed.")
            try:
                region = (max(0, x-30), max(0, y-30), x+30, y+30)
                screenshot = pyautogui.screenshot(region=region)
                safe_app = re.sub(r'[\\/*?:"<>|]', "", app_name)
                safe_desc = re.sub(r'[\\/*?:"<>|]', "", element_description)
                filename = f"calibration_{safe_app[:15]}_{safe_desc[:20]}_{int(time.time())}.png"
                screenshot.save(filename)
                self.after(0, self.display_message, "System", f"Saved verification image: {filename}")
            except Exception as img_e:
                print(f"[Calibration] Could not save verification image: {img_e}")
                pass 
        except Exception as e:
            self.after(0, self.display_message, "System", f"Error during manual calibration: {e}")
            self.after(0, self.update_status, "Calibration error.")
        finally:
            self.after(0, self._enable_inputs())
    def activate_training_mode(self):
        """Activate training mode for UI element identification"""
        self.training_mode_active = True
        print("Training mode activated")
        self.display_message("System", "Training mode activated. You can now identify UI elements by saying 'It's a [element name]'")
        speak_text("Training mode activated")
    def deactivate_training_mode(self):
        """Deactivate training mode"""
        self.training_mode_active = False
        print("Training mode deactivated")
        self.display_message("System", "Training mode deactivated")
        speak_text("Training mode deactivated")
    def process_training_command(self, text):
        """Process a training command to identify UI elements"""
        element_name = None
        if "it's a" in text:
            match = re.search(r"it's a\s+(.+)", text, re.IGNORECASE)
            if match:
                element_name = match.group(1).strip()
        elif "this is a" in text:
            match = re.search(r"this is a\s+(.+)", text, re.IGNORECASE)
            if match:
                element_name = match.group(1).strip()
        elif "this is the" in text:
            match = re.search(r"this is the\s+(.+)", text, re.IGNORECASE)
            if match:
                element_name = match.group(1).strip()
        if element_name:
            self.capture_ui_element(element_name)
        else:
            self.display_message("System", "Could not understand the element name. Please try again with 'It's a [element name]'.")
            speak_text("Could not understand the element name. Please try again.")
    def capture_ui_element(self, element_name):
        """Capture the current mouse position and store it with the element name"""
        try:
            x, y = pyautogui.position()
            save_coordinates(element_name, x, y, True)
            screenshot = pyautogui.screenshot(region=(x-50, y-50, 100, 100))
            os.makedirs("training_data", exist_ok=True)
            timestamp = int(time.time())
            screenshot_path = f"training_data/{element_name.replace(' ', '_')}_{timestamp}.png"
            screenshot.save(screenshot_path)
            if element_name not in self.training_elements_db:
                self.training_elements_db[element_name] = []
            self.training_elements_db[element_name].append({
                "x": x,
                "y": y,
                "timestamp": timestamp,
                "screenshot_path": screenshot_path
            })
            try:
                with open("training_data/elements.json", "w") as f:
                    json.dump(self.training_elements_db, f, indent=2)
            except Exception as json_err:
                print(f"Error saving training data to JSON: {json_err}")
            self.display_message("System", f"Captured UI element: '{element_name}' at position ({x}, {y})")
            speak_text(f"Captured {element_name}")
            print(f"Captured UI element: '{element_name}' at position ({x}, {y})")
        except Exception as e:
            error_msg = f"Error capturing UI element: {e}"
            self.display_message("System", error_msg)
            speak_text("Error capturing UI element")
            print(error_msg)
    def get_shortcut_mappings(self):
        """Returns a dictionary mapping voice keywords/phrases to shortcut keys."""
        return {
            "copy": ('ctrl', 'c'),
            "paste": ('ctrl', 'v'),
            "cut": ('ctrl', 'x'),
            "undo": ('ctrl', 'z'),
            "redo": ('ctrl', 'y'),
            "select all": ('ctrl', 'a'),
            "save": ('ctrl', 's'),
            "open file": ('ctrl', 'o'),
            "new file": ('ctrl', 'n'),
            "print": ('ctrl', 'p'),
            "close tab": ('ctrl', 'w'),
            "close window": ('alt', 'f4'),
            "new tab": ('ctrl', 't'),
            "switch window": ('alt', 'tab'),
            "next window": ('alt', 'tab'),
            "previous window": ('alt', 'shift', 'tab'),
            "minimize all": ('win', 'd'),
            "show desktop": ('win', 'd'),
            "find": ('ctrl', 'f'),
            "refresh": ('f5',),
            "escape": ('esc',),
            "enter key": ('enter',),
            "tab key": ('tab',),
        }
    def is_shortcut_command(self, text):
        """Checks if the text contains any known shortcut keywords."""
        mappings = self.get_shortcut_mappings()
        text_lower = text.lower()
        for key_phrase in mappings.keys():
            pattern = r'\b' + re.escape(key_phrase) + r'\b'
            if re.search(pattern, text_lower, re.IGNORECASE):
                return True
        return False
    def extract_shortcut_command(self, text):
        """Finds the best matching shortcut command from the text."""
        mappings = self.get_shortcut_mappings()
        text_lower = text.lower()
        best_match = None
        best_match_len = 0
        for key_phrase in sorted(mappings.keys(), key=len, reverse=True):
            pattern = r'\b' + re.escape(key_phrase) + r'\b'
            match = re.search(pattern, text_lower, re.IGNORECASE)
            if match:
                phrase_len = len(key_phrase)
                if phrase_len > best_match_len:
                    best_match = key_phrase
                    best_match_len = phrase_len
        return best_match
    def show_command_history(self):
        """Display a window with the command history."""
        try:
            history_window = ctk.CTkToplevel(self)
            history_window.title("Conversation & Command History")
            history_window.geometry("900x700")
            history_window.transient(self)
            history_window.grab_set()
            history_window.grid_columnconfigure(0, weight=1)
            history_window.grid_rowconfigure(0, weight=0)
            history_window.grid_rowconfigure(1, weight=0)
            history_window.grid_rowconfigure(2, weight=0)
            history_window.grid_rowconfigure(3, weight=1)
            history_window.grid_rowconfigure(4, weight=0)
            title_label = ctk.CTkLabel(history_window, text="Command & Conversation History", 
                                      font=ctk.CTkFont(size=20, weight="bold"))
            title_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
            filter_frame = ctk.CTkFrame(history_window)
            filter_frame.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="ew")
            filter_frame.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)
            self.filter_var = tk.StringVar(value="all")
            filter_label = ctk.CTkLabel(filter_frame, text="Filter by:")
            filter_label.grid(row=0, column=0, padx=5, pady=10, sticky="w")
            all_radio = ctk.CTkRadioButton(filter_frame, text="All Entries", variable=self.filter_var, 
                                         value="all", command=lambda: self.refresh_history_display())
            all_radio.grid(row=0, column=1, padx=10, pady=10)
            commands_radio = ctk.CTkRadioButton(filter_frame, text="Commands Only", variable=self.filter_var, 
                                              value="commands", command=lambda: self.refresh_history_display())
            commands_radio.grid(row=0, column=2, padx=10, pady=10)
            convos_radio = ctk.CTkRadioButton(filter_frame, text="Conversations Only", variable=self.filter_var, 
                                            value="conversations", command=lambda: self.refresh_history_display())
            convos_radio.grid(row=0, column=3, padx=10, pady=10)
            coords_radio = ctk.CTkRadioButton(filter_frame, text="Coordinates", variable=self.filter_var, 
                                           value="coordinates", command=lambda: self.refresh_history_display())
            coords_radio.grid(row=0, column=4, padx=10, pady=10)
            search_frame = ctk.CTkFrame(history_window)
            search_frame.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
            search_frame.grid_columnconfigure(0, weight=0)
            search_frame.grid_columnconfigure(1, weight=1)
            search_frame.grid_columnconfigure(2, weight=0)
            search_label = ctk.CTkLabel(search_frame, text="Search:")
            search_label.grid(row=0, column=0, padx=5, pady=10, sticky="w")
            self.search_entry = ctk.CTkEntry(search_frame, placeholder_text="Enter element description or app name...")
            self.search_entry.grid(row=0, column=1, padx=5, pady=10, sticky="ew")
            search_button = ctk.CTkButton(search_frame, text="Search", 
                                       command=lambda: self._search_coordinates_display())
            search_button.grid(row=0, column=2, padx=5, pady=10, sticky="e")
            self.search_entry.bind('<Return>', lambda event: self._search_coordinates_display())
            self.history_content_frame = ctk.CTkScrollableFrame(history_window)
            self.history_content_frame.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")
            self.history_content_frame.grid_columnconfigure(0, weight=1)
            self.history_window = history_window
            try:
                self.history_items = log_manager.get_all_conversations(limit=100)
                self.coordinates_items = log_manager.get_saved_coordinates()
                self.filtered_coordinates = None
                self.refresh_history_display()
            except Exception as e:
                error_label = ctk.CTkLabel(self.history_content_frame, text=f"Error loading history: {str(e)}")
                error_label.grid(row=0, column=0, pady=10, padx=10, sticky="w")
            button_frame = ctk.CTkFrame(history_window, fg_color="transparent")
            button_frame.grid(row=4, column=0, pady=10, padx=20, sticky="e")
            refresh_button = ctk.CTkButton(button_frame, text="Refresh", 
                                        command=lambda: self._reload_and_refresh_display())
            refresh_button.grid(row=0, column=0, padx=(0, 10))
            clear_button = ctk.CTkButton(button_frame, text="Clear History", 
                                      command=lambda: self._clear_history_display())
            clear_button.grid(row=0, column=1, padx=(0, 10))
            close_button = ctk.CTkButton(button_frame, text="Close", command=history_window.destroy)
            close_button.grid(row=0, column=2)
        except Exception as e:
            self.display_message("Error", f"Failed to display history: {e}")
            print(f"Error displaying history: {e}")
    def _search_coordinates_display(self):
        """Search and filter coordinates based on user input."""
        search_text = self.search_entry.get().lower().strip()
        if not search_text:
            self.filtered_coordinates = None
        else:
            self.filtered_coordinates = [
                item for item in self.coordinates_items
                if (search_text in str(item.get('element_description', '')).lower() or
                    search_text in str(item.get('app_name', '')).lower())
            ]
        self.refresh_history_display()
    def _reload_and_refresh_display(self):
        """Reload data from the database and refresh the display."""
        self.coordinates_items = log_manager.get_saved_coordinates()
        self.history_items = log_manager.get_all_conversations(limit=100)
        self.filtered_coordinates = None
        if hasattr(self, 'search_entry'):
            self.search_entry.delete(0, tk.END)
        self.refresh_history_display()
        self.display_message("System", "History data refreshed from database.")
    def _clear_history_display(self):
        """Clears the history display."""
        confirm = messagebox.askyesno("Confirm", "Are you sure you want to clear the history?")
        if confirm:
            filter_type = self.filter_var.get()
            if filter_type == "coordinates":
                self.coordinates_items = []
            else:
                self.history_items = []
            self.refresh_history_display()
    def refresh_history_display(self):
        """Refresh the history display based on selected filter."""
        for widget in self.history_content_frame.winfo_children():
            widget.destroy()
        filter_type = self.filter_var.get()
        if filter_type == "coordinates":
            display_items = self.filtered_coordinates if self.filtered_coordinates is not None else self.coordinates_items
            if not display_items:
                if self.filtered_coordinates is not None:
                    label = ctk.CTkLabel(self.history_content_frame, text="No coordinates match your search criteria.")
                    label.grid(row=0, column=0, pady=10, sticky="w")
                else:
                    label = ctk.CTkLabel(self.history_content_frame, text="No saved coordinates available.")
                    label.grid(row=0, column=0, pady=10, sticky="w")
                return
            coord_count = len(display_items)
            total_count = len(self.coordinates_items)
            if self.filtered_coordinates is not None:
                header = ctk.CTkLabel(self.history_content_frame, 
                                    text=f"FILTERED COORDINATES: Showing {coord_count} of {total_count} total items",
                                    font=ctk.CTkFont(size=14, weight="bold"))
            else:
                header = ctk.CTkLabel(self.history_content_frame, 
                                    text=f"DISPLAYING ALL COORDINATES ({coord_count} total items)",
                                    font=ctk.CTkFont(size=14, weight="bold"))
            header.grid(row=0, column=0, pady=(0, 5), sticky="w")
            instructions = ctk.CTkLabel(self.history_content_frame, 
                                       text="Use Edit buttons to modify coordinates and Delete to remove them.")
            instructions.grid(row=1, column=0, pady=(0, 10), sticky="w")
            print(f"Displaying {coord_count} coordinates in history view")
            for i, item in enumerate(display_items):
                self._create_coordinate_frame(i+2, item)
            print(f"Successfully displayed {coord_count} coordinates in frames")
            return
        if not self.history_items:
            label = ctk.CTkLabel(self.history_content_frame, text="No history available.")
            label.grid(row=0, column=0, pady=10, sticky="w")
            return
        row_count = 0
        for item in self.history_items:
            is_command = item.get('command_type') is not None
            has_ai_response = item.get('ai_response') is not None and item.get('ai_response') != ''
            if filter_type == "commands" and not is_command:
                continue
            if filter_type == "conversations" and not has_ai_response:
                continue
            item_frame = ctk.CTkFrame(self.history_content_frame)
            item_frame.grid(row=row_count, column=0, pady=(0, 10), sticky="ew")
            item_frame.grid_columnconfigure(0, weight=1)
            timestamp = item.get('timestamp', 'Unknown')
            if is_command:
                cmd_type = item.get('command_type', 'Unknown')
                cmd_text = item.get('user_query', 'Unknown')
                target = item.get('target', 'None')
                success = "Success" if item.get('success', 0) == 1 else "Failed"
                details = item.get('details', 'None')
                header_label = ctk.CTkLabel(item_frame, 
                                         text=f"[{timestamp}] COMMAND ({cmd_type}): {cmd_text}",
                                         font=ctk.CTkFont(weight="bold"))
                header_label.grid(row=0, column=0, padx=10, pady=(5, 0), sticky="w")
                target_label = ctk.CTkLabel(item_frame, text=f"Target: {target}")
                target_label.grid(row=1, column=0, padx=10, sticky="w")
                status_label = ctk.CTkLabel(item_frame, text=f"Status: {success}")
                status_label.grid(row=2, column=0, padx=10, sticky="w")
                if details != 'None':
                    details_label = ctk.CTkLabel(item_frame, text=f"Details: {details}")
                    details_label.grid(row=3, column=0, padx=10, pady=(0, 5), sticky="w")
            elif has_ai_response:
                user_query = item.get('user_query', 'Unknown')
                ai_response = item.get('ai_response', 'Unknown')
                has_screenshot = "Yes" if item.get('has_screenshot', 0) == 1 else "No"
                header_label = ctk.CTkLabel(item_frame, 
                                         text=f"[{timestamp}] CONVERSATION",
                                         font=ctk.CTkFont(weight="bold"))
                header_label.grid(row=0, column=0, padx=10, pady=(5, 0), sticky="w")
                user_label = ctk.CTkLabel(item_frame, text=f"User: {user_query}")
                user_label.grid(row=1, column=0, padx=10, sticky="w")
                ai_label = ctk.CTkLabel(item_frame, text=f"AI: {ai_response}")
                ai_label.grid(row=2, column=0, padx=10, sticky="w")
                screenshot_label = ctk.CTkLabel(item_frame, text=f"Screenshot: {has_screenshot}")
                screenshot_label.grid(row=3, column=0, padx=10, pady=(0, 5), sticky="w")
            row_count += 1
    def _create_coordinate_frame(self, row_position, item):
        """Create a frame for a single coordinate entry with edit and delete buttons."""
        item_id = item.get('id', 0)
        timestamp = item.get('timestamp', 'Unknown')
        app_name = item.get('app_name', 'Unknown App')
        element = item.get('element_description', 'Unknown Element')
        x, y = item.get('x', 0), item.get('y', 0)
        success_count = item.get('success_count', 0)
        fail_count = item.get('fail_count', 0)
        item_frame = ctk.CTkFrame(self.history_content_frame, fg_color=("#f0f0f0", "#333333"))
        item_frame.grid(row=row_position, column=0, pady=(0, 10), sticky="ew")
        item_frame.grid_columnconfigure(0, weight=1)
        header_frame = ctk.CTkFrame(item_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=(5, 0), sticky="ew")
        header_frame.grid_columnconfigure(0, weight=1)
        header_label = ctk.CTkLabel(header_frame, 
                                  text=f"[{timestamp}] COORDINATES #{item_id} - {app_name}",
                                  font=ctk.CTkFont(weight="bold"))
        header_label.grid(row=0, column=0, sticky="w")
        element_label = ctk.CTkLabel(item_frame, text=f"Element: {element}")
        element_label.grid(row=1, column=0, padx=10, sticky="w")
        coords_label = ctk.CTkLabel(item_frame, text=f"Position: ({x}, {y})")
        coords_label.grid(row=2, column=0, padx=10, sticky="w")
        stats_label = ctk.CTkLabel(item_frame, text=f"Success/Fail: {success_count}/{fail_count}")
        stats_label.grid(row=3, column=0, padx=10, sticky="w")
        button_frame = ctk.CTkFrame(item_frame, fg_color="transparent")
        button_frame.grid(row=4, column=0, padx=10, pady=(0, 5), sticky="w")
        edit_button = ctk.CTkButton(button_frame, text="Edit", width=80,
                                  command=lambda id=item_id, elm=element, x_val=x, y_val=y, app=app_name: 
                                      self._edit_coordinates_dialog(id, elm, x_val, y_val, app))
        edit_button.grid(row=0, column=0, padx=(0, 5))
        delete_button = ctk.CTkButton(button_frame, text="Delete", width=80,
                                    fg_color="darkred", hover_color="red",
                                    command=lambda id=item_id: self._delete_coordinates_dialog(id))
        delete_button.grid(row=0, column=1)
    def _edit_coordinates_dialog(self, coord_id, element, x, y, app_name):
        """Open a dialog to edit a coordinate record."""
        edit_window = ctk.CTkToplevel(self.history_window)
        edit_window.title(f"Edit Coordinates #{coord_id}")
        edit_window.geometry("400x300")
        edit_window.transient(self.history_window)
        edit_window.grab_set()
        edit_window.grid_columnconfigure(0, weight=1)
        edit_window.grid_columnconfigure(1, weight=2)
        ctk.CTkLabel(edit_window, text="Element:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        element_entry = ctk.CTkEntry(edit_window, width=250)
        element_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        element_entry.insert(0, element)
        ctk.CTkLabel(edit_window, text="X Coordinate:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        x_entry = ctk.CTkEntry(edit_window, width=100)
        x_entry.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        x_entry.insert(0, str(x))
        ctk.CTkLabel(edit_window, text="Y Coordinate:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        y_entry = ctk.CTkEntry(edit_window, width=100)
        y_entry.grid(row=2, column=1, padx=10, pady=10, sticky="w")
        y_entry.insert(0, str(y))
        ctk.CTkLabel(edit_window, text="App Name:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
        app_entry = ctk.CTkEntry(edit_window, width=250)
        app_entry.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        app_entry.insert(0, app_name)
        status_label = ctk.CTkLabel(edit_window, text="")
        status_label.grid(row=4, column=0, columnspan=2, padx=10, pady=10)
        button_frame = ctk.CTkFrame(edit_window, fg_color="transparent")
        button_frame.grid(row=5, column=0, columnspan=2, padx=10, pady=10)
        def save_changes():
            try:
                new_element = element_entry.get()
                new_x = int(x_entry.get())
                new_y = int(y_entry.get())
                new_app = app_entry.get()
                result = log_manager.update_coordinates(coord_id, new_element, new_x, new_y, new_app)
                if result:
                    status_label.configure(text="Changes saved successfully!", text_color="green")
                    self.coordinates_items = log_manager.get_saved_coordinates()
                    self.refresh_history_display()
                    edit_window.after(1500, edit_window.destroy)
                else:
                    status_label.configure(text="Failed to save changes.", text_color="red")
            except ValueError:
                status_label.configure(text="X and Y must be integers.", text_color="red")
            except Exception as e:
                status_label.configure(text=f"Error: {e}", text_color="red")
        save_button = ctk.CTkButton(button_frame, text="Save Changes", command=save_changes)
        save_button.grid(row=0, column=0, padx=(0, 10))
        cancel_button = ctk.CTkButton(button_frame, text="Cancel", command=edit_window.destroy)
        cancel_button.grid(row=0, column=1)
    def _delete_coordinates_dialog(self, coord_id):
        """Delete a coordinate record after confirmation."""
        confirm = messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete coordinate #{coord_id}?")
        if confirm:
            result = log_manager.delete_coordinates(coord_id)
            if result:
                self.coordinates_items = log_manager.get_saved_coordinates()
                self.refresh_history_display()
                self.display_message("System", f"Coordinate #{coord_id} deleted successfully.")
            else:
                self.display_message("Error", f"Failed to delete coordinate #{coord_id}.")
    def _do_shortcut_thread(self, keys_to_press, command_name):
        """Performs the shortcut action in a thread."""
        try:
            pyautogui.hotkey(*keys_to_press)
            print(f"Shortcut '{command_name}' executed successfully.")
            self.after(0, self.display_message, "System", f"Shortcut '{command_name.capitalize()}' executed.")
            process_command_conversation("shortcut_execute", f"Executed shortcut: {command_name}", 
                                      command_name, True, f"Keys: {' + '.join(keys_to_press)}")
        except Exception as e:
            print(f"Error executing shortcut '{command_name}': {e}")
            self.after(0, self.display_message, "System", f"Failed to execute shortcut '{command_name.capitalize()}'. Error: {e}")
            process_command_conversation("shortcut_execute", f"Failed shortcut: {command_name}", 
                                      command_name, False, f"Error: {e}")
        finally:
            if self.listening_active:
                self.after(0, self.update_status, "Listener: Active (Listening...)")
            else:
                self.after(0, self.update_status, f"Listener: Ready (Say '{LISTENING_KEYWORD}' to activate)")
            self.after(0, self._enable_inputs())
    def execute_action(self, action_name):
        """Execute a direct system action like 'minimize all'."""
        self._disable_inputs()
        self.display_message("Action", f"Executing: {action_name}")
        process_command_conversation("direct_action", f"Action: {action_name}", action_name)
        if action_name.lower() in ["minimize all", "minimise all", "show desktop"]:
            try:
                import subprocess
                self.display_message("System", "Minimizing all windows...")
                result = subprocess.run("powershell -c \"(New-Object -ComObject Shell.Application).MinimizeAll()\"", 
                                     shell=True, capture_output=True, text=True)
                if result.returncode == 0:
                    self.display_message("System", "All windows minimized successfully")
                    process_command_conversation("direct_action", "Minimize all windows", "desktop", True, "Success")
                else:
                    self.display_message("System", "Using keyboard shortcut instead...")
                    pyautogui.hotkey('win', 'd')
                    self.display_message("System", "All windows minimized using keyboard shortcut")
                    process_command_conversation("direct_action", "Minimize all windows", "desktop", True, "Used shortcut")
            except Exception as e:
                self.display_message("System", f"Error minimizing windows: {str(e)}")
                process_command_conversation("direct_action", "Minimize all windows", "desktop", False, str(e))
        else:
            self.display_message("System", f"Unknown action: {action_name}")
            process_command_conversation("direct_action", f"Unknown action: {action_name}", action_name, False, "Not implemented")
        self._enable_inputs()
if __name__ == "__main__":
    if not configure_genai():
        input("API Key configuration failed. Please set the HARDCODED_API_KEY in the script and restart. Press Enter to exit.")
    else:
        setup_database()
        app = ChatScreenApp()
        app.mainloop() 