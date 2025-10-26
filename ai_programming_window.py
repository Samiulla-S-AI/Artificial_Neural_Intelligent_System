import tkinter as tk
from tkinter import ttk, scrolledtext, Label, Frame, Button, Text, END, NORMAL, DISABLED, WORD, filedialog, messagebox
import google.generativeai as genai
import threading
import logging
import re

class AIProgrammingWindow:
    def __init__(self, parent):
        self.window = tk.Toplevel(parent.root)
        self.window.title("AI Programming Assistant")
        self.window.geometry("1200x800")
        self.window.configure(bg="#0C141F")
        
        # Store the main application instance as parent
        self.main_app = parent
        self.buttons = {}
        
        # Initialize Gemini
        try:
            genai.configure(api_key="AIzaSyAqlSk_zL7ID0a_tiBP_E6sIurmXB43F4k")
            self.model = genai.GenerativeModel('gemini-2.0-flash')
            self.chat = self.model.start_chat(history=[])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to initialize Gemini: {str(e)}")
            self.window.destroy()
            return
            
        # Initialize other attributes
        self.requirements_text = None
        self.code_display = None
        self.status_var = tk.StringVar()
        self.suggestion_window = None
        self.error_analysis_window = None
        self.code_completion_active = False
        self.completion_thread = None
        
        self.create_widgets()
        self.center_window()
        self.setup_bindings()
        
        # Make window modal
        self.window.transient(parent.root)
        self.window.grab_set()
        
        # Set window close handler
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        # Create main container
        main_frame = ttk.Frame(self.window, style="TRON.TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Requirements Frame
        req_frame = ttk.Frame(main_frame, style="TRON.TFrame")
        req_frame.pack(fill=tk.X, pady=(0, 10))
        
        req_label = ttk.Label(req_frame, 
                            text="PROGRAM REQUIREMENTS", 
                            style="TRON.TLabel")
        req_label.pack(fill=tk.X, padx=5, pady=(0, 5))

        self.requirements_text = tk.Text(req_frame, 
                              height=5, 
                              wrap=tk.WORD,
                              bg="#1A2030",
                              fg="#00F0FF",
                              insertbackground="#00F0FF",
                              font=('Share Tech Mono', 12),
                              bd=2,
                              relief="solid",
                              highlightthickness=1,
                              highlightcolor="#00F0FF",
                              highlightbackground="#004A5D")
        self.requirements_text.pack(fill=tk.X, padx=5, pady=5)
        
        # Example requirements
        example_text = "Example: Create a Python program that reads a CSV file, processes the data, and generates a bar chart using matplotlib."
        self.requirements_text.insert('1.0', example_text)
        self.requirements_text.bind('<FocusIn>', lambda e: self.clear_example_text(e))

        # Buttons Frame with glowing effect
        btn_frame = ttk.Frame(main_frame, style="TRON.TFrame")
        btn_frame.pack(fill=tk.X, pady=(0, 15))

        button_configs = [
            ("GENERATE PROGRAM", self.generate_program, "Ctrl+Enter"),
            ("ANALYZE CODE", self.analyze_code, "Ctrl+A"),
            ("GET SUGGESTIONS", self.get_code_suggestions, "Ctrl+S"),
            ("CLEAR ALL", self.clear_all, "Ctrl+L"),
            ("COPY CODE", self.copy_to_clipboard, "Ctrl+C"),
            ("SAVE CODE", self.save_code, "Ctrl+Shift+S")
        ]

        for text, command, shortcut in button_configs:
            btn = tk.Button(btn_frame,
                          text=f"{text}\n{shortcut}",
                          command=command,
                          font=('Orbitron', 9),
                          bg="#0C141F",
                          fg="#00F0FF",
                          activebackground="#00F0FF",
                          activeforeground="#0C141F",
                          relief="solid",
                          bd=1,
                          padx=15,
                          pady=5)
            btn.pack(side=tk.LEFT, padx=5, expand=True)
            
            # Store button reference
            self.buttons[text] = btn

            # Bind hover events for glow effect
            btn.bind('<Enter>', lambda e, b=btn: self.on_button_hover(e, b))
            btn.bind('<Leave>', lambda e, b=btn: self.on_button_leave(e, b))

        # Code Editor Frame
        code_frame = ttk.Frame(main_frame, style="TRON.TFrame")
        code_frame.pack(fill=tk.BOTH, expand=True)

        # Code editor with line numbers
        editor_frame = ttk.Frame(code_frame, style="TRON.TFrame")
        editor_frame.pack(fill=tk.BOTH, expand=True)

        # Line numbers
        self.line_numbers = tk.Text(editor_frame,
                                  width=4,
                                  padx=3,
                                  pady=5,
                                  takefocus=0,
                                  border=0,
                                  background="#1A2030",
                                  foreground="#00F0FF",
                                  font=('Share Tech Mono', 12),
                                  state='disabled')
        self.line_numbers.pack(side=tk.LEFT, fill=tk.Y)

        # Code editor
        self.code_display = tk.Text(editor_frame,
                                  wrap=tk.NONE,
                                  bg="#1A2030",
                                  fg="#00F0FF",
                                  insertbackground="#00F0FF",
                                  font=('Share Tech Mono', 12),
                                  bd=2,
                                  relief="solid",
                                  highlightthickness=1,
                                  highlightcolor="#00F0FF",
                                  highlightbackground="#004A5D")
        self.code_display.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Scrollbars
        y_scrollbar = ttk.Scrollbar(editor_frame, 
                                  orient=tk.VERTICAL,
                                  command=self.code_display.yview)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        x_scrollbar = ttk.Scrollbar(code_frame,
                                  orient=tk.HORIZONTAL,
                                  command=self.code_display.xview)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.code_display.configure(yscrollcommand=y_scrollbar.set,
                                  xscrollcommand=x_scrollbar.set)

        # Status bar
        status_bar = ttk.Label(main_frame,
                             textvariable=self.status_var,
                             style="TRON.TLabel",
                             relief=tk.SUNKEN,
                             anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)

        # Configure syntax highlighting
        self.setup_syntax_highlighting()

        # Bind code completion events
        self.code_display.bind('<KeyRelease>', self.check_for_completion)
        self.code_display.bind('<Tab>', self.accept_completion)

    def analyze_code(self):
        """Analyze code for potential issues and improvements"""
        try:
            code = self.code_display.get("1.0", tk.END).strip()
            if not code:
                self.show_notification("No code to analyze!")
                return

            # Create analysis window if it doesn't exist
            if not self.error_analysis_window or not self.error_analysis_window.winfo_exists():
                self.error_analysis_window = tk.Toplevel(self.window)
                self.error_analysis_window.title("Code Analysis")
                self.error_analysis_window.geometry("600x400")
                self.error_analysis_window.configure(bg="#0C141F")

                # Analysis text area
                analysis_text = tk.Text(self.error_analysis_window,
                                      wrap=tk.WORD,
                                      bg="#1A2030",
                                      fg="#00F0FF",
                                      font=('Share Tech Mono', 12))
                analysis_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

                # Scrollbar
                scrollbar = ttk.Scrollbar(self.error_analysis_window,
                                        orient=tk.VERTICAL,
                                        command=analysis_text.yview)
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                analysis_text.configure(yscrollcommand=scrollbar.set)

                # Close button
                close_btn = tk.Button(self.error_analysis_window,
                                    text="Close",
                                    command=self.error_analysis_window.destroy,
                                    bg="#0C141F",
                                    fg="#00F0FF",
                                    font=('Orbitron', 10))
                close_btn.pack(pady=10)

            # Get analysis from Gemini
            prompt = f"""Analyze this Python code and provide:
            1. Potential bugs and issues
            2. Performance improvements
            3. Code style suggestions
            4. Security considerations
            5. Best practices recommendations

            Code:
            {code}

            Provide a detailed analysis with specific examples and fixes."""

            response = self.chat.send_message(prompt)
            analysis = response.text.strip()

            # Update analysis window
            analysis_text.delete("1.0", tk.END)
            analysis_text.insert("1.0", analysis)
            self.error_analysis_window.lift()

        except Exception as e:
            self.show_notification(f"Error analyzing code: {str(e)}")
            logging.error(f"Error in analyze_code: {str(e)}")

    def get_code_suggestions(self):
        """Get AI-powered code suggestions"""
        try:
            code = self.code_display.get("1.0", tk.END).strip()
            if not code:
                self.show_notification("No code to get suggestions for!")
                return

            # Create suggestions window if it doesn't exist
            if not self.suggestion_window or not self.suggestion_window.winfo_exists():
                self.suggestion_window = tk.Toplevel(self.window)
                self.suggestion_window.title("Code Suggestions")
                self.suggestion_window.geometry("600x400")
                self.suggestion_window.configure(bg="#0C141F")

                # Suggestions text area
                suggestions_text = tk.Text(self.suggestion_window,
                                         wrap=tk.WORD,
                                         bg="#1A2030",
                                         fg="#00F0FF",
                                         font=('Share Tech Mono', 12))
                suggestions_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

                # Scrollbar
                scrollbar = ttk.Scrollbar(self.suggestion_window,
                                        orient=tk.VERTICAL,
                                        command=suggestions_text.yview)
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                suggestions_text.configure(yscrollcommand=scrollbar.set)

                # Apply suggestion button
                apply_btn = tk.Button(self.suggestion_window,
                                    text="Apply Selected Suggestion",
                                    command=lambda: self.apply_suggestion(suggestions_text),
                                    bg="#0C141F",
                                    fg="#00F0FF",
                                    font=('Orbitron', 10))
                apply_btn.pack(pady=10)

                # Close button
                close_btn = tk.Button(self.suggestion_window,
                                    text="Close",
                                    command=self.suggestion_window.destroy,
                                    bg="#0C141F",
                                    fg="#00F0FF",
                                    font=('Orbitron', 10))
                close_btn.pack(pady=10)

            # Get suggestions from Gemini
            prompt = f"""Provide code suggestions for this Python code:
            1. Alternative implementations
            2. Modern Python features that could be used
            3. Optimization opportunities
            4. Design pattern applications
            5. Error handling improvements

            Code:
            {code}

            Format each suggestion with a clear title and explanation."""

            response = self.chat.send_message(prompt)
            suggestions = response.text.strip()

            # Update suggestions window
            suggestions_text.delete("1.0", tk.END)
            suggestions_text.insert("1.0", suggestions)
            self.suggestion_window.lift()

        except Exception as e:
            self.show_notification(f"Error getting suggestions: {str(e)}")
            logging.error(f"Error in get_code_suggestions: {str(e)}")

    def apply_suggestion(self, suggestions_text):
        """Apply the selected suggestion to the code"""
        try:
            # Get selected text
            selected_text = suggestions_text.get("sel.first", "sel.last")
            if not selected_text:
                self.show_notification("Please select a suggestion to apply!")
                return

            # Extract code from suggestion
            code_match = re.search(r'```python\n(.*?)\n```', selected_text, re.DOTALL)
            if code_match:
                code_to_apply = code_match.group(1).strip()
                
                # Insert at cursor position
                cursor_pos = self.code_display.index(tk.INSERT)
                self.code_display.insert(cursor_pos, code_to_apply)
                
                self.show_notification("Suggestion applied successfully!")
            else:
                self.show_notification("No valid code found in selected suggestion!")

        except Exception as e:
            self.show_notification(f"Error applying suggestion: {str(e)}")
            logging.error(f"Error in apply_suggestion: {str(e)}")

    def check_for_completion(self, event):
        """Check if code completion should be triggered"""
        try:
            if event.keysym in ('Return', 'Tab', 'space'):
                # Get current line and cursor position
                current_line = self.code_display.get("insert linestart", "insert")
                cursor_pos = self.code_display.index(tk.INSERT)
                
                # Extract the word being typed
                words = current_line.split()
                if words:
                    current_word = words[-1]
                    
                    # Only trigger completion for words longer than 2 characters
                    if len(current_word) > 2:
                        self.get_code_completion(current_word, cursor_pos)
                        
        except Exception as e:
            logging.error(f"Error in check_for_completion: {str(e)}")

    def get_code_completion(self, partial_word, cursor_pos):
        """Get code completion suggestions from Gemini"""
        try:
            if self.completion_thread and self.completion_thread.is_alive():
                return

            def completion_task():
                try:
                    # Get context around cursor
                    context = self.code_display.get("insert-5c linestart", "insert+5c lineend")
                    
                    # Get completion from Gemini
                    prompt = f"""Provide code completion suggestions for this Python code context.
                    Current word: {partial_word}
                    
                    Context:
                    {context}
                    
                    Provide only the most relevant completion suggestions."""

                    response = self.chat.send_message(prompt)
                    suggestions = response.text.strip().split('\n')
                    
                    # Filter and format suggestions
                    valid_suggestions = [s.strip() for s in suggestions if s.strip()]
                    
                    if valid_suggestions:
                        # Show completion popup
                        self.show_completion_popup(valid_suggestions, cursor_pos)
                        
                except Exception as e:
                    logging.error(f"Error in completion task: {str(e)}")

            self.completion_thread = threading.Thread(target=completion_task)
            self.completion_thread.daemon = True
            self.completion_thread.start()

        except Exception as e:
            logging.error(f"Error in get_code_completion: {str(e)}")

    def show_completion_popup(self, suggestions, cursor_pos):
        """Show completion suggestions in a popup"""
        try:
            # Create popup window
            popup = tk.Toplevel(self.window)
            popup.overrideredirect(True)  # Remove window decorations
            
            # Position popup near cursor
            x, y = self.code_display.bbox(cursor_pos)[:2]
            popup_x = self.code_display.winfo_rootx() + x
            popup_y = self.code_display.winfo_rooty() + y + 20
            popup.geometry(f"+{popup_x}+{popup_y}")
            
            # Style popup
            popup.configure(bg="#1A2030")
            
            # Create listbox for suggestions
            listbox = tk.Listbox(popup,
                               bg="#1A2030",
                               fg="#00F0FF",
                               selectbackground="#00F0FF",
                               selectforeground="#0C141F",
                               font=('Share Tech Mono', 12),
                               relief="solid",
                               bd=1,
                               highlightthickness=1,
                               highlightcolor="#00F0FF",
                               highlightbackground="#004A5D")
            listbox.pack(fill=tk.BOTH, expand=True)
            
            # Add suggestions
            for suggestion in suggestions:
                listbox.insert(tk.END, suggestion)
            
            # Bind selection events
            def on_select(event):
                selected = listbox.get(listbox.curselection())
                self.apply_completion(selected, cursor_pos)
                popup.destroy()
            
            listbox.bind('<<ListboxSelect>>', on_select)
            
            # Bind escape to close
            popup.bind('<Escape>', lambda e: popup.destroy())
            
            # Show popup
            popup.lift()
            popup.focus_force()
            
        except Exception as e:
            logging.error(f"Error showing completion popup: {str(e)}")

    def apply_completion(self, completion, cursor_pos):
        """Apply the selected completion"""
        try:
            # Get current line and cursor position
            current_line = self.code_display.get("insert linestart", "insert")
            words = current_line.split()
            
            if words:
                # Replace the last word with completion
                current_line = ' '.join(words[:-1]) + ' ' + completion
                self.code_display.delete("insert linestart", "insert")
                self.code_display.insert("insert linestart", current_line)
                
        except Exception as e:
            logging.error(f"Error applying completion: {str(e)}")

    def accept_completion(self, event):
        """Handle tab key for completion"""
        if self.code_completion_active:
            # Accept current completion
            self.code_completion_active = False
            return "break"
        return None

    def setup_bindings(self):
        """Setup keyboard shortcuts"""
        self.window.bind('<Control-Return>', lambda e: self.generate_program())
        self.window.bind('<Control-a>', lambda e: self.analyze_code())
        self.window.bind('<Control-s>', lambda e: self.get_code_suggestions())
        self.window.bind('<Control-l>', lambda e: self.clear_all())
        self.window.bind('<Control-c>', lambda e: self.copy_to_clipboard())
        self.window.bind('<Control-Shift-S>', lambda e: self.save_code())

    def setup_syntax_highlighting(self):
        """Setup syntax highlighting tags"""
        self.code_display.tag_configure("keyword", foreground="#FF00FF")
        self.code_display.tag_configure("builtin", foreground="#00FFFF")
        self.code_display.tag_configure("string", foreground="#00FF00")
        self.code_display.tag_configure("comment", foreground="#808080")
        self.code_display.tag_configure("number", foreground="#FFA500")
        self.code_display.tag_configure("class", foreground="#FFD700")
        self.code_display.tag_configure("function", foreground="#00FF00")
        self.code_display.tag_configure("type_hint", foreground="#FF69B4")
        self.code_display.tag_configure("special", foreground="#FF4500")
        self.code_display.tag_configure("operator", foreground="#FFD700")

    def generate_program(self):
        """Generate program based on requirements"""
        try:
            requirements = self.requirements_text.get("1.0", tk.END).strip()
            if not requirements or requirements == "Example: Create a Python program that reads a CSV file, processes the data, and generates a bar chart using matplotlib.":
                self.show_notification("Please enter program requirements!")
                return

            # Disable generate button
            if "GENERATE PROGRAM" in self.buttons:
                self.buttons["GENERATE PROGRAM"].configure(state="disabled")

            # Show loading label
            loading_label = tk.Label(self.window,
                                   text="Generating program...",
                                   font=('Orbitron', 12),
                                   bg="#0C141F",
                                   fg="#00F0FF")
            loading_label.pack(pady=10)

            def generate_task():
                try:
                    # Get program from Gemini
                    prompt = f"""Generate a Python program based on these requirements:
                    {requirements}

                    Guidelines:
                    1. Use clear and concise code
                    2. Include proper error handling
                    3. Add comments explaining complex logic
                    4. Follow PEP 8 style guide
                    5. Make the code reusable
                    6. Include example usage
                    7. Add type hints where appropriate
                    8. Use modern Python features
                    9. Include proper logging
                    10. Make the code production-ready

                    Generate ONLY the Python code without any explanations or markdown formatting."""

                    response = self.chat.send_message(prompt)
                    generated_code = response.text.strip()

                    # Clean up the code
                    generated_code = re.sub(r'```python\n?', '', generated_code)
                    generated_code = re.sub(r'```\n?', '', generated_code)
                    generated_code = re.sub(r'^\s*def\s+', 'def ', generated_code)
                    generated_code = re.sub(r'^\s*class\s+', 'class ', generated_code)
                    generated_code = generated_code.split('\n\n')[0]

                    # Apply syntax highlighting
                    self.apply_syntax_highlighting(generated_code)

                    # Show success notification
                    self.show_notification("Program generated successfully! Use Ctrl+S to save or Ctrl+C to copy.")

                except Exception as e:
                    self.show_notification(f"Error generating program: {str(e)}")
                    logging.error(f"Error in generate_program: {str(e)}")
                finally:
                    # Re-enable generate button
                    if "GENERATE PROGRAM" in self.buttons:
                        self.buttons["GENERATE PROGRAM"].configure(state="normal")
                    # Safely destroy loading label
                    if loading_label and loading_label.winfo_exists():
                        loading_label.destroy()

            # Run generation in a separate thread
            threading.Thread(target=generate_task, daemon=True).start()

        except Exception as e:
            self.show_notification(f"Error generating program: {str(e)}")
            logging.error(f"Error in generate_program: {str(e)}")

    def apply_syntax_highlighting(self, code):
        """Apply syntax highlighting to the code"""
        # Clear existing text and tags
        self.code_display.delete("1.0", tk.END)
        
        # Python keywords, built-ins, and special methods
        keywords = ["def", "class", "import", "from", "return", "if", "else", "elif",
                   "try", "except", "finally", "for", "while", "in", "and", "or", 
                   "not", "is", "with", "as", "raise", "break", "continue", "pass",
                   "lambda", "assert", "del", "yield", "async", "await"]
                   
        builtins = ["True", "False", "None", "print", "len", "range", "str", "int",
                    "float", "list", "dict", "set", "tuple", "sum", "min", "max",
                    "enumerate", "zip", "map", "filter", "any", "all", "round", "abs"]
                    
        operators = ["+", "-", "*", "/", "//", "%", "**", "=", "==", "!=", "<", ">",
                    "<=", ">=", "+=", "-=", "*=", "/=", "&", "|", "^", "~", "<<", ">>"]
                    
        special_methods = [f"__{name}__" for name in ["init", "str", "repr", "len", 
                         "getitem", "setitem", "delitem", "iter", "next", "enter",
                         "exit", "call", "add", "sub", "mul", "div", "mod", "pow"]]
        
        # Split code into lines
        lines = code.split("\n")
        
        for line in lines:
            # Split line into words
            words = re.findall(r'\b\w+\b|\W+', line)
            i = 0
            
            while i < len(words):
                word = words[i]
                
                if word.isspace():
                    self.code_display.insert(tk.END, word)
                elif word in keywords:
                    self.code_display.insert(tk.END, word, "keyword")
                elif word in builtins:
                    self.code_display.insert(tk.END, word, "builtin")
                elif word in special_methods:
                    self.code_display.insert(tk.END, word, "special")
                elif word in operators:
                    self.code_display.insert(tk.END, word, "operator")
                elif word.startswith(("'", '"')):
                    self.code_display.insert(tk.END, word, "string")
                elif word.startswith("#"):
                    self.code_display.insert(tk.END, word, "comment")
                elif re.match(r'^[0-9]+(\.[0-9]+)?$', word):
                    self.code_display.insert(tk.END, word, "number")
                elif re.match(r'^[A-Z][a-zA-Z0-9]*$', word):
                    self.code_display.insert(tk.END, word, "class")
                elif re.match(r'^[a-z][a-zA-Z0-9_]*\($', word):
                    self.code_display.insert(tk.END, word, "function")
                elif re.match(r'^[A-Z][a-zA-Z0-9_]*(\[.*\])?$', word):
                    self.code_display.insert(tk.END, word, "type_hint")
                else:
                    self.code_display.insert(tk.END, word)

                i += 1
            
            self.code_display.insert(tk.END, "\n")

    def clear_example_text(self, event):
        """Clear example text when focused"""
        if self.requirements_text.get("1.0", tk.END).strip() == "Example: Create a Python program that reads a CSV file, processes the data, and generates a bar chart using matplotlib.":
            self.requirements_text.delete("1.0", tk.END)

    def on_button_hover(self, event, button):
        """Handle button hover effect"""
        button.configure(bg="#00F0FF", fg="#0C141F")

    def on_button_leave(self, event, button):
        """Handle button leave effect"""
        button.configure(bg="#0C141F", fg="#00F0FF")

    def copy_to_clipboard(self):
        """Copy code to clipboard"""
        try:
            code = self.code_display.get("1.0", tk.END).strip()
            if code:
                self.window.clipboard_clear()
                self.window.clipboard_append(code)
                self.show_notification("Code copied to clipboard!")
            else:
                self.show_notification("No code to copy!")
        except Exception as e:
            self.show_notification(f"Error copying code: {str(e)}")
            logging.error(f"Error in copy_to_clipboard: {str(e)}")

    def save_code(self):
        """Save code to file"""
        try:
            code = self.code_display.get("1.0", tk.END).strip()
            if not code:
                self.show_notification("No code to save!")
                return

            file_path = filedialog.asksaveasfilename(
                defaultextension=".py",
                filetypes=[("Python files", "*.py"), ("All files", "*.*")],
                title="Save Generated Code"
            )

            if file_path:
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(code)
                self.show_notification(f"Code saved successfully to {file_path}")

        except Exception as e:
            self.show_notification(f"Error saving code: {str(e)}")
            logging.error(f"Error in save_code: {str(e)}")

    def clear_all(self):
        """Clear all text fields"""
        self.requirements_text.delete("1.0", tk.END)
        self.code_display.delete("1.0", tk.END)
        self.status_var.set("")
        # Reset example text
        self.requirements_text.insert('1.0', "Example: Create a Python program that reads a CSV file, processes the data, and generates a bar chart using matplotlib.")

    def show_notification(self, message):
        """Show a notification message in the status bar"""
        if hasattr(self, 'status_var'):
            self.status_var.set(message)
        else:
            messagebox.showinfo("Notification", message)

    def center_window(self):
        """Center the window on the screen"""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f'{width}x{height}+{x}+{y}')

    def on_closing(self):
        """Handle window closing"""
        self.window.destroy() 