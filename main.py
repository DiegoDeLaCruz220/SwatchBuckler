from PIL import Image, ImageTk, ImageDraw, ImageEnhance
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
import os
import sys

try:
    import pytesseract
    import os
    if os.name == 'nt':  # Windows
        possible_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]
        for path in possible_paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                break
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

class ImageViewer:
    """Simple image viewer with zoom and pan."""
    
    def __init__(self, parent, image):
        self.parent = parent
        self.original_image = image
        self.img_width, self.img_height = image.size
        
        # Zoom settings
        self.zoom_level = 1.0
        self.min_zoom = 0.1
        self.max_zoom = 5.0
        
        # Pan offset
        self.offset_x = 0
        self.offset_y = 0
        
        # Display canvas size
        self.canvas_width = 1200
        self.canvas_height = 800
        
        # Calculate initial zoom to fit
        self.fit_to_window()
        
        self.current_image = None
        self.photo = None
        
    def fit_to_window(self):
        """Calculate zoom to fit entire image in window."""
        zoom_x = self.canvas_width / self.img_width
        zoom_y = self.canvas_height / self.img_height
        self.zoom_level = min(zoom_x, zoom_y, 1.0)
        
        # Center the image
        display_w = int(self.img_width * self.zoom_level)
        display_h = int(self.img_height * self.zoom_level)
        self.offset_x = (self.canvas_width - display_w) // 2
        self.offset_y = (self.canvas_height - display_h) // 2
        
    def get_display_image(self):
        """Get the image at current zoom level."""
        new_width = int(self.img_width * self.zoom_level)
        new_height = int(self.img_height * self.zoom_level)
        
        if new_width < 1 or new_height < 1:
            new_width = max(1, new_width)
            new_height = max(1, new_height)
            
        self.current_image = self.original_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        return self.current_image
    
    def screen_to_image(self, screen_x, screen_y):
        """Convert screen coordinates to original image coordinates."""
        # Remove offset
        x = screen_x - self.offset_x
        y = screen_y - self.offset_y
        
        # Scale to original
        img_x = int(x / self.zoom_level)
        img_y = int(y / self.zoom_level)
        
        return img_x, img_y
    
    def image_to_screen(self, img_x, img_y):
        """Convert image coordinates to screen coordinates."""
        screen_x = int(img_x * self.zoom_level) + self.offset_x
        screen_y = int(img_y * self.zoom_level) + self.offset_y
        return screen_x, screen_y
    
    def zoom(self, factor, mouse_x=None, mouse_y=None):
        """Zoom in or out, optionally around a point."""
        old_zoom = self.zoom_level
        self.zoom_level *= factor
        self.zoom_level = max(self.min_zoom, min(self.max_zoom, self.zoom_level))
        
        if mouse_x is not None and mouse_y is not None:
            # Zoom around mouse position
            # Get image coordinate at mouse before zoom
            img_x, img_y = self.screen_to_image(mouse_x, mouse_y)
            
            # After zoom, where would that coordinate be on screen?
            new_screen_x, new_screen_y = self.image_to_screen(img_x, img_y)
            
            # Adjust offset to keep mouse over same image point
            self.offset_x += (mouse_x - new_screen_x)
            self.offset_y += (mouse_y - new_screen_y)
    
    def pan(self, dx, dy):
        """Pan the image."""
        self.offset_x += dx
        self.offset_y += dy


class SwatchExtractor:
    def __init__(self, root, image_path):
        self.root = root
        self.root.title("Color Swatch Extractor")
        
        # Load image
        self.original_image = Image.open(image_path)
        self.viewer = ImageViewer(root, self.original_image)
        
        # State
        self.selection_enabled = False
        self.texture_mode = False  # For textured swatches
        self.text_offset_from_color_x1 = None
        self.text_offset_from_color_y1 = None
        self.text_width = None
        self.text_height = None
        self.learning_text_position = False
        self.drawing_text_box = False
        self.text_box_start = None
        self.first_color_bounds = None
        self.last_swatch_bounds = None
        self.last_color_x = None
        self.last_color_y = None
        self.rectangles = []
        self.extracted_count = 0
        
        # Pan state
        self.panning = False
        self.pan_start = None
        
        # Output
        self.output_dir = "color_swatches"
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
        
        # Create UI
        self.create_ui()
        
        # Initial display
        self.update_canvas()
    
    def create_ui(self):
        # Make window fullscreen or large and centered
        self.root.state('zoomed')  # Maximize window on Windows
        
        # Use PanedWindow for resizable divider
        paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True)
        
        # Canvas frame (left side)
        canvas_frame = tk.Frame(paned_window)
        paned_window.add(canvas_frame, weight=3)
        
        # Canvas
        self.canvas = tk.Canvas(canvas_frame, 
                               width=self.viewer.canvas_width, 
                               height=self.viewer.canvas_height,
                               bg='gray')
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # Bind events
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.canvas.bind("<Button-3>", self.start_pan)
        self.canvas.bind("<B3-Motion>", self.do_pan)
        self.canvas.bind("<ButtonRelease-3>", self.end_pan)
        self.canvas.bind("<Motion>", self.on_motion)
        
        # Side panel (right side) - resizable
        panel = tk.Frame(paned_window, bg='#f0f0f0', width=600)
        paned_window.add(panel, weight=1)
        
        # Add inner frame with padding
        inner_panel = tk.Frame(panel, bg='#f0f0f0')
        inner_panel.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        tk.Label(inner_panel, text="Color Swatch Extractor", 
                font=("Arial", 16, "bold"), bg='#f0f0f0').pack(pady=10)
        
        # Output directory selector
        output_frame = tk.Frame(inner_panel, bg='#f0f0f0')
        output_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(output_frame, text="Save to:", font=("Arial", 10, "bold"), bg='#f0f0f0').pack(anchor=tk.W)
        
        path_row = tk.Frame(output_frame, bg='#f0f0f0')
        path_row.pack(fill=tk.X, pady=3)
        
        self.output_entry = tk.Entry(path_row, font=("Arial", 9), 
                                     relief=tk.SUNKEN, bg='white')
        self.output_entry.insert(0, self.output_dir)
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        browse_btn = tk.Button(path_row, text="Browse", 
                              command=self.change_output_dir, 
                              font=("Arial", 9),
                              width=10)
        browse_btn.pack(side=tk.LEFT)
        
        tk.Label(inner_panel, text="", height=1, bg='#f0f0f0').pack()
        
        # Selection toggle - taller button with better text
        self.selection_btn = tk.Button(inner_panel, 
                                      text="🔒 Selection: OFF\nClick to enable",
                                      command=self.toggle_selection,
                                      bg="#FF5722",
                                      fg="white",
                                      font=("Arial", 13, "bold"),
                                      height=3,
                                      wraplength=550)
        self.selection_btn.pack(fill=tk.X, pady=10)
        
        # Texture mode checkbox
        self.texture_var = tk.BooleanVar(value=False)
        texture_check = tk.Checkbutton(inner_panel, 
                                       text="Texture Mode",
                                       variable=self.texture_var,
                                       command=self.toggle_texture_mode,
                                       font=("Arial", 10),
                                       bg='#f0f0f0',
                                       activebackground='#f0f0f0')
        texture_check.pack(anchor=tk.W, pady=5)
        
        tk.Label(inner_panel, text="", height=1, bg='#f0f0f0').pack()
        
        # Instructions
        tk.Label(inner_panel, text="Instructions:", 
                font=("Arial", 12, "bold"), bg='#f0f0f0').pack(anchor=tk.W)
        instructions = [
            "1. Enable Selection mode",
            "2. Click on first color",
            "3. Drag box around its name",
            "4. Click other colors",
        ]
        for inst in instructions:
            tk.Label(inner_panel, text=inst, font=("Arial", 11), 
                    wraplength=550, justify=tk.LEFT, bg='#f0f0f0').pack(anchor=tk.W, padx=5, pady=3)
        
        tk.Label(inner_panel, text="", height=1, bg='#f0f0f0').pack()
        
        # Controls
        tk.Label(inner_panel, text="Controls:", 
                font=("Arial", 11, "bold"), bg='#f0f0f0').pack(anchor=tk.W)
        tk.Label(inner_panel, text="• Mouse wheel: Zoom", 
                font=("Arial", 10), bg='#f0f0f0').pack(anchor=tk.W, padx=5)
        tk.Label(inner_panel, text="• Right-drag: Pan", 
                font=("Arial", 10), bg='#f0f0f0').pack(anchor=tk.W, padx=5)
        tk.Label(inner_panel, text="• F: Fit to window", 
                font=("Arial", 10), bg='#f0f0f0').pack(anchor=tk.W, padx=5)
        self.root.bind("<f>", lambda e: self.fit_to_window())
        self.root.bind("<F>", lambda e: self.fit_to_window())
        
        tk.Label(inner_panel, text="", height=1, bg='#f0f0f0').pack()
        
        # Status
        self.coord_label = tk.Label(inner_panel, text="Position: ", 
                                    font=("Courier", 10), bg='#f0f0f0')
        self.coord_label.pack(anchor=tk.W)
        
        self.zoom_label = tk.Label(inner_panel, 
                                   text=f"Zoom: {self.viewer.zoom_level:.0%}", 
                                   font=("Courier", 10), bg='#f0f0f0')
        self.zoom_label.pack(anchor=tk.W)
        
        tk.Label(inner_panel, text="", height=1, bg='#f0f0f0').pack()
        
        self.mode_label = tk.Label(inner_panel, text="Mode: Navigation", 
                                   font=("Arial", 11, "bold"), fg="orange", bg='#f0f0f0')
        self.mode_label.pack(anchor=tk.W)
        
        self.status_label = tk.Label(inner_panel, text="Ready", 
                                     font=("Arial", 10), wraplength=550, 
                                     justify=tk.LEFT, bg='#f0f0f0')
        self.status_label.pack(anchor=tk.W, pady=5)
        
        self.extracted_label = tk.Label(inner_panel, text="Extracted: 0", 
                                        font=("Arial", 12, "bold"), bg='#f0f0f0')
        self.extracted_label.pack(anchor=tk.W, pady=5)
    
    def toggle_texture_mode(self):
        """Toggle texture mode for patterned swatches."""
        self.texture_mode = self.texture_var.get()
        if self.texture_mode:
            self.status_label.config(text="Texture mode: Will use edge detection for boundaries", fg="blue")
        else:
            self.status_label.config(text="Normal mode: Detecting solid color boundaries", fg="blue")
    
    def toggle_selection(self):
        self.selection_enabled = not self.selection_enabled
        if self.selection_enabled:
            self.selection_btn.config(text="✓ Selection: ON\nClick to disable", bg="#4CAF50")
            self.mode_label.config(text="Mode: Click on color", fg="green")
            self.status_label.config(text="Click on a color swatch", fg="blue")
        else:
            self.selection_btn.config(text="🔒 Selection: OFF\nClick to enable", bg="#FF5722")
            self.mode_label.config(text="Mode: Navigation", fg="orange")
            self.status_label.config(text="Zoom/pan freely", fg="orange")
    
    def change_output_dir(self):
        """Allow user to change output directory."""
        # Check if user manually edited the entry field
        current_entry_value = self.output_entry.get().strip()
        
        new_dir = filedialog.askdirectory(
            title="Select output folder for color swatches",
            initialdir=current_entry_value if os.path.exists(current_entry_value) else self.output_dir
        )
        if new_dir:
            self.output_dir = new_dir
            # Update the entry with full path
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, new_dir)
            
            # Create directory if it doesn't exist
            if not os.path.exists(self.output_dir):
                os.makedirs(self.output_dir)
            
            self.status_label.config(text=f"Output changed to: {os.path.basename(self.output_dir)}", fg="green")
    
    def get_output_dir(self):
        """Get the current output directory from entry field."""
        # Allow user to type path directly
        path = self.output_entry.get().strip()
        if path and path != self.output_dir:
            # User manually changed the path
            self.output_dir = path
            if not os.path.exists(self.output_dir):
                os.makedirs(self.output_dir)
        return self.output_dir
    
    def fit_to_window(self):
        self.viewer.fit_to_window()
        self.update_canvas()
        self.zoom_label.config(text=f"Zoom: {self.viewer.zoom_level:.0%}")
    
    def update_canvas(self):
        """Redraw the canvas with current image and overlays."""
        # Get display image
        display_img = self.viewer.get_display_image().copy()
        draw = ImageDraw.Draw(display_img)
        
        # Draw rectangles on the image
        for rect_data in self.rectangles:
            x1, y1, x2, y2, color = rect_data
            # Convert to display coordinates (relative to displayed image)
            dx1 = int(x1 * self.viewer.zoom_level)
            dy1 = int(y1 * self.viewer.zoom_level)
            dx2 = int(x2 * self.viewer.zoom_level)
            dy2 = int(y2 * self.viewer.zoom_level)
            draw.rectangle([dx1, dy1, dx2, dy2], outline=color, width=max(2, int(2 * self.viewer.zoom_level)))
        
        self.viewer.photo = ImageTk.PhotoImage(display_img)
        
        # Clear and redraw
        self.canvas.delete("all")
        self.canvas.create_image(self.viewer.offset_x, self.viewer.offset_y, 
                                anchor=tk.NW, image=self.viewer.photo)
    
    def on_mousewheel(self, event):
        if event.delta > 0:
            self.viewer.zoom(1.2, event.x, event.y)
        else:
            self.viewer.zoom(0.8, event.x, event.y)
        self.update_canvas()
        self.zoom_label.config(text=f"Zoom: {self.viewer.zoom_level:.0%}")
    
    def start_pan(self, event):
        self.panning = True
        self.pan_start = (event.x, event.y)
        self.canvas.config(cursor="fleur")
    
    def do_pan(self, event):
        if self.panning and self.pan_start:
            dx = event.x - self.pan_start[0]
            dy = event.y - self.pan_start[1]
            self.viewer.pan(dx, dy)
            self.pan_start = (event.x, event.y)
            self.update_canvas()
    
    def end_pan(self, event):
        self.panning = False
        self.pan_start = None
        self.canvas.config(cursor="")
    
    def on_motion(self, event):
        img_x, img_y = self.viewer.screen_to_image(event.x, event.y)
        if 0 <= img_x < self.viewer.img_width and 0 <= img_y < self.viewer.img_height:
            self.coord_label.config(text=f"Position: ({img_x}, {img_y})")
    
    def on_click(self, event):
        if self.panning or not self.selection_enabled:
            return
        
        if self.learning_text_position:
            self.drawing_text_box = True
            self.text_box_start = (event.x, event.y)
            self.status_label.config(text="Drag to draw box around text...", fg="blue")
            return
        
        img_x, img_y = self.viewer.screen_to_image(event.x, event.y)
        
        # Detect color boundaries
        if self.texture_mode:
            x1, y1, x2, y2 = self.find_textured_swatch_boundaries(img_x, img_y)
        else:
            x1, y1, x2, y2 = self.find_color_boundaries(img_x, img_y)
        
        if x2 - x1 < 20 or y2 - y1 < 20:
            self.status_label.config(text="Region too small, click on color center", fg="red")
            return
        
        # Store and draw
        self.last_swatch_bounds = (x1, y1, x2, y2)
        self.last_color_x = img_x
        self.last_color_y = img_y
        
        if self.first_color_bounds is None:
            self.first_color_bounds = (x1, y1, x2, y2)
        
        self.rectangles.append((x1, y1, x2, y2, "red"))
        self.update_canvas()
        
        # Handle naming
        if HAS_OCR and self.text_offset_from_color_x1 is None:
            self.learning_text_position = True
            self.mode_label.config(text="Mode: Draw box around name", fg="orange")
            self.status_label.config(text="Now drag box around the color's name", fg="blue")
            messagebox.showinfo("Next Step", 
                              "Good! Now click and drag to draw a box\naround the NAME of this selected color.\n\n"
                              "The box should tightly contain just the text label.",
                              parent=self.root)
            return
        elif HAS_OCR and self.text_offset_from_color_x1 is not None:
            # Auto-detect name
            text_x1 = x1 + self.text_offset_from_color_x1
            text_y1 = y1 + self.text_offset_from_color_y1
            text_x2 = text_x1 + self.text_width
            text_y2 = text_y1 + self.text_height
            
            # Show preview
            self.rectangles.append((text_x1, text_y1, text_x2, text_y2, "cyan"))
            self.update_canvas()
            
            name = self.extract_text_from_box(text_x1, text_y1, text_x2, text_y2)
            
            # Remove preview
            self.rectangles.pop()
            
            if name:
                confirmed_name = simpledialog.askstring("Swatch Name", 
                                                       f"Detected: {name}\n\nEdit or press OK to accept:", 
                                                       initialvalue=name)
                if confirmed_name:
                    self.save_swatch(x1, y1, x2, y2, confirmed_name)
                    self.rectangles[-1] = (x1, y1, x2, y2, "green")
                    self.update_canvas()
                else:
                    self.rectangles.pop()
                    self.update_canvas()
            else:
                self.status_label.config(text="Couldn't read text, enter manually", fg="orange")
                name = simpledialog.askstring("Swatch Name", "Couldn't detect text.\n\nEnter color name (or cancel to skip):")
                if name:
                    self.save_swatch(x1, y1, x2, y2, name)
                    self.rectangles[-1] = (x1, y1, x2, y2, "green")
                    self.update_canvas()
                else:
                    self.rectangles.pop()
                    self.update_canvas()
        else:
            name = simpledialog.askstring("Swatch Name", "Enter color name (or cancel to skip):")
            if name:
                self.save_swatch(x1, y1, x2, y2, name)
                self.rectangles[-1] = (x1, y1, x2, y2, "green")
                self.update_canvas()
            else:
                self.rectangles.pop()
                self.update_canvas()
    
    def on_drag(self, event):
        if not self.selection_enabled or not self.drawing_text_box or not self.text_box_start:
            return
        
        # Show preview rectangle
        self.update_canvas()
        start_x, start_y = self.text_box_start
        self.canvas.create_rectangle(start_x, start_y, event.x, event.y, 
                                     outline="blue", width=2, dash=(5, 5))
    
    def on_release(self, event):
        if not self.drawing_text_box or not self.text_box_start:
            return
        
        start_x, start_y = self.text_box_start
        end_x, end_y = event.x, event.y
        
        # Convert to image coords
        img_x1, img_y1 = self.viewer.screen_to_image(start_x, start_y)
        img_x2, img_y2 = self.viewer.screen_to_image(end_x, end_y)
        
        x1, x2 = min(img_x1, img_x2), max(img_x1, img_x2)
        y1, y2 = min(img_y1, img_y2), max(img_y1, img_y2)
        
        # Calculate offsets
        color_x1, color_y1, color_x2, color_y2 = self.first_color_bounds
        self.text_offset_from_color_x1 = x1 - color_x1
        self.text_offset_from_color_y1 = y1 - color_y1
        self.text_width = x2 - x1
        self.text_height = y2 - y1
        
        self.drawing_text_box = False
        self.text_box_start = None
        self.learning_text_position = False
        self.mode_label.config(text="Mode: Auto-detect names", fg="green")
        
        text = self.extract_text_from_box(x1, y1, x2, y2)
        
        if text:
            name = simpledialog.askstring("Swatch Name", 
                                         f"Detected: {text}\n\nEdit or press OK to accept:", 
                                         initialvalue=text)
            if name:
                sx1, sy1, sx2, sy2 = self.last_swatch_bounds
                self.save_swatch(sx1, sy1, sx2, sy2, name)
                self.rectangles[-1] = (sx1, sy1, sx2, sy2, "green")
                self.status_label.config(text=f"Learned! Click other swatches", fg="green")
            else:
                self.rectangles.pop()
                self.status_label.config(text="Cancelled", fg="orange")
        else:
            name = simpledialog.askstring("Swatch Name", "Enter color name:")
            if name:
                sx1, sy1, sx2, sy2 = self.last_swatch_bounds
                self.save_swatch(sx1, sy1, sx2, sy2, name)
                self.rectangles[-1] = (sx1, sy1, sx2, sy2, "green")
                self.status_label.config(text=f"Learned! Click other swatches", fg="green")
            else:
                self.rectangles.pop()
        
        self.update_canvas()
    
    def find_color_boundaries(self, click_x, click_y, threshold=30):
        pixels = self.original_image.load()
        center_color = pixels[click_x, click_y]
        
        def color_matches(x, y):
            if x < 0 or y < 0 or x >= self.viewer.img_width or y >= self.viewer.img_height:
                return False
            pixel = pixels[x, y]
            diff = sum(abs(pixel[i] - center_color[i]) for i in range(3))
            return diff < threshold
        
        x1 = click_x
        while x1 > 0 and color_matches(x1 - 1, click_y):
            x1 -= 1
        
        x2 = click_x
        while x2 < self.viewer.img_width - 1 and color_matches(x2 + 1, click_y):
            x2 += 1
        
        y1 = click_y
        while y1 > 0 and color_matches(click_x, y1 - 1):
            y1 -= 1
        
        y2 = click_y
        while y2 < self.viewer.img_height - 1 and color_matches(click_x, y2 + 1):
            y2 += 1
        
        margin = 2
        x1 = min(x1 + margin, click_x)
        y1 = min(y1 + margin, click_y)
        x2 = max(x2 - margin, click_x)
        y2 = max(y2 - margin, click_y)
        
        return x1, y1, x2, y2
    
    def find_textured_swatch_boundaries(self, click_x, click_y):
        """Find boundaries of textured swatches using edge detection approach."""
        pixels = self.original_image.load()
        
        # Get a sample of colors around click point to understand the texture
        sample_radius = 10
        color_samples = []
        for dy in range(-sample_radius, sample_radius + 1, 3):
            for dx in range(-sample_radius, sample_radius + 1, 3):
                sx = click_x + dx
                sy = click_y + dy
                if 0 <= sx < self.viewer.img_width and 0 <= sy < self.viewer.img_height:
                    color_samples.append(pixels[sx, sy])
        
        # Calculate average color and variance for the texture
        if not color_samples:
            return self.find_color_boundaries(click_x, click_y)
        
        avg_r = sum(c[0] for c in color_samples) / len(color_samples)
        avg_g = sum(c[1] for c in color_samples) / len(color_samples)
        avg_b = sum(c[2] for c in color_samples) / len(color_samples)
        
        # Calculate standard deviation for texture
        variance = sum(
            (c[0] - avg_r)**2 + (c[1] - avg_g)**2 + (c[2] - avg_b)**2 
            for c in color_samples
        ) / len(color_samples)
        texture_threshold = max(60, min(120, variance ** 0.5 * 3))  # Adaptive threshold
        
        def texture_matches(x, y):
            """Check if pixel is part of the textured swatch."""
            if x < 0 or y < 0 or x >= self.viewer.img_width or y >= self.viewer.img_height:
                return False
            pixel = pixels[x, y]
            # Check if pixel is within the texture's color range
            diff = abs(pixel[0] - avg_r) + abs(pixel[1] - avg_g) + abs(pixel[2] - avg_b)
            return diff < texture_threshold
        
        # Find boundaries by expanding from center
        x1 = click_x
        while x1 > 0 and texture_matches(x1 - 1, click_y):
            x1 -= 1
        
        x2 = click_x
        while x2 < self.viewer.img_width - 1 and texture_matches(x2 + 1, click_y):
            x2 += 1
        
        y1 = click_y
        while y1 > 0 and texture_matches(click_x, y1 - 1):
            y1 -= 1
        
        y2 = click_y
        while y2 < self.viewer.img_height - 1 and texture_matches(click_x, y2 + 1):
            y2 += 1
        
        # Apply smaller margin for textured swatches
        margin = 3
        x1 = min(x1 + margin, click_x)
        y1 = min(y1 + margin, click_y)
        x2 = max(x2 - margin, click_x)
        y2 = max(y2 - margin, click_y)
        
        return x1, y1, x2, y2
    
    def extract_text_from_box(self, x1, y1, x2, y2):
        if not HAS_OCR:
            return None
        
        x1, x2 = min(x1, x2), max(x1, x2)
        y1, y2 = min(y1, y2), max(y1, y2)
        x1 = max(0, x1)
        y1 = max(0, y1)
        x2 = min(self.viewer.img_width, x2)
        y2 = min(self.viewer.img_height, y2)
        
        if x2 - x1 < 10 or y2 - y1 < 10:
            return None
        
        text_region = self.original_image.crop((x1, y1, x2, y2))
        
        # Enhance for OCR
        text_region = text_region.convert('L')
        enhancer = ImageEnhance.Contrast(text_region)
        text_region = enhancer.enhance(2.0)
        enhancer = ImageEnhance.Sharpness(text_region)
        text_region = enhancer.enhance(2.0)
        text_region = text_region.resize((text_region.width * 3, text_region.height * 3), Image.Resampling.LANCZOS)
        
        try:
            configs = ['--psm 7 --oem 3', '--psm 8 --oem 3', '--psm 13 --oem 3']
            for config in configs:
                text = pytesseract.image_to_string(text_region, config=config).strip()
                text = ''.join(c for c in text if c.isalnum() or c in ' _-')
                text = text.strip().replace(' ', '_').lower()
                if text and len(text) > 2:
                    return text
            return None
        except Exception as e:
            print(f"OCR error: {e}")
            return None
    
    def save_swatch(self, x1, y1, x2, y2, name):
        # Get current output directory (in case user typed a new path)
        output_dir = self.get_output_dir()
        
        swatch = self.original_image.crop((x1, y1, x2, y2))
        filename = f"{name}.png"
        filepath = os.path.join(output_dir, filename)
        swatch.save(filepath)
        
        self.extracted_count += 1
        self.status_label.config(text=f"Saved: {filename}", fg="green")
        self.extracted_label.config(text=f"Extracted: {self.extracted_count}")


def main():
    import sys
    
    root = tk.Tk()
    
    # Set window icon early
    try:
        # Handle PyInstaller bundled resources
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            base_path = sys._MEIPASS
        else:
            # Running as script
            base_path = os.path.dirname(os.path.abspath(__file__))
        
        # Try .ico file first (works best on Windows)
        ico_path = os.path.join(base_path, 'logo.ico')
        if os.path.exists(ico_path):
            root.wm_iconbitmap(ico_path)
        else:
            # Fallback to .png with iconphoto
            icon_path = os.path.join(base_path, 'logo.png')
            if os.path.exists(icon_path):
                icon = tk.PhotoImage(file=icon_path)
                root.iconphoto(True, icon)
    except Exception as e:
        print(f"Could not load icon: {e}")
    
    root.withdraw()
    
    file_dialog = tk.Toplevel()
    
    # Set icon on file dialog too
    try:
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        ico_path = os.path.join(base_path, 'logo.ico')
        if os.path.exists(ico_path):
            file_dialog.iconbitmap(ico_path)
    except Exception as e:
        print(f"Could not load icon for dialog: {e}")
    
    file_dialog.title("Color Swatch Extractor - Select Image")
    file_dialog.geometry("600x200")
    file_dialog.resizable(False, False)
    
    file_dialog.update_idletasks()
    x = (file_dialog.winfo_screenwidth() // 2) - 300
    y = (file_dialog.winfo_screenheight() // 2) - 100
    file_dialog.geometry(f'+{x}+{y}')
    
    selected_file = {"path": None}
    
    def on_browse():
        filepath = filedialog.askopenfilename(
            title="Select image file",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.bmp *.tiff *.gif"),
                ("PNG files", "*.png"),
                ("All files", "*.*")
            ]
        )
        if filepath:
            path_entry.delete(0, tk.END)
            path_entry.insert(0, filepath)
    
    def on_start():
        filepath = path_entry.get().strip().strip('"')
        if filepath and os.path.exists(filepath):
            selected_file["path"] = filepath
            file_dialog.destroy()
        else:
            messagebox.showerror("Error", "Please select a valid image file.", parent=file_dialog)
    
    def on_cancel():
        file_dialog.destroy()
        root.destroy()
    
    tk.Label(file_dialog, text="Color Swatch Extractor", font=("Arial", 16, "bold")).pack(pady=15)
    
    frame = tk.Frame(file_dialog)
    frame.pack(pady=10, padx=20, fill=tk.X)
    
    tk.Label(frame, text="Image File:", font=("Arial", 10)).pack(anchor=tk.W)
    
    path_frame = tk.Frame(frame)
    path_frame.pack(fill=tk.X, pady=5)
    
    path_entry = tk.Entry(path_frame, font=("Arial", 10), width=50)
    path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
    
    if len(sys.argv) > 1:
        path_entry.insert(0, sys.argv[1].strip('"'))
    
    browse_btn = tk.Button(path_frame, text="Browse...", command=on_browse, width=10)
    browse_btn.pack(side=tk.LEFT)
    
    button_frame = tk.Frame(file_dialog)
    button_frame.pack(pady=15)
    
    start_btn = tk.Button(button_frame, text="Start", command=on_start, width=15, bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
    start_btn.pack(side=tk.LEFT, padx=5)
    
    cancel_btn = tk.Button(button_frame, text="Cancel", command=on_cancel, width=15)
    cancel_btn.pack(side=tk.LEFT, padx=5)
    
    path_entry.bind("<Return>", lambda e: on_start())
    file_dialog.protocol("WM_DELETE_WINDOW", on_cancel)
    file_dialog.wait_window()
    
    if selected_file["path"]:
        root.deiconify()
        try:
            app = SwatchExtractor(root, selected_file["path"])
            root.mainloop()
            print(f"\n✓ Extracted {app.extracted_count} swatches to: {app.output_dir}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load image:\n{e}")
            print(f"Error: {e}")
    else:
        root.destroy()

if __name__ == "__main__":
    main()
