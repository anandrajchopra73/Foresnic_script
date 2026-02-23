import os
import shutil
import sqlite3
import tempfile
import platform
import json
import base64
import stat
from datetime import datetime, timedelta
from getpass import getuser
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry 
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.dates as mdates
from matplotlib.widgets import Cursor
import numpy as np
from urllib.parse import urlparse, parse_qsl, urlunparse, urlencode
import threading
import queue

try:
    from PIL import Image, ImageTk, ImageDraw, ImageFont
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False

try:
    from Crypto.Cipher import AES
    from Crypto.Protocol.KDF import PBKDF2
    from Crypto.Util.Padding import unpad
    CRYPTO_AVAILABLE = True
except Exception:
    CRYPTO_AVAILABLE = False

try:
    import win32crypt
    WIN32CRYPT_AVAILABLE = True
except Exception:
    WIN32CRYPT_AVAILABLE = False

# optional: requests for link checking
try:
    import requests
    REQUESTS_AVAILABLE = True
except Exception:
    REQUESTS_AVAILABLE = False

import ctypes
from ctypes import wintypes

# Tooltip class for showing full URLs on hover
class ToolTip:
    def __init__(self, widget, text='widget info'):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(500, self.showtip)  # Show after 500ms

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def showtip(self, event=None):
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        
        # Create tooltip window
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry("+%d+%d" % (x, y))
        
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                        background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                        font=("tahoma", "8", "normal"), wraplength=400)
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

    def update_text(self, new_text):
        self.text = new_text

# -----------------------------
# CONFIG
# -----------------------------
BROWSERS = {
    "Chrome": {
        "win": f"C:/Users/{getuser()}/AppData/Local/Google/Chrome/User Data/Default",
        "mac": f"/Users/{getuser()}/Library/Application Support/Google/Chrome/Default",
        "linux": f"/home/{getuser()}/.config/google-chrome/Default"
    },
    "Opera": {
        "win": f"C:/Users/{getuser()}/AppData/Roaming/Opera Software/Opera Stable",
        "mac": f"/Users/{getuser()}/Library/Application Support/com.operasoftware.Opera",
        "linux": f"/home/{getuser()}/.config/opera"
    },
    "Opera GX": {
        "win": f"C:/Users/{getuser()}/AppData/Roaming/Opera Software/Opera GX Stable",
        "mac": f"/Users/{getuser()}/Library/Application Support/com.operasoftware.OperaGX",
        "linux": f"/home/{getuser()}/.config/opera-gx"
    },
    "Brave": {
        "win": f"C:/Users/{getuser()}/AppData/Local/BraveSoftware/Brave-Browser/User Data/Default",
        "mac": f"/Users/{getuser()}/Library/Application Support/BraveSoftware/Brave-Browser/Default",
        "linux": f"/home/{getuser()}/.config/BraveSoftware/Brave-Browser/Default"
    },
    "Edge": {
        "win": f"C:/Users/{getuser()}/AppData/Local/Microsoft/Edge/User Data/Default",
        "mac": f"/Users/{getuser()}/Library/Application Support/Microsoft Edge/Default",
        "linux": f"/home/{getuser()}/.config/microsoft-edge/Default"
    },
    "Firefox": {
        "win": f"C:/Users/{getuser()}/AppData/Roaming/Mozilla/Firefox/Profiles",
        "mac": f"/Users/{getuser()}/Library/Application Support/Firefox/Profiles",
        "linux": f"/home/{getuser()}/.mozilla/firefox"
    }
}

POSSIBLE_COOKIE_DB_NAMES = ["Cookies", os.path.join("Network", "Cookies")]
POSSIBLE_LOGIN_DB_NAMES = ["Login Data", "Login Data For Account"]

# -----------------------------
# Helpers
# -----------------------------

def get_platform():
    pf = platform.system()
    if pf == "Darwin":
        return "mac"
    elif pf == "Linux":
        return "linux"
    else:
        return "win"


def copy_db_file(db_path):
    temp_dir = tempfile.mkdtemp()
    filename = os.path.basename(db_path)
    temp_db = os.path.join(temp_dir, filename)
    try:
        shutil.copy2(db_path, temp_db)
    except Exception:
        shutil.copyfile(db_path, temp_db)
    try:
        os.chmod(temp_db, stat.S_IWRITE | stat.S_IREAD)
    except Exception:
        pass
    return temp_db


def _safe_exists(path):
    try:
        return os.path.exists(path)
    except Exception:
        return False


def discover_opera_profile_path():
    """Try hard to locate an Opera profile directory that contains a Chromium 'History' DB."""
    user_home = os.path.expanduser("~")
    candidates = [
        os.path.join(user_home, "AppData", "Roaming", "Opera Software", "Opera Stable"),
        os.path.join(user_home, "AppData", "Local", "Opera Software", "Opera Stable"),
        # Opera GX common locations
        os.path.join(user_home, "AppData", "Roaming", "Opera Software", "Opera GX Stable"),
        os.path.join(user_home, "AppData", "Local", "Opera Software", "Opera GX Stable"),
        os.path.join(user_home, "AppData", "Roaming", "Opera Software"),
        os.path.join(user_home, "AppData", "Local", "Opera Software"),
        os.path.join(user_home, "OneDrive", "AppData", "Roaming", "Opera Software", "Opera Stable"),
        os.path.join(user_home, "OneDrive", "AppData", "Local", "Opera Software", "Opera Stable"),
        os.path.join(user_home, "OneDrive", "AppData", "Roaming", "Opera Software", "Opera GX Stable"),
        os.path.join(user_home, "OneDrive", "AppData", "Local", "Opera Software", "Opera GX Stable"),
    ]
    # Also look directly under Opera Software for any sub-profiles
    for root in list(candidates):
        if _safe_exists(root) and os.path.isdir(root):
            history_path = os.path.join(root, "History")
            if _safe_exists(history_path):
                return root
            try:
                for name in os.listdir(root):
                    cand = os.path.join(root, name)
                    if os.path.isdir(cand) and _safe_exists(os.path.join(cand, "History")):
                        return cand
            except Exception:
                pass
    # As a last resort, do a limited-depth scan from user_home and Desktop
    limited_roots = [user_home, os.path.join(user_home, "Desktop")]
    max_depth = 4
    try:
        for base in limited_roots:
            if not _safe_exists(base):
                continue
            queue = [(base, 0)]
            seen = set()
            while queue:
                current, depth = queue.pop(0)
                if current in seen or depth > max_depth:
                    continue
                seen.add(current)
                name_lower = os.path.basename(current).lower()
                if current.lower().endswith("opera stable") or current.lower().endswith("opera gx stable") or "opera" in name_lower:
                    if _safe_exists(os.path.join(current, "History")):
                        return current
                try:
                    for name in os.listdir(current):
                        child = os.path.join(current, name)
                        if os.path.isdir(child):
                            queue.append((child, depth + 1))
                except Exception:
                    pass
    except Exception:
        pass
    return None


def chrome_time_to_datetime(chrome_time):
    if chrome_time:
        try:
            # Chrome stores times as microseconds since Jan 1, 1601
            return datetime(1601, 1, 1) + timedelta(microseconds=int(chrome_time))
        except Exception:
            return None
    return None


def datetime_to_ddmmyy(dt):
    if not dt:
        return ""
    return dt.strftime("%d/%m/%y %H:%M:%S")


def date_only_ddmmyy(dt):
    if not dt:
        return ""
    return dt.strftime("%d/%m/%y")


def month_name(dt):
    if not dt:
        return ""
    return dt.strftime("%B")


def get_date_range_from_chrome_history(db_path):
    try:
        temp_db = copy_db_file(db_path)
        conn = sqlite3.connect(temp_db)
        cursor = conn.cursor()
        cursor.execute("SELECT MIN(last_visit_time), MAX(last_visit_time) FROM urls")
        row = cursor.fetchone()
        conn.close()
        if row:
            min_time, max_time = row
            return chrome_time_to_datetime(min_time), chrome_time_to_datetime(max_time)
        return None, None
    except Exception:
        return None, None

# -----------------------------
# Firefox helpers
# -----------------------------

def firefox_time_to_datetime(firefox_time):
    """Convert Firefox timestamp to datetime"""
    if firefox_time:
        try:
            # Firefox stores times as microseconds since Unix epoch
            return datetime.fromtimestamp(firefox_time / 1000000)
        except Exception:
            return None
    return None

def get_firefox_profiles(base_path):
    """Get list of Firefox profile directories"""
    profiles = []
    if not os.path.exists(base_path):
        return profiles
    
    try:
        # Look for profiles.ini file
        profiles_ini = os.path.join(base_path, "profiles.ini")
        if os.path.exists(profiles_ini):
            import configparser
            config = configparser.ConfigParser()
            config.read(profiles_ini, encoding='utf-8')
            
            for section in config.sections():
                if section.startswith('Profile'):
                    if config.has_option(section, 'Path'):
                        profile_path = config[section]['Path']
                        if config.has_option(section, 'IsRelative') and config[section]['IsRelative'] == '1':
                            profile_path = os.path.join(base_path, profile_path)
                        else:
                            profile_path = os.path.join(base_path, profile_path)
                        
                        if os.path.exists(profile_path):
                            profiles.append(profile_path)
        
        # Always try fallback method to find more profiles
        try:
            for item in os.listdir(base_path):
                item_path = os.path.join(base_path, item)
                if os.path.isdir(item_path):
                    # Look for Firefox profile indicators
                    if (item.endswith('.default') or 
                        item.endswith('.default-release') or 
                        item.endswith('.default-esr') or
                        'default' in item.lower() or
                        os.path.exists(os.path.join(item_path, 'places.sqlite'))):
                        if item_path not in profiles:
                            profiles.append(item_path)
        except Exception:
            pass
            
    except Exception:
        # If all else fails, try to find any directory with places.sqlite
        try:
            for item in os.listdir(base_path):
                item_path = os.path.join(base_path, item)
                if os.path.isdir(item_path) and os.path.exists(os.path.join(item_path, 'places.sqlite')):
                    profiles.append(item_path)
        except Exception:
            pass
    
    return profiles

def get_date_range_from_firefox_history(profile_path):
    """Get date range from Firefox history"""
    try:
        places_path = os.path.join(profile_path, "places.sqlite")
        if not os.path.exists(places_path):
            return None, None
            
        temp_db = copy_db_file(places_path)
        conn = sqlite3.connect(temp_db)
        cursor = conn.cursor()
        cursor.execute("SELECT MIN(last_visit_date), MAX(last_visit_date) FROM moz_places WHERE last_visit_date IS NOT NULL")
        row = cursor.fetchone()
        conn.close()
        if row:
            min_time, max_time = row
            return firefox_time_to_datetime(min_time), firefox_time_to_datetime(max_time)
        return None, None
    except Exception:
        return None, None

# Enhanced Graph Class with Perfect X-axis and Hover Details
class EnhancedGraph:
    def __init__(self, parent_frame):
        self.parent_frame = parent_frame
        self.fig, self.ax = plt.subplots(figsize=(12, 6))
        self.canvas = FigureCanvasTkAgg(self.fig, parent_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Data storage for hover functionality
        self.data_points = []
        self.hover_annotation = None
        self.line_objects = []  # Store line objects for hover detection
        
        # Connect mouse events
        self.canvas.mpl_connect('motion_notify_event', self.on_hover)
        self.canvas.mpl_connect('button_press_event', self.on_click)
        
    def _bytes_to_readable(self, num_bytes):
        """Convert bytes to human readable format"""
        try:
            num = float(num_bytes)
        except Exception:
            return ""
        if num < 1024:
            return f"{int(num)} B"
        for unit in ["KB", "MB", "GB", "TB"]:
            num /= 1024.0
            if num < 1024.0:
                if unit == "KB":
                    return f"{num:.1f} {unit}"
                else:
                    return f"{num:.2f} {unit}"
        return f"{num:.2f} PB"
    
    def plot_daily_activity(self, history_data, title="Browser Activity", context="Browsing History"):      
        """Plot daily browser activity with perfect formatting and hover"""
        self.current_context = context
        self.ax.clear()
        
        if not history_data:
            self.ax.text(0.5, 0.5, 'कोई डेटा उपलब्ध नहीं है\nNo data available', 
                        horizontalalignment='center', verticalalignment='center',
                        transform=self.ax.transAxes, fontsize=14, 
                        bbox=dict(boxstyle="round,pad=0.5", facecolor='lightgray', alpha=0.7))
            self.canvas.draw()
            return
        
        # Group data by date with complete information
        daily_counts = {}
        self.data_points = []  # Reset data points
        
        for item in history_data:
            if 'visit_time' in item and item['visit_time']:
                date = item['visit_time'].date()
                if date not in daily_counts:
                    daily_counts[date] = {'count': 0, 'urls': [], 'browsers': set(), 'times': [], 'total_size': 0}
                daily_counts[date]['count'] += 1
                url_info = {
                    'url': item.get('url', 'Unknown'),
                    'title': item.get('title', 'No Title'),
                    'time': item['visit_time'].strftime('%H:%M:%S'),
                    'full_time': item['visit_time'].strftime('%d/%m/%Y %H:%M:%S')
                }
                # Add size info for downloads
                if 'size_bytes' in item and item['size_bytes']:
                    url_info['size_bytes'] = item['size_bytes']
                    daily_counts[date]['total_size'] += item['size_bytes']
                daily_counts[date]['urls'].append(url_info)
                # Extract browser info if available
                if 'browser' in item:
                    daily_counts[date]['browsers'].add(item['browser'])
                # Track times for the day (for hover first/last time)
                daily_counts[date]['times'].append(item['visit_time'])
        
        if not daily_counts:
            self.ax.text(0.5, 0.5, 'कोई वैध डेटा नहीं मिला\nNo valid date data found', 
                        horizontalalignment='center', verticalalignment='center',
                        transform=self.ax.transAxes, fontsize=14,
                        bbox=dict(boxstyle="round,pad=0.5", facecolor='lightcoral', alpha=0.7))
            self.canvas.draw()
            return
        
        # Sort dates and fill missing dates with zero counts for continuous display
        if daily_counts:
            min_date = min(daily_counts.keys())
            max_date = max(daily_counts.keys())
            
            # Create complete date range
            all_dates = []
            current_date = min_date
            while current_date <= max_date:
                all_dates.append(current_date)
                current_date += timedelta(days=1)
            
            # Fill data for all dates
            dates = all_dates
            counts = []
            for date in dates:
                if date in daily_counts:
                    counts.append(daily_counts[date]['count'])
                else:
                    counts.append(0)
                    # Add empty data for missing dates
                    daily_counts[date] = {'count': 0, 'urls': [], 'browsers': set(), 'times': [], 'total_size': 0}
        else:
            dates = []
            counts = []
        
        # Store enhanced data for hover functionality
        for i, date in enumerate(dates):
            self.data_points.append({
                'x': mdates.date2num(date),
                'y': counts[i],
                'date': date,
                'count': counts[i],
                'details': daily_counts[date]['urls'],
                'browsers': list(daily_counts[date]['browsers']),
                'day_name': date.strftime('%A'),  # Day name in English
                'formatted_date': date.strftime('%d/%m/%Y'),
                'first_time': min(daily_counts[date]['times']).strftime('%H:%M:%S') if daily_counts[date]['times'] else None,
                'last_time': max(daily_counts[date]['times']).strftime('%H:%M:%S') if daily_counts[date]['times'] else None
            })
        
        # Create the enhanced plot with better styling
        line = self.ax.plot(dates, counts, marker='o', linewidth=3, markersize=8, 
                           color='#2E86AB', markerfacecolor='#A23B72', 
                           markeredgecolor='white', markeredgewidth=2,
                           label='Daily Activity')
        self.line_objects = line
        
        # Dynamic timestamp formatting based on data range
        total_days = len(dates)
        date_range = (dates[-1] - dates[0]).days if len(dates) > 1 else 1
        
        if date_range > 1095:  # More than 3 years
            self.ax.xaxis.set_major_locator(mdates.YearLocator())
            self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
            self.ax.xaxis.set_minor_locator(mdates.MonthLocator(interval=6))
            rotation = 0
        elif date_range > 730:  # 2-3 years
            self.ax.xaxis.set_major_locator(mdates.MonthLocator(interval=6))
            self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))
            self.ax.xaxis.set_minor_locator(mdates.MonthLocator(interval=3))
            rotation = 45
        elif date_range > 365:  # 1-2 years
            self.ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
            self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %y'))
            self.ax.xaxis.set_minor_locator(mdates.MonthLocator(interval=1))
            rotation = 45
        elif date_range > 180:  # 6 months - 1 year
            self.ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
            self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %y'))
            self.ax.xaxis.set_minor_locator(mdates.WeekdayLocator(interval=2))
            rotation = 45
        elif date_range > 90:  # 3-6 months
            self.ax.xaxis.set_major_locator(mdates.MonthLocator(interval=1))
            self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %y'))
            self.ax.xaxis.set_minor_locator(mdates.WeekdayLocator(interval=1))
            rotation = 45
        elif date_range > 30:  # 1-3 months
            self.ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
            self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m'))
            self.ax.xaxis.set_minor_locator(mdates.DayLocator(interval=2))
            rotation = 45
        elif date_range > 14:  # 2 weeks - 1 month
            self.ax.xaxis.set_major_locator(mdates.DayLocator(interval=3))
            self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m'))
            self.ax.xaxis.set_minor_locator(mdates.DayLocator(interval=1))
            rotation = 45
        else:  # Less than 2 weeks - SHOW EVERY DAY
            self.ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
            self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m'))
            self.ax.xaxis.set_minor_locator(mdates.HourLocator(interval=12))
            rotation = 90
        
        # Apply rotation and styling
        plt.setp(self.ax.xaxis.get_majorticklabels(), rotation=rotation, ha='right', fontsize=9, fontweight='bold')
        
        # Adjust margins based on rotation
        bottom_margin = 0.2 if rotation == 90 else 0.15
        self.fig.subplots_adjust(bottom=bottom_margin)
        
        # Set X-axis limits to touch left and right edges
        if dates:
            start_date = dates[0]
            end_date = dates[-1]
            self.ax.set_xlim(start_date, end_date)
        
        # Add minor ticks for better precision - every 6 hours
        self.ax.xaxis.set_minor_locator(mdates.HourLocator(interval=6))
        
        # Style the ticks for better visibility
        self.ax.tick_params(axis='x', which='major', length=8, width=2, color='#2c3e50')
        self.ax.tick_params(axis='x', which='minor', length=4, width=1, color='#7f8c8d')
        
        # Add grid lines for each day
        self.ax.grid(True, which='major', axis='x', alpha=0.6, linestyle='-', linewidth=1)
        self.ax.grid(True, which='minor', axis='x', alpha=0.2, linestyle=':', linewidth=0.5)
        
        # Enhanced plot styling
        self.ax.set_title(f'{title}', fontsize=16, fontweight='bold', pad=20, color='#2c3e50')
        self.ax.set_xlabel('Date', fontsize=12, fontweight='bold', color='#34495e')
        self.ax.set_ylabel('Visit Count', fontsize=12, fontweight='bold', color='#34495e')
        
        # Enhanced grid
        self.ax.grid(True, alpha=0.3, linestyle='--', linewidth=0.8)
        self.ax.grid(True, which='minor', alpha=0.1, linestyle=':')
        
        # Professional background
        self.ax.set_facecolor('#f8f9fa')
        self.fig.patch.set_facecolor('white')
        
        # Add value labels on points for better visibility
        for i, (date, count) in enumerate(zip(dates, counts)):
            if count > 0:  # Only show labels for non-zero values
                self.ax.annotate(str(count), (date, count), 
                               textcoords="offset points", xytext=(0,10), 
                               ha='center', fontsize=9, fontweight='bold',
                               color='#2c3e50', alpha=0.8)
        
        # Perfect layout adjustment
        self.fig.tight_layout(pad=2.0)
        
        # Create enhanced hover annotation
        self.hover_annotation = self.ax.annotate('', xy=(0,0), xytext=(30,30), 
                                               textcoords="offset points",
                                               bbox=dict(boxstyle="round,pad=0.8", 
                                                        facecolor='#fff3cd', 
                                                        edgecolor='#856404',
                                                        linewidth=2, alpha=0.95),
                                               arrowprops=dict(arrowstyle="->", 
                                                             connectionstyle="arc3,rad=0.2",
                                                             color='#856404', lw=2),
                                               fontsize=10, visible=False,
                                               fontfamily='monospace')
        
        # Legend removed as requested
        
        self.canvas.draw()
    
    def on_hover(self, event):
        """Enhanced hover with complete information display"""
        if event.inaxes != self.ax or not self.data_points:
            if self.hover_annotation:
                self.hover_annotation.set_visible(False)
                self.canvas.draw_idle()
            return
        
        # Find closest data point with better detection
        closest_point = None
        min_distance = float('inf')
        
        for point in self.data_points:
            # Calculate both X and Y distance for better detection
            x_dist = abs(event.xdata - point['x']) if event.xdata else float('inf')
            y_dist = abs(event.ydata - point['y']) if event.ydata else float('inf')
            
            # Normalize distances
            x_range = max([p['x'] for p in self.data_points]) - min([p['x'] for p in self.data_points])
            y_range = max([p['y'] for p in self.data_points]) - min([p['y'] for p in self.data_points])
            
            if x_range > 0 and y_range > 0:
                normalized_dist = (x_dist/x_range)**2 + (y_dist/y_range)**2
            else:
                normalized_dist = x_dist
            
            if normalized_dist < min_distance:
                min_distance = normalized_dist
                closest_point = point
        
        # Show enhanced details with better threshold
        if closest_point and min_distance < 0.01:  # Adjusted threshold
            # Context-specific tooltip display
            details_text = f"Date: {closest_point['formatted_date']}\n"
            
            # Check if this is Downloads context for special formatting
            if hasattr(self, 'current_context') and self.current_context == 'Downloads':
                details_text += f"Total Downloads: {closest_point['count']}\n"
                # Calculate total size if available
                total_size = 0
                size_count = 0
                for detail in closest_point.get('details', []):
                    if 'size_bytes' in detail and detail['size_bytes']:
                        total_size += detail['size_bytes']
                        size_count += 1
                if total_size > 0:
                    details_text += f"Total Size: {self._bytes_to_readable(total_size)}\n"
            else:
                details_text += f"Total Visits: {closest_point['count']}\n"
            
            if closest_point['browsers']:
                details_text += f"Browsers: {', '.join(closest_point['browsers'])}"
            
            # Update annotation with enhanced positioning
            self.hover_annotation.xy = (closest_point['x'], closest_point['y'])
            self.hover_annotation.set_text(details_text)
            self.hover_annotation.set_visible(True)
            
            # Smart positioning - show left for right-side points
            if event.x > self.ax.bbox.width * 0.5:  # Right half of graph
                self.hover_annotation.xytext = (-180, 30)  # Show on left
            else:  # Left half of graph
                self.hover_annotation.xytext = (30, 30)  # Show on right
                
        else:
            self.hover_annotation.set_visible(False)
        
        self.canvas.draw_idle()
    
    def on_click(self, event):
        """Handle click events for detailed information"""
        if event.inaxes != self.ax or not self.data_points:
            return
            
        # Find clicked point
        for point in self.data_points:
            x_dist = abs(event.xdata - point['x']) if event.xdata else float('inf')
            y_dist = abs(event.ydata - point['y']) if event.ydata else float('inf')
            
            if x_dist < 1 and y_dist < point['count'] * 0.1:
                # Show detailed popup or print to console
                print(f"\n📊 Detailed Information for {point['formatted_date']}:")
                print(f"Total Activities: {point['count']}")
                print(f"Day: {point['day_name']}")
                if point['browsers']:
                    print(f"Browsers: {', '.join(point['browsers'])}")
                print("\nAll Activities:")
                for i, activity in enumerate(point['details'], 1):
                    print(f"{i:2d}. {activity['full_time']} - {activity['title'][:50]}")
                    print(f"     URL: {activity['url'][:80]}")
                break
    
    def clear_plot(self):
        """Clear the current plot"""
        self.ax.clear()
        self.data_points = []
        self.hover_annotation = None
        self.canvas.draw()

# -----------------------------
# Bytes human-readable
# -----------------------------

def bytes_to_readable(num_bytes):
    try:
        num = float(num_bytes)
    except Exception:
        return ""
    if num < 1024:
        return f"{int(num)} B"
    for unit in ["KB", "MB", "GB", "TB"]:
        num /= 1024.0
        if num < 1024.0:
            if unit == "KB":
                return f"{num:.1f} {unit}"
            else:
                return f"{num:.2f} {unit}"
    return f"{num:.2f} PB"

# -----------------------------
# Cookie decryption helpers (unchanged)
# -----------------------------

def _read_local_state_key(base_path):
    parent = os.path.dirname(base_path)
    # Try typical Chromium parent and Opera's profile-local placement
    candidate_paths = [
        os.path.join(parent, "Local State"),
        os.path.join(base_path, "Local State"),
    ]
    local_state_path = None
    for cand in candidate_paths:
        if os.path.exists(cand):
            local_state_path = cand
            break
    if not local_state_path:
        return None
    try:
        with open(local_state_path, "r", encoding="utf-8") as f:
            local = json.load(f)
        enc_key_b64 = local.get("os_crypt", {}).get("encrypted_key")
        if not enc_key_b64:
            return None
        enc_key = base64.b64decode(enc_key_b64)
        if enc_key.startswith(b"DPAPI"):
            enc_key = enc_key[5:]
        try:
            if get_platform() == "win":
                if WIN32CRYPT_AVAILABLE:
                    return win32crypt.CryptUnprotectData(enc_key, None, None, None, 0)[1]
                else:
                    return _crypt_unprotect_data_ctypes(enc_key)
            elif get_platform() == "mac":
                return enc_key
            else:
                return enc_key
        except Exception:
            return enc_key
    except Exception:
        return None


def _crypt_unprotect_data_ctypes(encrypted_bytes):
    try:
        crypt32 = ctypes.windll.crypt32
        kernel32 = ctypes.windll.kernel32

        class DATA_BLOB(ctypes.Structure):
            _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

        def _to_blob(data: bytes):
            blob = DATA_BLOB()
            blob.cbData = len(data)
            blob.pbData = ctypes.cast(ctypes.create_string_buffer(data, len(data)), ctypes.POINTER(ctypes.c_char))
            return blob

        in_blob = _to_blob(encrypted_bytes)
        out_blob = DATA_BLOB()
        if crypt32.CryptUnprotectData(ctypes.byref(in_blob), None, None, None, None, 0, ctypes.byref(out_blob)):
            ptr = out_blob.pbData
            length = int(out_blob.cbData)
            buf = ctypes.string_at(ptr, length)
            kernel32.LocalFree(out_blob.pbData)
            return buf
    except Exception:
        pass
    raise RuntimeError("DPAPI decrypt failed")


def _decrypt_aes_gcm(encrypted_value, key):
    if not CRYPTO_AVAILABLE:
        raise RuntimeError("pycryptodome required for AES-GCM decryption (pip install pycryptodome)")

    try:
        if isinstance(encrypted_value, (bytes, bytearray)) and (encrypted_value.startswith(b'v10') or encrypted_value.startswith(b'v11')):
            iv = encrypted_value[3:15]
            ciphertext = encrypted_value[15:]
            try:
                tag = ciphertext[-16:]
                ct = ciphertext[:-16]
                cipher = AES.new(key, AES.MODE_GCM, nonce=iv)
                plaintext = cipher.decrypt_and_verify(ct, tag)
                return plaintext.decode("utf-8", errors="ignore")
            except Exception:
                try:
                    return ciphertext.decode("utf-8", errors="ignore")
                except Exception:
                    return ""
        else:
            if not CRYPTO_AVAILABLE:
                raise RuntimeError("pycryptodome required for AES-CBC fallback decryption")
            salt = b"saltysalt"
            iterations = 1003
            length = 16
            derived = PBKDF2("peanuts", salt, dkLen=length, count=iterations)
            iv = b" " * 16
            cipher = AES.new(derived, AES.MODE_CBC, iv)
            try:
                decrypted = unpad(cipher.decrypt(encrypted_value), 16)
                return decrypted.decode("utf-8", errors="ignore")
            except Exception:
                try:
                    return encrypted_value.decode("utf-8", errors="ignore")
                except Exception:
                    return ""
    except Exception:
        raise


def _get_key_via_keychain(service_names=("Chrome Safe Storage", "Chromium Safe Storage", "Brave Safe Storage", "Microsoft Edge Safe Storage")):
    if get_platform() != "mac":
        return None
    for svc in service_names:
        try:
            import subprocess
            p = subprocess.Popen(["security", "find-generic-password", "-s", svc, "-w"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            out, err = p.communicate()
            if p.returncode == 0:
                password = out.strip()
                if not password:
                    continue
                if not CRYPTO_AVAILABLE:
                    return None
                key = PBKDF2(password, b"saltysalt", dkLen=16, count=1003)
                return key
        except Exception:
            continue
    return None


def identify_encryption_type(encrypted_value):
    """Identify the type of encryption used on the password"""
    if not encrypted_value:
        return "No encryption (empty)"
    
    if isinstance(encrypted_value, str):
        return "Plain text"
    
    if isinstance(encrypted_value, (bytes, bytearray, memoryview)):
        enc_bytes = bytes(encrypted_value)
        
        # Check for Chrome v10/v11 encryption
        if enc_bytes.startswith(b'v10') or enc_bytes.startswith(b'v11'):
            return "Chrome AES-GCM (v10/v11)"
        
        # Check for DPAPI (Windows)
        if get_platform() == "win":
            try:
                if WIN32CRYPT_AVAILABLE:
                    win32crypt.CryptUnprotectData(enc_bytes, None, None, None, 0)
                    return "Windows DPAPI"
                else:
                    _crypt_unprotect_data_ctypes(enc_bytes)
                    return "Windows DPAPI (ctypes)"
            except Exception:
                pass
        
        # Check for Chrome legacy encryption (Linux/Mac)
        if len(enc_bytes) >= 16:
            return "Chrome AES-CBC (legacy)"
        
        return f"Unknown binary format ({len(enc_bytes)} bytes)"
    
    return f"Unknown type: {type(encrypted_value)}"

def decrypt_chrome_password(encrypted_value, base_path=None):
    """Decrypt Chrome saved password with detailed error handling"""
    if not encrypted_value:
        return "", "No password data"
    
    if isinstance(encrypted_value, str):
        return encrypted_value, "Already decrypted"
    
    encryption_type = identify_encryption_type(encrypted_value)
    
    try:
        if get_platform() == "win":
            # Try DPAPI first
            try:
                if WIN32CRYPT_AVAILABLE:
                    decrypted = win32crypt.CryptUnprotectData(encrypted_value, None, None, None, 0)[1]
                    if decrypted:
                        return decrypted.decode("utf-8", errors="ignore"), "Windows DPAPI"
                else:
                    dec = _crypt_unprotect_data_ctypes(encrypted_value)
                    if dec:
                        return dec.decode("utf-8", errors="ignore"), "Windows DPAPI (ctypes)"
            except Exception as e:
                pass
            
            # Try AES-GCM with local state key
            key = _read_local_state_key(base_path) if base_path else None
            if key:
                try:
                    decrypted = _decrypt_aes_gcm(encrypted_value, key)
                    return decrypted, "Chrome AES-GCM"
                except Exception as e:
                    pass
        
        elif get_platform() == "mac":
            # Try keychain key first
            try:
                kc_key = _get_key_via_keychain()
                if kc_key and CRYPTO_AVAILABLE:
                    decrypted = _decrypt_aes_gcm(encrypted_value, kc_key)
                    return decrypted, "Mac Keychain AES-GCM"
            except Exception as e:
                pass
            
            # Try local state key
            if base_path:
                key = _read_local_state_key(base_path)
                if key and CRYPTO_AVAILABLE:
                    try:
                        decrypted = _decrypt_aes_gcm(encrypted_value, key)
                        return decrypted, "Chrome AES-GCM"
                    except Exception as e:
                        pass
        
        else:  # Linux
            # Try local state key first
            key = _read_local_state_key(base_path) if base_path else None
            if key and CRYPTO_AVAILABLE:
                try:
                    decrypted = _decrypt_aes_gcm(encrypted_value, key)
                    return decrypted, "Chrome AES-GCM"
                except Exception as e:
                    pass
            
            # Try legacy method
            if CRYPTO_AVAILABLE:
                try:
                    salt = b"saltysalt"
                    iterations = 1_003
                    derived = PBKDF2("peanuts", salt, dkLen=16, count=iterations)
                    iv = b" " * 16
                    cipher = AES.new(derived, AES.MODE_CBC, iv)
                    decrypted = unpad(cipher.decrypt(encrypted_value), 16)
                    return decrypted.decode("utf-8", errors="ignore"), "Chrome AES-CBC (legacy)"
                except Exception as e:
                    pass
        
        # If all decryption methods fail, return raw string representation
        try:
            return encrypted_value.decode("utf-8", errors="ignore"), f"Failed to decrypt ({encryption_type})"
        except Exception:
            return repr(encrypted_value), f"Failed to decrypt ({encryption_type})"
    
    except Exception as e:
        return f"Error: {str(e)}", encryption_type

def decrypt_chrome_cookie_value(encrypted_value, base_path=None):
    if not encrypted_value:
        return ""

    if isinstance(encrypted_value, str):
        return encrypted_value

    try:
        if get_platform() == "win":
            try:
                if WIN32CRYPT_AVAILABLE:
                    decrypted = win32crypt.CryptUnprotectData(encrypted_value, None, None, None, 0)[1]
                    if decrypted:
                        try:
                            return decrypted.decode("utf-8", errors="ignore")
                        except Exception:
                            return str(decrypted)
                else:
                    try:
                        dec = _crypt_unprotect_data_ctypes(encrypted_value)
                        if dec:
                            return dec.decode("utf-8", errors="ignore")
                    except Exception:
                        pass
            except Exception:
                pass

            key = _read_local_state_key(base_path) if base_path else None
            if key:
                try:
                    return _decrypt_aes_gcm(encrypted_value, key)
                except Exception:
                    pass

            try:
                return encrypted_value.decode("utf-8", errors="ignore")
            except Exception:
                return repr(encrypted_value)

        elif get_platform() == "mac":
            try:
                kc_key = _get_key_via_keychain()
                if kc_key and CRYPTO_AVAILABLE:
                    try:
                        return _decrypt_aes_gcm(encrypted_value, kc_key)
                    except Exception:
                        pass
            except Exception:
                pass

            if base_path:
                key = _read_local_state_key(base_path)
                if key and CRYPTO_AVAILABLE:
                    try:
                        return _decrypt_aes_gcm(encrypted_value, key)
                    except Exception:
                        pass

            try:
                return encrypted_value.decode("utf-8", errors="ignore")
            except Exception:
                return repr(encrypted_value)

        else:
            key = _read_local_state_key(base_path) if base_path else None
            if key and CRYPTO_AVAILABLE:
                try:
                    return _decrypt_aes_gcm(encrypted_value, key)
                except Exception:
                    pass

            if CRYPTO_AVAILABLE:
                try:
                    salt = b"saltysalt"
                    iterations = 1_003
                    derived = PBKDF2("peanuts", salt, dkLen=16, count=iterations)
                    iv = b" " * 16
                    cipher = AES.new(derived, AES.MODE_CBC, iv)
                    decrypted = unpad(cipher.decrypt(encrypted_value), 16)
                    return decrypted.decode("utf-8", errors="ignore")
                except Exception:
                    pass

            try:
                return encrypted_value.decode("utf-8", errors="ignore")
            except Exception:
                return repr(encrypted_value)
    except Exception:
        try:
            return encrypted_value.decode("utf-8", errors="ignore")
        except Exception:
            return str(encrypted_value)

# -----------------------------
# URL helpers for Bookmarks features
# -----------------------------

def normalize_url(u: str) -> str:
    if not u:
        return ""
    try:
        p = urlparse(u.strip())
        scheme = p.scheme.lower() or "http"
        netloc = (p.netloc or "").lower()
        path = p.path or "/"
        # drop common tracking params
        q = [(k, v) for k, v in parse_qsl(p.query, keep_blank_values=True) if not k.lower().startswith("utm_") and k.lower() not in {"gclid","fbclid"}]
        query = urlencode(q)
        # strip trailing slash (except root)
        if path != "/":
            path = path.rstrip("/")
        return urlunparse((scheme, netloc, path, "", query, ""))
    except Exception:
        return u.strip()


def extract_domain(u: str) -> str:
    try:
        return (urlparse(u).hostname or "").lower()
    except Exception:
        return ""

def shorten_url(url: str, max_length: int = 50) -> str:
    """Shorten URL for display while preserving full URL for hover/tooltip"""
    if not url:
        return ""
    
    # If URL is already short enough, return as is
    if len(url) <= max_length:
        return url
    
    try:
        parsed = urlparse(url)
        domain = parsed.netloc or ""
        path = parsed.path or ""
        
        # If domain itself is too long, truncate it
        if len(domain) > max_length - 10:
            domain = domain[:max_length-13] + "..."
        
        # Calculate remaining space for path
        remaining = max_length - len(domain) - 10  # Reserve space for protocol and ellipsis
        
        if remaining > 5 and path:
            if len(path) > remaining:
                path = path[:remaining-3] + "..."
            short_url = f"{parsed.scheme}://{domain}{path}"
        else:
            short_url = f"{parsed.scheme}://{domain}"
        
        return short_url
    except Exception:
        # Fallback: simple truncation
        if len(url) > max_length:
            return url[:max_length-3] + "..."
        return url

def shorten_download_url(url: str, max_length: int = 30) -> str:
    """Shorten download URL for compact display"""
    if not url:
        return ""
    
    # If URL is already short enough, return as is
    if len(url) <= max_length:
        return url
    
    try:
        parsed = urlparse(url)
        domain = parsed.netloc or ""
        
        # For downloads, show domain + "..." for very short display
        if len(domain) > max_length - 5:
            return domain[:max_length-8] + "..."
        else:
            return domain + "..."
    except Exception:
        # Fallback: simple truncation
        if len(url) > max_length:
            return url[:max_length-3] + "..."
        return url

def shorten_url_for_export(url: str, export_type: str = "pdf") -> str:
    """Shorten URL specifically for PDF/HTML exports to match GUI display"""
    if not url:
        return ""
    
    # Use different lengths based on export type
    if export_type.lower() == "pdf":
        max_length = 60  # Slightly longer for PDF as it has more space
    elif export_type.lower() == "html":
        max_length = 50  # Standard length for HTML
    else:
        max_length = 50  # Default
    
    return shorten_url(url, max_length)

class URLTooltip:
    """Tooltip class for showing full URLs on hover"""
    def __init__(self, widget):
        self.widget = widget
        self.tooltip_window = None
        self.original_values = {}  # Store original full URLs
        
    def bind_tooltip(self, tree):
        """Bind tooltip events to treeview"""
        tree.bind('<Motion>', self.on_motion)
        tree.bind('<Leave>', self.hide_tooltip)
        tree.bind('<Button-1>', self.hide_tooltip)
        
    def on_motion(self, event):
        """Handle mouse motion over treeview"""
        tree = event.widget
        item = tree.identify('item', event.x, event.y)
        column = tree.identify('column', event.x, event.y)
        
        if item and column:
            # Get column index
            col_index = int(column.replace('#', '')) - 1
            values = tree.item(item, 'values')
            
            if col_index < len(values):
                cell_value = values[col_index]
                
                # Check if this is a URL column (contains http or www) or a shortened value with ellipsis
                is_url = cell_value and (cell_value.startswith('http') or 'www.' in cell_value or '.' in cell_value)
                is_shortened = cell_value and '...' in cell_value
                
                if is_url or is_shortened:
                    # Get the original full value if it was shortened
                    item_id = tree.item(item)['values']
                    full_value = self.original_values.get((item, col_index), cell_value)
                    
                    # Only show tooltip if value was shortened
                    if len(full_value) > len(cell_value) or '...' in cell_value:
                        self.show_tooltip(event, full_value)
                        return
        
        self.hide_tooltip()
    
    def show_tooltip(self, event, text):
        """Show tooltip with full URL"""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            
        self.tooltip_window = tk.Toplevel()
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
        
        # Create tooltip content
        label = tk.Label(
            self.tooltip_window,
            text=text,
            background='#ffffe0',
            foreground='#000000',
            relief='solid',
            borderwidth=1,
            font=('Arial', 9),
            wraplength=600,
            justify='left',
            padx=5,
            pady=3
        )
        label.pack()
    
    def hide_tooltip(self, event=None):
        """Hide tooltip"""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None
    
    def store_original_url(self, item, col_index, full_url):
        """Store original full URL for tooltip display"""
        self.original_values[(item, col_index)] = full_url

def remove_bullets_from_text(text: str) -> str:
    """Remove bullet points and list markers from text"""
    if not text:
        return text
    
    import re
    
    # Remove various bullet point patterns
    patterns = [
        r'^[\s]*[•·▪▫◦‣⁃]\s*',  # Unicode bullets
        r'^[\s]*[-*+]\s*',       # Dash, asterisk, plus bullets
        r'^[\s]*\d+[.)]\s*',     # Numbered lists (1. or 1))
        r'^[\s]*[a-zA-Z][.)]\s*', # Lettered lists (a. or a))
        r'^[\s]*[ivxlcdm]+[.)]\s*', # Roman numerals (i. or i))
        r'^[\s]*►\s*',           # Arrow bullets
        r'^[\s]*→\s*',           # Right arrow
        r'^[\s]*✓\s*',           # Checkmark
        r'^[\s]*○\s*',           # Circle bullet
        r'^[\s]*■\s*',           # Square bullet
        r'^[\s]*◆\s*',           # Diamond bullet
    ]
    
    # Apply each pattern to remove bullets
    cleaned_text = text
    for pattern in patterns:
        cleaned_text = re.sub(pattern, '', cleaned_text, flags=re.MULTILINE | re.IGNORECASE)
    
    # Remove extra whitespace and empty lines
    lines = [line.strip() for line in cleaned_text.split('\n') if line.strip()]
    return '\n'.join(lines)

def shorten_cookie_value(cookie_value: str, max_length: int = 50) -> str:
    """Shorten cookie values for better display in PDF reports"""
    if not cookie_value:
        return ""
    
    cookie_str = str(cookie_value).strip()
    
    # If already short enough, return as is
    if len(cookie_str) <= max_length:
        return cookie_str
    
    # For very long values, show beginning and end with ellipsis
    if len(cookie_str) > max_length:
        # Calculate how much to show from start and end
        start_length = max_length // 2 - 2
        end_length = max_length - start_length - 3  # 3 for "..."
        
        if start_length > 0 and end_length > 0:
            return f"{cookie_str[:start_length]}...{cookie_str[-end_length:]}"
        else:
            # Fallback: just truncate with ellipsis
            return cookie_str[:max_length-3] + "..."
    
    return cookie_str

# -----------------------------
# Data extraction (history + cookies + downloads + bookmarks)
# -----------------------------

def extract_chrome_history_data(base_path, start_date, end_date):
    data = {"urls": [], "searches": [], "downloads": [], "cookies": [], "bookmarks": [], "passwords": []}

    # History
    history_path = os.path.join(base_path, "History")
    if os.path.exists(history_path):
        try:
            db = copy_db_file(history_path)
            conn = sqlite3.connect(db)
            cursor = conn.cursor()
            start_micro = int((start_date - datetime(1601, 1, 1)).total_seconds() * 1_000_000)
            end_micro = int((end_date - datetime(1601, 1, 1)).total_seconds() * 1_000_000)
            cursor.execute(
                """
                SELECT url, title, last_visit_time
                FROM urls
                WHERE last_visit_time BETWEEN ? AND ?
                ORDER BY last_visit_time DESC
                """,
                (start_micro, end_micro),
            )
            for url, title, ts in cursor.fetchall():
                ts_dt = chrome_time_to_datetime(ts)
                if url and ("search" in url or "google.com/search" in url or "bing.com/search" in url):
                    data["searches"].append((title or "(no title)", url, ts_dt))
                else:
                    data["urls"].append((title or "(no title)", url, ts_dt))
            conn.close()
        except Exception:
            pass

    # Downloads
    if os.path.exists(history_path):
        try:
            db = copy_db_file(history_path)
            conn = sqlite3.connect(db)
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(downloads)")
            cols = [r[1] for r in cursor.fetchall()]
            select_cols = ["current_path", "tab_url", "start_time"]
            if "total_bytes" in cols:
                select_cols.append("total_bytes")
            if "received_bytes" in cols:
                select_cols.append("received_bytes")

            q = f"SELECT {', '.join(select_cols)} FROM downloads"
            try:
                cursor.execute(q)
                fetch = cursor.fetchall()
                for row in fetch:
                    rowd = dict(zip(select_cols, row))
                    path = rowd.get("current_path")
                    source = rowd.get("tab_url") or ""
                    ts = rowd.get("start_time")
                    ts_dt = chrome_time_to_datetime(ts)
                    size_bytes = None
                    try:
                        if path and isinstance(path, str) and os.path.exists(path):
                            size_bytes = os.path.getsize(path)
                    except Exception:
                        size_bytes = None
                    if size_bytes is None:
                        if "total_bytes" in rowd and rowd.get("total_bytes") not in (None, 0):
                            try:
                                size_bytes = int(rowd.get("total_bytes") or 0)
                            except Exception:
                                size_bytes = None
                    if size_bytes is None:
                        if "received_bytes" in rowd and rowd.get("received_bytes") not in (None, 0):
                            try:
                                size_bytes = int(rowd.get("received_bytes") or 0)
                            except Exception:
                                size_bytes = None

                    data["downloads"].append((os.path.basename(path or ""), source or "", ts_dt, size_bytes))
            except sqlite3.OperationalError:
                try:
                    cursor.execute("SELECT current_path, tab_url, start_time FROM downloads")
                    for path, source, ts in cursor.fetchall():
                        ts_dt = chrome_time_to_datetime(ts)
                        size_bytes = None
                        try:
                            if path and isinstance(path, str) and os.path.exists(path):
                                size_bytes = os.path.getsize(path)
                        except Exception:
                            size_bytes = None
                        data["downloads"].append((os.path.basename(path or ""), source or "", ts_dt, size_bytes))
                except Exception:
                    pass
            conn.close()
        except Exception:
            pass

    # Cookies
    for cookie_db_name in POSSIBLE_COOKIE_DB_NAMES:
        cookie_path = os.path.join(base_path, cookie_db_name)
        if os.path.exists(cookie_path):
            try:
                db = copy_db_file(cookie_path)
                conn = sqlite3.connect(db)
                cursor = conn.cursor()
                try:
                    cursor.execute("PRAGMA table_info(cookies)")
                    cols = [r[1] for r in cursor.fetchall()]
                    use_encrypted = "encrypted_value" in cols
                    if use_encrypted:
                        cursor.execute("SELECT host_key, name, value, encrypted_value, expires_utc FROM cookies")
                        rows = cursor.fetchall()
                        for host, name, val, enc_val, exp in rows:
                            try:
                                if enc_val and isinstance(enc_val, (bytes, memoryview)):
                                    enc_bytes = bytes(enc_val)
                                    dec = decrypt_chrome_cookie_value(enc_bytes, base_path)
                                    cookie_val = dec if dec else (val or "")
                                else:
                                    cookie_val = val or ""
                            except Exception:
                                cookie_val = val or ""
                            exp_dt = chrome_time_to_datetime(exp) if exp else None
                            data["cookies"].append((host or "", name or "", cookie_val or "", exp_dt))
                    else:
                        cursor.execute("SELECT host_key, name, value, expires_utc FROM cookies")
                        for host, name, val, exp in cursor.fetchall():
                            exp_dt = chrome_time_to_datetime(exp) if exp else None
                            data["cookies"].append((host or "", name or "", val or "", exp_dt))
                except sqlite3.OperationalError:
                    try:
                        cursor.execute("SELECT host_key, name, value FROM cookies")
                        for host, name, val in cursor.fetchall():
                            data["cookies"].append((host or "", name or "", val or "", None))
                    except Exception:
                        pass
                conn.close()
            except Exception:
                pass
            break

    # Bookmarks - FIXED: Bookmarks live directly inside base_path
    bookmarks_path = os.path.join(base_path, "Bookmarks")
    if os.path.exists(bookmarks_path):
        try:
            with open(bookmarks_path, "r", encoding="utf-8") as f:
                bm = json.load(f)

            def walk_nodes(node, folder_path=""):
                if isinstance(node, dict):
                    if node.get("type") == "url":
                        name = node.get("name") or "(no name)"
                        url = node.get("url") or ""
                        date_added = node.get("date_added")
                        dt = None
                        try:
                            if date_added:
                                dt = chrome_time_to_datetime(int(date_added))
                        except Exception:
                            dt = None
                        data["bookmarks"].append((folder_path, name, url, dt))
                    elif node.get("type") == "folder":
                        new_folder = f"{folder_path}/{node.get('name', '')}" if folder_path else node.get('name', '')
                        for c in node.get("children", []) or []:
                            walk_nodes(c, new_folder)
                    else:
                        for c in node.get("children", []) or []:
                            walk_nodes(c, folder_path)
                elif isinstance(node, list):
                    for c in node:
                        walk_nodes(c, folder_path)

            roots = bm.get("roots", {})
            for root_name, root_node in roots.items():
                walk_nodes(root_node, root_name)
        except Exception:
            pass

    # Passwords/Login Data
    for login_db_name in POSSIBLE_LOGIN_DB_NAMES:
        login_path = os.path.join(base_path, login_db_name)
        if os.path.exists(login_path):
            try:
                db = copy_db_file(login_path)
                conn = sqlite3.connect(db)
                cursor = conn.cursor()
                try:
                    # Check table structure
                    cursor.execute("PRAGMA table_info(logins)")
                    cols = [r[1] for r in cursor.fetchall()]
                    
                    # Build query based on available columns
                    select_cols = ["origin_url", "username_value"]
                    if "password_value" in cols:
                        select_cols.append("password_value")
                    if "date_created" in cols:
                        select_cols.append("date_created")
                    if "date_last_used" in cols:
                        select_cols.append("date_last_used")
                    
                    query = f"SELECT {', '.join(select_cols)} FROM logins"
                    cursor.execute(query)
                    rows = cursor.fetchall()
                    
                    for row in rows:
                        row_dict = dict(zip(select_cols, row))
                        origin_url = row_dict.get("origin_url", "")
                        username = row_dict.get("username_value", "")
                        encrypted_password = row_dict.get("password_value", b"")
                        date_created = row_dict.get("date_created")
                        date_last_used = row_dict.get("date_last_used")
                        
                        # Convert timestamps
                        created_dt = chrome_time_to_datetime(date_created) if date_created else None
                        used_dt = chrome_time_to_datetime(date_last_used) if date_last_used else None
                        
                        # Decrypt password
                        if encrypted_password:
                            decrypted_password, encryption_info = decrypt_chrome_password(encrypted_password, base_path)
                        else:
                            decrypted_password = ""
                            encryption_info = "No password data"
                        
                        # Store password data
                        data["passwords"].append({
                            "url": origin_url,
                            "username": username,
                            "password_decrypted": decrypted_password,
                            "password_encrypted": encrypted_password,
                            "encryption_type": encryption_info,
                            "date_created": created_dt,
                            "date_last_used": used_dt,
                            "raw_password_string": repr(encrypted_password) if encrypted_password else ""
                        })
                        
                except sqlite3.OperationalError as e:
                    # Fallback for different table structures
                    try:
                        cursor.execute("SELECT origin_url, username_value, password_value FROM logins")
                        for origin_url, username, encrypted_password in cursor.fetchall():
                            if encrypted_password:
                                decrypted_password, encryption_info = decrypt_chrome_password(encrypted_password, base_path)
                            else:
                                decrypted_password = ""
                                encryption_info = "No password data"
                            
                            data["passwords"].append({
                                "url": origin_url or "",
                                "username": username or "",
                                "password_decrypted": decrypted_password,
                                "password_encrypted": encrypted_password,
                                "encryption_type": encryption_info,
                                "date_created": None,
                                "date_last_used": None,
                                "raw_password_string": repr(encrypted_password) if encrypted_password else ""
                            })
                    except Exception:
                        pass
                conn.close()
            except Exception:
                pass
            break  # Only process first found login database

    return data

# -----------------------------
# Firefox data extraction
# -----------------------------

def extract_firefox_history_data(base_path, start_date, end_date):
    data = {"urls": [], "searches": [], "downloads": [], "cookies": [], "bookmarks": [], "passwords": []}
    
    profiles = get_firefox_profiles(base_path)
    if not profiles:
        # If no profiles found, still return empty data structure
        return data
    
    # Convert dates to Firefox timestamp format (microseconds since epoch)
    start_micro = int(start_date.timestamp() * 1000000)
    end_micro = int(end_date.timestamp() * 1000000)
    
    for profile_path in profiles:
        # History and URLs
        places_path = os.path.join(profile_path, "places.sqlite")
        if os.path.exists(places_path):
            try:
                db = copy_db_file(places_path)
                conn = sqlite3.connect(db)
                cursor = conn.cursor()
                
                # Get URLs with visit history - more flexible query
                try:
                    cursor.execute("""
                        SELECT p.url, p.title, h.visit_date, p.visit_count
                        FROM moz_places p
                        JOIN moz_historyvisits h ON p.id = h.place_id
                        WHERE h.visit_date BETWEEN ? AND ?
                        ORDER BY h.visit_date DESC
                        LIMIT 1000
                    """, (start_micro, end_micro))
                    
                    for url, title, visit_date, visit_count in cursor.fetchall():
                        visit_dt = firefox_time_to_datetime(visit_date)
                        if url and ("search" in url or "google.com/search" in url or "bing.com/search" in url):
                            data["searches"].append((title or "(no title)", url, visit_dt))
                        else:
                            data["urls"].append((title or "(no title)", url, visit_dt))
                except Exception:
                    # Fallback: get all URLs without date filter
                    try:
                        cursor.execute("""
                            SELECT p.url, p.title, p.last_visit_date, p.visit_count
                            FROM moz_places p
                            WHERE p.last_visit_date IS NOT NULL
                            ORDER BY p.last_visit_date DESC
                            LIMIT 500
                        """)
                        
                        for url, title, visit_date, visit_count in cursor.fetchall():
                            visit_dt = firefox_time_to_datetime(visit_date)
                            if url and ("search" in url or "google.com/search" in url or "bing.com/search" in url):
                                data["searches"].append((title or "(no title)", url, visit_dt))
                            else:
                                data["urls"].append((title or "(no title)", url, visit_dt))
                    except Exception:
                        pass
                
                # Get downloads - try multiple approaches
                try:
                    cursor.execute("""
                        SELECT a.content, p.url, a.dateAdded
                        FROM moz_annos a
                        JOIN moz_places p ON a.place_id = p.id
                        WHERE a.anno_attribute_id IN (
                            SELECT id FROM moz_anno_attributes WHERE name LIKE '%download%'
                        )
                        ORDER BY a.dateAdded DESC
                        LIMIT 100
                    """)
                    
                    for content, source, date_added in cursor.fetchall():
                        date_dt = firefox_time_to_datetime(date_added)
                        data["downloads"].append((content or "Unknown", source or "", date_dt, None))
                except Exception:
                    pass
                
                # Get bookmarks - more flexible approach
                try:
                    cursor.execute("""
                        SELECT b.title, p.url, b.dateAdded, b.parent
                        FROM moz_bookmarks b
                        LEFT JOIN moz_places p ON b.fk = p.id
                        WHERE b.type = 1 AND p.url IS NOT NULL
                        ORDER BY b.dateAdded DESC
                        LIMIT 500
                    """)
                    
                    for title, url, date_added, parent in cursor.fetchall():
                        date_dt = firefox_time_to_datetime(date_added)
                        # Get folder name
                        folder = "Bookmarks"
                        if parent:
                            try:
                                cursor.execute("SELECT title FROM moz_bookmarks WHERE id = ?", (parent,))
                                parent_result = cursor.fetchone()
                                if parent_result and parent_result[0]:
                                    folder = parent_result[0]
                            except Exception:
                                folder = "Bookmarks"
                        
                        data["bookmarks"].append((folder, title or "(no title)", url or "", date_dt))
                except Exception:
                    pass
                
                conn.close()
            except Exception:
                pass
        
        # Cookies
        cookies_path = os.path.join(profile_path, "cookies.sqlite")
        if os.path.exists(cookies_path):
            try:
                db = copy_db_file(cookies_path)
                conn = sqlite3.connect(db)
                cursor = conn.cursor()
                
                try:
                    cursor.execute("""
                        SELECT host, name, value, expiry
                        FROM moz_cookies
                        ORDER BY creationTime DESC
                        LIMIT 500
                    """)
                    
                    for host, name, value, expiry in cursor.fetchall():
                        expiry_dt = None
                        if expiry:
                            try:
                                expiry_dt = datetime.fromtimestamp(expiry)
                            except Exception:
                                pass
                        data["cookies"].append((host or "", name or "", value or "", expiry_dt))
                except Exception:
                    # Fallback without expiry filter
                    try:
                        cursor.execute("SELECT host, name, value FROM moz_cookies LIMIT 500")
                        for host, name, value in cursor.fetchall():
                            data["cookies"].append((host or "", name or "", value or "", None))
                    except Exception:
                        pass
                
                conn.close()
            except Exception:
                pass
        
        # Passwords (Firefox stores passwords in logins.json)
        try:
            logins_path = os.path.join(profile_path, "logins.json")
            if os.path.exists(logins_path):
                try:
                    with open(logins_path, 'r', encoding='utf-8') as f:
                        logins_data = json.load(f)
                    
                    for login in logins_data.get('logins', []):
                        url = login.get('hostname', '')
                        username = login.get('encryptedUsername', '')
                        password = login.get('encryptedPassword', '')
                        timeCreated = login.get('timeCreated', 0)
                        timeLastUsed = login.get('timeLastUsed', 0)
                        
                        created_dt = None
                        used_dt = None
                        if timeCreated:
                            try:
                                created_dt = datetime.fromtimestamp(timeCreated / 1000)
                            except Exception:
                                pass
                        if timeLastUsed:
                            try:
                                used_dt = datetime.fromtimestamp(timeLastUsed / 1000)
                            except Exception:
                                pass
                        
                        data["passwords"].append({
                            "url": url,
                            "username": username,  # Encrypted, would need NSS decryption
                            "password_decrypted": "[Firefox Encrypted - NSS Required]",
                            "password_encrypted": password,
                            "encryption_type": "Firefox NSS (encrypted)",
                            "date_created": created_dt,
                            "date_last_used": used_dt,
                            "raw_password_string": repr(password) if password else ""
                        })
                except Exception:
                    pass
        except Exception:
            pass
    
    return data

# -----------------------------
# GUI APP (styled + header logo + enhanced Bookmarks)
# -----------------------------
class ForensicBrowserApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Browser Forensic Tool")
        self.root.geometry("1200x900")
        self.root.minsize(1000, 720)
        
        # Set window icon
        try:
            icon_path = os.path.join(os.path.dirname(__file__) if '__file__' in globals() else os.getcwd(), "ncfl.jpg")
            if os.path.exists(icon_path) and PIL_AVAILABLE:
                icon_img = Image.open(icon_path)
                icon_img = icon_img.resize((32, 32), Image.LANCZOS)
                self.icon = ImageTk.PhotoImage(icon_img)
                self.root.iconphoto(True, self.icon)
        except Exception:
            pass

        # State for bookmarks enhanced features
        self.bookmarks_master = []  # list of dicts
        
        # State for password features
        self.all_data = {}

        # Color palette
        primary_bg = "#0b4f6c"      # deep teal
        header_text = "#ffffff"
        accent = "#00b894"
        card_bg = "#f8fafc"
        subtext = "#cfeef6"

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Header.TFrame", background=primary_bg)
        style.configure("Header.TLabel", background=primary_bg, foreground=header_text, font=("Segoe UI", 18, "bold"))
        style.configure("SubHeader.TLabel", background=primary_bg, foreground=subtext, font=("Segoe UI", 10))
        style.configure("TButton", padding=6, font=("Segoe UI", 10, "bold"))
        style.map("TButton", foreground=[("active", primary_bg)], background=[("active", card_bg)])
        style.configure("Action.TButton", background=accent, foreground=primary_bg)
        style.configure("Danger.TButton", background="#e17055", foreground=header_text)
        style.configure("TNotebook", tabposition='n')
        style.configure("Treeview", font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
        
        # Custom button style for toggle
        style.configure("Toggle.TButton", 
                       background="#e74c3c", 
                       foreground="white",
                       font=("Arial", 10, "bold"))

        # Header
        header = ttk.Frame(self.root, style="Header.TFrame")
        header.pack(fill="x")

        # --- Header Layout: Use grid for left logo, center title, right logo ---
        header.columnconfigure(0, weight=1)
        header.columnconfigure(1, weight=3)
        header.columnconfigure(2, weight=1)

        # Left logo (ifso_left_img)
        self.ifso_left_img = None
        ifso_left_img_path = os.path.join(os.path.dirname(__file__) if '__file__' in globals() else os.getcwd(), "ifso.jpg")
        if os.path.exists(ifso_left_img_path) and PIL_AVAILABLE:
            try:
                im_ifso_left = Image.open(ifso_left_img_path)
                im_ifso_left = im_ifso_left.resize((120, 100), Image.LANCZOS)
                self.ifso_left_img = ImageTk.PhotoImage(im_ifso_left)
            except Exception:
                self.ifso_left_img = None
        ifso_left_label = tk.Label(header, image=self.ifso_left_img, bg=primary_bg, borderwidth=0, highlightthickness=0) if self.ifso_left_img else tk.Label(header, bg=primary_bg, borderwidth=0, highlightthickness=0)
        ifso_left_label.grid(row=0, column=0, padx=(0, 0), pady=0, sticky="nsw")

        # Center title and subtitle
        title_frame = ttk.Frame(header, style="Header.TFrame")
        title_frame.grid(row=0, column=1, padx=12, pady=12, sticky="nsew")
        title_frame.columnconfigure(0, weight=1)
        ttk.Label(title_frame, text="Browser Forensic Tool", style="Header.TLabel").grid(row=0, column=0, sticky="n", pady=(0, 2))
        self.subtitle_label = ttk.Label(title_frame, text="Cross-browser triage and investigation tool", style="SubHeader.TLabel")
        self.subtitle_label.grid(row=1, column=0, sticky="n")

        # Right logo (ifso_img)
        self.ifso_img = None
        ifso_img_path = os.path.join(os.path.dirname(__file__) if '__file__' in globals() else os.getcwd(), "ncfl.jpg")
        if os.path.exists(ifso_img_path) and PIL_AVAILABLE:
            try:
                im_ifso = Image.open(ifso_img_path)
                im_ifso = im_ifso.resize((120, 100), Image.LANCZOS)
                self.ifso_img = ImageTk.PhotoImage(im_ifso)
            except Exception:
                self.ifso_img = None
        ifso_label = tk.Label(header, image=self.ifso_img, bg=primary_bg, borderwidth=0, highlightthickness=0) if self.ifso_img else tk.Label(header, bg=primary_bg, borderwidth=0, highlightthickness=0)
        ifso_label.grid(row=0, column=2, padx=(0, 0), pady=0, sticky="nse")
        try:
            for i in range(20):
                c = int(11 + i * 4)
                c_hex = f"#{c:02x}{79:02x}{108:02x}"
                logo_holder.create_rectangle(0, i * 4, 140, (i + 1) * 4, outline=c_hex, fill=c_hex)
        except Exception:
            pass

        self.logo_img = None
        # Try multiple default logo names; user can replace any with their own file
        logo_dir = os.path.dirname(__file__) if '__file__' in globals() else os.getcwd()
        candidate_logos = [
            "browser_forensic_logo.png",
            "BrowserForensicLogo.png",
            "logo.png",
            "app_logo.png",
            "delhi_police_logo.png"
        ]
        logo_path = None
        for cand in candidate_logos:
            p = os.path.join(logo_dir, cand)
            if os.path.exists(p):
                logo_path = p
                break
        if logo_path and os.path.exists(logo_path):
            try:
                if PIL_AVAILABLE:
                    im = Image.open(logo_path)
                    im.thumbnail((120, 72), Image.LANCZOS)
                    self.logo_img = ImageTk.PhotoImage(im)
                else:
                    self.logo_img = tk.PhotoImage(file=logo_path)
            except Exception:
                self.logo_img = None
        if self.logo_img:
            logo_holder.create_image(70, 42, image=self.logo_img)

    # (Removed duplicate/old pack() calls for title_frame and its children; using grid above)

        controls = ttk.Frame(self.root)
        controls.pack(fill="x", padx=12, pady=10)

        self.os_type = get_platform()
        self.available_browsers = self.check_browsers()

        left_controls = ttk.Frame(controls)
        left_controls.pack(side="left", padx=4)

        ttk.Label(left_controls, text="Browser:").grid(row=0, column=0, padx=4)
        browser_options = ["All Browsers"] + list(self.available_browsers.keys())
        self.browser_combo = ttk.Combobox(left_controls, values=browser_options, state="readonly", width=22)
        self.browser_combo.current(0)
        self.browser_combo.grid(row=0, column=1, padx=4)

        date_frame = ttk.Frame(controls)
        date_frame.pack(side="left", padx=12)
        ttk.Label(date_frame, text="Start: DD/MM/YY").grid(row=0, column=0, padx=4)
        self.start_cal = DateEntry(date_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yy')
        self.start_cal.grid(row=0, column=1, padx=4)
        ttk.Label(date_frame, text="End: DD/MM/YY").grid(row=0, column=2, padx=4)
        self.end_cal = DateEntry(date_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yy')
        self.end_cal.grid(row=0, column=3, padx=4)
        
        # Apply button right after End Date
        ttk.Button(date_frame, text="Apply", command=self.apply_date_filter, style="TButton").grid(row=0, column=4, padx=6)
        
        # Reset button after Apply button
        ttk.Button(date_frame, text="Reset", command=self.reset_ui_defaults, style="Danger.TButton").grid(row=0, column=5, padx=6)
        
        # Search functionality - REMOVED

        ttk.Label(date_frame, text="Search:").grid(row=0, column=6, padx=(12, 4))
        self.global_search_var = tk.StringVar()
        self.search_entry = ttk.Entry(date_frame, textvariable=self.global_search_var, width=25, 
                                     font=('Arial', 9))
        self.search_entry.grid(row=0, column=7, padx=4)
        self.search_entry.bind("<KeyRelease>", lambda e: self.apply_global_search())
        # Add placeholder text
        self.search_entry.insert(0, "Search across all data...")
        self.search_entry.bind("<FocusIn>", self.on_search_focus_in)
        self.search_entry.bind("<FocusOut>", self.on_search_focus_out)
        self.search_entry.config(foreground='gray')
        ttk.Button(date_frame, text="Clear", command=self.clear_global_search, style="Action.TButton").grid(row=0, column=8, padx=4)

    # Search button removed as per user request

        right_controls = ttk.Frame(controls)
        right_controls.pack(side="right", padx=4)
        ttk.Button(right_controls, text="Save All as Single HTML", command=self.save_all_as_single_html, style="TButton").grid(row=0, column=0, padx=6)
        ttk.Button(right_controls, text="Save All as CSV", command=self.save_all_as_csv, style="TButton").grid(row=0, column=1, padx=6)
        ttk.Button(right_controls, text="Save All as PDF", command=self.save_all_as_pdf, style="TButton").grid(row=0, column=2, padx=6)
        ttk.Button(right_controls, text="i", width=3, command=self.show_credits, style="TButton").grid(row=0, column=3, padx=6)

        self.tabs = ttk.Notebook(self.root)
        self.tabs.pack(expand=1, fill="both", padx=8, pady=8)
        
        # Store original tab expand setting
        self.original_tab_expand = True

        # Overview tab
        self.overview_frame = ttk.Frame(self.tabs)
        self.tabs.add(self.overview_frame, text="Installed Browsers")
        # Overview Tree with vertical scrollbar
        _ov_container = ttk.Frame(self.overview_frame)
        _ov_container.pack(expand=1, fill="both")
        _ov_scroll = ttk.Scrollbar(_ov_container, orient="vertical")
        self.overview_tree = ttk.Treeview(_ov_container, columns=("Browser", "First Activity", "Last Activity", "History File"), show="headings", yscrollcommand=_ov_scroll.set)
        _ov_scroll.config(command=self.overview_tree.yview)
        for col, text in [
            ("Browser", "Browser"),
            ("First Activity", "Available From (DD/MM/YY HH:MM:SS)"),
            ("Last Activity", "Available To (DD/MM/YY HH:MM:SS)"),
            ("History File", "Data Source"),
        ]:
            self.overview_tree.heading(col, text=text, command=lambda c=col: self._sort_tree(self.overview_tree, c, False))
            self.overview_tree.column(col, width=160, stretch=True)
        self.overview_tree.pack(side="left", expand=1, fill="both")
        _ov_scroll.pack(side="right", fill="y")

        self.tree_views = {}
        self.url_tooltips = {}  # Store tooltip handlers for each tree
        self.download_urls = {}  # Store full URLs for downloads
        for tab_name in ["Browsing History", "Downloads", "Cookies", "Bookmarks", "Secrets"]:
            frame = ttk.Frame(self.tabs)
            self.tabs.add(frame, text=tab_name)
            if tab_name == "Cookies":
                cols = ("col1", "col2", "col3", "col4")
            elif tab_name == "Downloads":
                cols = ("col1", "col2", "col3", "col4")
            elif tab_name == "Bookmarks":
                # Enhanced columns for Bookmarks (removed Browser and Live columns)
                cols = ("col1", "col2", "col3", "col4", "col5", "col6")
            elif tab_name == "Secrets":
                # Columns for Secrets
                cols = ("col1", "col2", "col3", "col4", "col5", "col6")
            else:
                cols = ("col1", "col2", "col3")
            # Tree with vertical scrollbar per tab
            _tab_container = ttk.Frame(frame)
            _tab_container.pack(expand=1, fill="both")
            _vscroll = ttk.Scrollbar(_tab_container, orient="vertical")
            tree = ttk.Treeview(_tab_container, columns=cols, show="headings", yscrollcommand=_vscroll.set)
            _vscroll.config(command=tree.yview)

            if tab_name == "Browsing History":
                tree.heading("col1", text="Title", command=lambda c="col1": self._sort_tree(tree, c, False))
                tree.heading("col2", text="URL", command=lambda c="col2": self._sort_tree(tree, c, False))
                tree.heading("col3", text="Timestamp", command=lambda c="col3": self._sort_tree(tree, c, False))
                
                # Right-click context menu for Browsing History
                history_menu = tk.Menu(tree, tearoff=0)
                history_menu.add_command(label="Copy Title", command=lambda: self._copy_data(tree, 0, "Title"))
                history_menu.add_command(label="Copy URL", command=lambda: self._copy_data(tree, 1, "URL"))
                # Copy Selected Row and Copy All Selected Rows - REMOVED
                tree.bind("<Button-3>", lambda e, t=tree, m=history_menu: self._show_context_menu(e, t, m))
            elif tab_name == "Downloads":
                tree.heading("col1", text="File", command=lambda c="col1": self._sort_tree(tree, c, False))
                tree.heading("col2", text="Source URL", command=lambda c="col2": self._sort_tree(tree, c, False))
                tree.heading("col3", text="Size", command=lambda c="col3": self._sort_tree(tree, c, False))
                tree.heading("col4", text="Timestamp", command=lambda c="col4": self._sort_tree(tree, c, False))

                
                # Right-click context menu for Downloads
                downloads_menu = tk.Menu(tree, tearoff=0)
                downloads_menu.add_command(label="Copy File Name", command=lambda: self._copy_data(tree, 0, "File Name"))
                downloads_menu.add_command(label="Copy Source URL", command=lambda: self._copy_data(tree, 1, "Source URL"))
                downloads_menu.add_command(label="Copy Size", command=lambda: self._copy_data(tree, 2, "Size"))
                # Copy Selected Row and Copy All Selected Rows - REMOVED
                tree.bind("<Button-3>", lambda e, t=tree, m=downloads_menu: self._show_context_menu(e, t, m))
            elif tab_name == "Cookies":
                tree.heading("col1", text="Domain | Name", command=lambda c="col1": self._sort_tree(tree, c, False))
                tree.heading("col2", text="Value", command=lambda c="col2": self._sort_tree(tree, c, False))
                tree.heading("col3", text="Expires (DD/MM/YY HH:MM:SS)", command=lambda c="col3": self._sort_tree(tree, c, False))
                tree.heading("col4", text="Browser", command=lambda c="col4": self._sort_tree(tree, c, False))
                
                # Right-click context menu for Cookies
                cookies_menu = tk.Menu(tree, tearoff=0)
                cookies_menu.add_command(label="Copy Domain/Name", command=lambda: self._copy_data(tree, 0, "Domain/Name"))
                cookies_menu.add_command(label="Copy Value", command=lambda: self._copy_data(tree, 1, "Value"))
                cookies_menu.add_command(label="Copy Expires", command=lambda: self._copy_data(tree, 2, "Expires"))
                cookies_menu.add_command(label="Copy Browser", command=lambda: self._copy_data(tree, 3, "Browser"))
                # Copy Selected Row and Copy All Selected Rows - REMOVED
                tree.bind("<Button-3>", lambda e, t=tree, m=cookies_menu: self._show_context_menu(e, t, m))
            elif tab_name == "Bookmarks":
                tree.heading("col1", text="Folder", command=lambda c="col1": self._sort_tree(tree, c, False))
                tree.heading("col2", text="Name", command=lambda c="col2": self._sort_tree(tree, c, False))
                tree.heading("col3", text="URL", command=lambda c="col3": self._sort_tree(tree, c, False))
                tree.heading("col4", text="Timestamp", command=lambda c="col4": self._sort_tree(tree, c, False))
                tree.heading("col5", text="Domain", command=lambda c="col5": self._sort_tree(tree, c, False))
                tree.heading("col6", text="Browser", command=lambda c="col6": self._sort_tree(tree, c, False))



                # Right-click context menu for Bookmarks
                bookmarks_menu = tk.Menu(tree, tearoff=0)
                bookmarks_menu.add_command(label="Open in browser", command=lambda: self._bookmark_open(tree))
                bookmarks_menu.add_separator()
                bookmarks_menu.add_command(label="Copy Folder", command=lambda: self._copy_data(tree, 0, "Folder"))
                bookmarks_menu.add_command(label="Copy Name", command=lambda: self._copy_data(tree, 1, "Name"))
                bookmarks_menu.add_command(label="Copy URL", command=lambda: self._copy_data(tree, 2, "URL"))
                bookmarks_menu.add_command(label="Copy Domain", command=lambda: self._copy_data(tree, 4, "Domain"))
                bookmarks_menu.add_command(label="Copy Browser", command=lambda: self._copy_data(tree, 5, "Browser"))
                # Copy Selected Row and Copy All Selected Rows - REMOVED
                tree.bind("<Button-3>", lambda e, t=tree, m=bookmarks_menu: self._show_context_menu(e, t, m))
            
            elif tab_name == "Secrets":
                tree.heading("col1", text="Website/URL", command=lambda c="col1": self._sort_tree(tree, c, False))
                tree.heading("col2", text="Username", command=lambda c="col2": self._sort_tree(tree, c, False))
                tree.heading("col3", text="Password (Decrypted)", command=lambda c="col3": self._sort_tree(tree, c, False))
                tree.heading("col4", text="Encryption Type", command=lambda c="col4": self._sort_tree(tree, c, False))
                tree.heading("col5", text="Timestamp", command=lambda c="col5": self._sort_tree(tree, c, False))
                tree.heading("col6", text="Browser", command=lambda c="col6": self._sort_tree(tree, c, False))
                
                # Password toolbar
                pwd_tb = ttk.Frame(frame)
                pwd_tb.pack(fill="x", padx=4, pady=6)
                
                # Toggle password visibility - REMOVED
                
                # Search functionality - REMOVED
                
                # Export and Copy buttons - REMOVED
                
                # Right-click context menu for Secrets
                secrets_menu = tk.Menu(tree, tearoff=0)
                secrets_menu.add_command(label="Copy URL", command=lambda: self._copy_data(tree, 0, "URL"))
                secrets_menu.add_command(label="Copy Username", command=lambda: self._copy_data(tree, 1, "Username"))
                secrets_menu.add_command(label="Copy Password", command=lambda: self._copy_password_special(tree))
                secrets_menu.add_command(label="Copy Encryption Type", command=lambda: self._copy_data(tree, 3, "Encryption Type"))
                secrets_menu.add_command(label="Copy Date Created", command=lambda: self._copy_data(tree, 4, "Date Created"))
                secrets_menu.add_command(label="Copy Browser", command=lambda: self._copy_data(tree, 5, "Browser"))
                secrets_menu.add_separator()
                secrets_menu.add_command(label="Show Encryption Details", command=lambda: self._show_encryption_details(tree))
                # Copy Selected Row and Copy All Selected Rows - REMOVED
                tree.bind("<Button-3>", lambda e, t=tree, m=secrets_menu: self._show_context_menu(e, t, m))

            # Add keyboard shortcuts for copy functionality
            tree.bind("<Control-c>", lambda e, t=tree: self._copy_selected_row(t))
            tree.bind("<Control-C>", lambda e, t=tree: self._copy_selected_row(t))
            
            tree.pack(side="left", expand=1, fill="both")
            _vscroll.pack(side="right", fill="y")
            self.tree_views[tab_name] = tree
            
            # Add URL tooltip functionality
            tooltip = URLTooltip(tree)
            tooltip.bind_tooltip(tree)
            self.url_tooltips[tab_name] = tooltip

        # Chart container with controls
        self.chart_page_frame = ttk.Frame(self.root)
        self.chart_page_frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        
        # Chart controls frame
        chart_control_frame = ttk.LabelFrame(self.chart_page_frame, text="Graph Controls", padding=8)
        chart_control_frame.pack(fill="x", pady=(0, 5))
        
        # Left controls for graph
        left_chart_controls = ttk.Frame(chart_control_frame)
        left_chart_controls.pack(side="left", fill="x", expand=True)
        
        # Hide/Show entire graph page button
        self.graph_page_visible = True
        self.toggle_graph_page_btn = ttk.Button(
            left_chart_controls,
            text="Hide Graph Page (Ctrl+G)",
            command=self.toggle_graph_page_visibility,
            style="Toggle.TButton"
        )
        self.toggle_graph_page_btn.pack(side="left", padx=(0, 10))
        
        # Full Screen Graph button
        self.graph_fullscreen = False
        self.fullscreen_btn = ttk.Button(
            left_chart_controls,
            text="Full Screen (F11)",
            command=self.toggle_fullscreen_graph,
            style="Action.TButton"
        )
        self.fullscreen_btn.pack(side="left", padx=(0, 10))
        
        # Graph type selector
        ttk.Label(left_chart_controls, text="Graph Type:").pack(side="left", padx=(0, 5))
        self.graph_type = tk.StringVar(value="pie")
        self.graph_type_combo = ttk.Combobox(
            left_chart_controls,
            textvariable=self.graph_type,
            values=["pie", "line"],
            state="readonly",
            width=12
        )
        self.graph_type_combo.pack(side="left", padx=(0, 10))
        self.graph_type_combo.bind("<<ComboboxSelected>>", lambda e: self.redraw_chart())
        # Default tab is Installed Browsers: force pie and disable switching at start
        try:
            self.graph_type_combo.set("pie")
            self.graph_type_combo.configure(state="disabled")
        except Exception:
            pass
        
        # Export chart button - moved to right of graph type
        ttk.Button(
            left_chart_controls,
            text="Export Chart",
            command=self.export_chart
        ).pack(side="left", padx=(10, 0))
        
        # Graph header showing current context (e.g., Installed Browsers, Browsing History)
        self.chart_header_frame = ttk.Frame(self.chart_page_frame)
        self.chart_header_frame.pack(fill="x", padx=0, pady=(0, 4))
        self.current_graph_context = "Installed Browsers"
        self.graph_context_label = ttk.Label(self.chart_header_frame, text=self.current_graph_context, font=("Segoe UI", 10, "bold"))
        self.graph_context_label.pack(side="left")

        # Chart container
        self.chart_container = ttk.Frame(self.chart_page_frame)
        self.chart_container.pack(fill="both", expand=True, padx=0, pady=(5, 0))
        self.chart_canvas_widget = None
        self.current_fig = None
        
        # Initialize enhanced graph
        self.enhanced_graph = EnhancedGraph(self.chart_container)
        
        # Full screen message for tabs
        self.full_screen_message = ttk.Label(
            self.root, 
            text="Graph Page Hidden | Tabs are now in Full Screen Mode", 
            font=("Segoe UI", 12, "bold"),
            foreground="#2c3e50",
            background="#ecf0f1",
            relief="raised",
            borderwidth=2
        )
        
        # Chart style variables
        self.chart_style = {
            'background_color': '#ffffff',
            'grid_color': '#e0e0e0',
            'bar_color': '#1f77b4',
            'line_color': '#ff7f0e',
            'title_color': '#2c3e50',
            'label_color': '#34495e',
            'font_size': 12,
            'grid_alpha': 0.3
        }

        self.status_var = tk.StringVar(value="Ready")
        status = ttk.Label(self.root, textvariable=self.status_var, anchor="w")
        status.pack(fill="x", padx=8, pady=(0, 6))

        # Prefill date ranges
        self.prefill_dates()
        self.refresh_overview()

        # Initial chart setup
        self.draw_browser_data_chart({b: 0 for b in self.available_browsers.keys()}, {}, {})
        
        # Ensure graph page is visible at startup
        self.graph_page_visible = True
        self.toggle_graph_page_btn.config(text="Hide Graph Page (Ctrl+G)")
        
        # Add keyboard shortcuts
        self.root.bind('<Control-g>', lambda e: self.toggle_graph_page_visibility())
        self.root.bind('<Control-G>', lambda e: self.toggle_graph_page_visibility())
        self.root.bind('<F11>', lambda e: self.toggle_fullscreen_graph())
        self.root.bind('<Escape>', lambda e: self.exit_fullscreen_graph())
        # Change graph automatically on tab change
        self.tabs.bind('<<NotebookTabChanged>>', self._on_tab_changed)
        
        try:
            # Prefer Opera, then Opera GX, else All Browsers
            if "Opera" in self.available_browsers:
                self.browser_combo.set("Opera")
            elif "Opera GX" in self.available_browsers:
                self.browser_combo.set("Opera GX")
            else:
                self.browser_combo.set("All Browsers")
            self.load_data()
        except Exception:
            pass

    def reset_ui_defaults(self):
        try:
            # Reset date range
            self.prefill_dates()
            # Reset browser selection
            if "All Browsers" in self.browser_combo["values"]:
                self.browser_combo.set("All Browsers")
            # Reset graph type
            if hasattr(self, 'graph_type'):
                self.graph_type.set("pie")
            # Reset chart style
            self.chart_style = {
                'background_color': '#ffffff',
                'grid_color': '#e0e0e0',
                'bar_color': '#1f77b4',
                'line_color': '#ff7f0e',
                'title_color': '#2c3e50',
                'label_color': '#34495e',
                'font_size': 12,
                'grid_alpha': 0.3
            }
            # Sync quick style entries if present
            if hasattr(self, 'bg_color_var'):
                self.bg_color_var.set(self.chart_style['background_color'])
            if hasattr(self, 'bar_color_var'):
                self.bar_color_var.set(self.chart_style['bar_color'])

            # Global search reset - REMOVED
            # Ensure graph page visible
            if not self.graph_page_visible:
                self.toggle_graph_page_visibility()
            # Reset graph context label
            if hasattr(self, 'graph_context_label'):
                self.current_graph_context = "Installed Browsers"
                self.graph_context_label.config(text=self.current_graph_context)

            # Reload data and redraw
            self.load_data()
            self.redraw_chart()
            self.status_var.set("Reset to defaults")
        except Exception:
            # Best-effort reset
            self.status_var.set("Reset attempted")

    def apply_date_filter(self):
        """Apply the selected date range filter and reload data"""
        try:
            start_date, end_date = self.get_selected_date_range()
            self.status_var.set(f"Applying date filter: {date_only_ddmmyy(start_date)} to {date_only_ddmmyy(end_date)}...")
            self.root.update_idletasks()
            
            # Reload data with new date range
            self.load_data()
            
            self.status_var.set(f"Date filter applied: {date_only_ddmmyy(start_date)} to {date_only_ddmmyy(end_date)}")
        except Exception as e:
            self.status_var.set(f"Error applying date filter: {str(e)}")
            messagebox.showerror("Error", f"Failed to apply date filter: {str(e)}")

    def apply_global_search(self):
        """Apply global search across all tabs"""
        if not hasattr(self, 'global_search_var'):
            return
        
        search_term = self.global_search_var.get().lower().strip()
        # Ignore placeholder text
        if search_term == "search across all data...":
            search_term = ""
        
        # Apply search to each tab
        for tab_name, tree in self.tree_views.items():
            if not tree:
                continue
                
            # Store original data if not already stored
            if not hasattr(self, 'original_data'):
                self.original_data = {}
            
            if tab_name not in self.original_data:
                # Store original data
                self.original_data[tab_name] = []
                for item in tree.get_children():
                    values = tree.item(item)['values']
                    self.original_data[tab_name].append(values)
            
            # Clear current tree
            for item in tree.get_children():
                tree.delete(item)
            
            # Filter and re-populate
            for values in self.original_data[tab_name]:
                # Search in all columns
                match_found = False
                if not search_term:  # If search is empty, show all
                    match_found = True
                else:
                    for value in values:
                        if str(value).lower().find(search_term) != -1:
                            match_found = True
                            break
                
                if match_found:
                    tree.insert("", "end", values=values)
        
        # Update status
        if search_term:
            self.status_var.set(f"Search applied: '{search_term}'")
        else:
            self.status_var.set("Search cleared - showing all data")

    def clear_global_search(self):
        """Clear global search and show all data"""
        if hasattr(self, 'global_search_var'):
            self.global_search_var.set("")
            self.apply_global_search()
            # Reset placeholder
            if hasattr(self, 'search_entry'):
                self.search_entry.delete(0, tk.END)
                self.search_entry.insert(0, "Search across all data...")
                self.search_entry.config(foreground='gray')

    def on_search_focus_in(self, event):
        """Handle search entry focus in"""
        if self.search_entry.get() == "Search across all data...":
            self.search_entry.delete(0, tk.END)
            self.search_entry.config(foreground='black')

    def on_search_focus_out(self, event):
        """Handle search entry focus out"""
        if not self.search_entry.get():
            self.search_entry.insert(0, "Search across all data...")
            self.search_entry.config(foreground='gray')

    def _show_menu(self, event, tree):
        iid = tree.identify_row(event.y)
        if iid:
            tree.selection_set(iid)
            self.bm_menu.post(event.x_root, event.y_root)

    def _bookmark_open(self, tree):
        try:
            import webbrowser
            sel = tree.selection()
            if not sel:
                return
            url = tree.item(sel[0], 'values')[2]
            if url:
                webbrowser.open(url)
        except Exception:
            pass

    def _bookmark_copy(self, tree, which="url"):
        try:
            sel = tree.selection()
            if not sel:
                return
            vals = tree.item(sel[0], 'values')
            text = vals[2] if which == "url" else vals[1]
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.status_var.set(f"Copied {which.upper()} to clipboard.")
        except Exception:
            pass

    def prefill_dates(self):
        mins = []
        maxs = []
        for info in self.available_browsers.values():
            if info.get("min_date"):
                mins.append(info["min_date"]) 
            if info.get("max_date"):
                maxs.append(info["max_date"]) 
        if mins:
            start = min(mins)
        else:
            start = datetime.now() - timedelta(days=7)
        if maxs:
            end = max(maxs)
        else:
            end = datetime.now()
        try:
            self.start_cal.set_date(start.date())
            self.end_cal.set_date(end.date())
        except Exception:
            self.start_cal.set_date((datetime.now() - timedelta(days=7)).date())
            self.end_cal.set_date(datetime.now().date())

    def check_browsers(self):
        found = {}
        for browser, paths in BROWSERS.items():
            path = paths.get(get_platform())
            if not path or not os.path.exists(path):
                # For Opera variants, aggressively discover a profile path containing History
                if browser == "Opera" or browser == "Opera GX":
                    discovered = discover_opera_profile_path()
                    if discovered:
                        path = discovered
                if not path or not os.path.exists(path):
                    continue
            
            if browser == "Firefox":
                # Handle Firefox profiles
                profiles = get_firefox_profiles(path)
                if profiles:
                    # Get date range from first profile
                    min_d, max_d = get_date_range_from_firefox_history(profiles[0])
                    # Add Firefox even if no history data found
                    if not min_d or not max_d:
                        min_d = datetime.now() - timedelta(days=30)  # Default range
                        max_d = datetime.now()
                    found[browser] = {"path": path, "min_date": min_d, "max_date": max_d, "profiles": profiles}
                else:
                    # Even if no profiles found, add Firefox if directory exists
                    min_d = datetime.now() - timedelta(days=30)
                    max_d = datetime.now()
                    found[browser] = {"path": path, "min_date": min_d, "max_date": max_d, "profiles": []}
            else:
                # Handle Chrome-based browsers
                history_file = os.path.join(path, "History")
                if os.path.exists(history_file):
                    min_d, max_d = get_date_range_from_chrome_history(history_file)
                    if min_d and max_d:
                        found[browser] = {"path": path, "min_date": min_d, "max_date": max_d}
                        continue
                # Add with default range so it shows in GUI even if empty/new
                min_d = datetime.now() - timedelta(days=30)
                max_d = datetime.now()
                found[browser] = {"path": path, "min_date": min_d, "max_date": max_d}
        return found

    def debug_firefox_detection(self):
        """Debug method to check Firefox detection"""
        firefox_path = BROWSERS["Firefox"].get(get_platform())
        print(f"Firefox path: {firefox_path}")
        print(f"Path exists: {os.path.exists(firefox_path) if firefox_path else False}")
        
        if firefox_path and os.path.exists(firefox_path):
            profiles = get_firefox_profiles(firefox_path)
            print(f"Found profiles: {len(profiles)}")
            for i, profile in enumerate(profiles):
                print(f"  Profile {i+1}: {profile}")
                places_path = os.path.join(profile, "places.sqlite")
                print(f"    places.sqlite exists: {os.path.exists(places_path)}")
        
        return firefox_path, os.path.exists(firefox_path) if firefox_path else False

    def refresh_overview(self):
        for item in self.overview_tree.get_children():
            self.overview_tree.delete(item)
        for browser, info in self.available_browsers.items():
            min_d = info.get("min_date")
            max_d = info.get("max_date")
            
            if browser == "Firefox":
                # Show profile count for Firefox without full path
                profile_count = len(info.get("profiles", []))
                hist_path = f"Firefox Profile Data ({profile_count} profiles)"
            else:
                hist_path = f"{browser} Browser Data"
            
            self.overview_tree.insert(
                "",
                "end",
                values=(
                    browser,
                    datetime_to_ddmmyy(min_d),
                    datetime_to_ddmmyy(max_d),
                    hist_path,
                ),
            )

    def get_selected_date_range(self):
        start_date = datetime.combine(self.start_cal.get_date(), datetime.min.time())
        end_date = datetime.combine(self.end_cal.get_date(), datetime.max.time())
        return start_date, end_date

    def load_data(self):
        start_date, end_date = self.get_selected_date_range()
        # Remember last date range for chart labels
        self.last_date_range = (start_date, end_date)
        selected = self.browser_combo.get()
        self.status_var.set(f"Loading data for {selected} from {date_only_ddmmyy(start_date)} to {date_only_ddmmyy(end_date)} ...")
        self.root.update_idletasks()

        for tree in self.tree_views.values():
            for item in tree.get_children():
                tree.delete(item)

        # Clear original data for search functionality
        if hasattr(self, 'original_data'):
            self.original_data.clear()

        self.bookmarks_master.clear()

        if selected == "All Browsers":
            targets = list(self.available_browsers.items())
        else:
            if selected not in self.available_browsers:
                messagebox.showerror("Error", "Please select a valid browser.")
                return
            targets = [(selected, self.available_browsers[selected])]

        total_rows = 0
        per_browser_counts = {}
        per_browser_download_bytes = {}
        # Time-series per tab: {tab_name: {browser: {date: count}}}
        time_series_counts = {
            "Browsing History": {},
            "Downloads": {},
            "Cookies": {},
            "Bookmarks": {},
            "Secrets": {},
        }
        self.all_data = {}  # Store all data for password handling
        
        for browser, info in targets:
            base_path = info["path"]
            
            if browser == "Firefox":
                data = extract_firefox_history_data(base_path, start_date, end_date)
            else:
                data = extract_chrome_history_data(base_path, start_date, end_date)
            
            browser_count = 0
            browser_download_bytes = 0
            
            # Store data for password methods
            self.all_data[browser] = data

            for title, url, when_dt in data["urls"]:
                # Show shortened URL and title with tooltips for full values
                short_url = shorten_url(url, 60)
                short_title = title[:40] + "..." if title and len(title) > 40 else title
                item = self.tree_views["Browsing History"].insert("", "end", values=(short_title, short_url, f"{datetime_to_ddmmyy(when_dt)} | {browser}"))
                # Store original values for tooltips
                if "Browsing History" in self.url_tooltips:
                    self.url_tooltips["Browsing History"].store_original_url(item, 0, title)  # Store full title
                    self.url_tooltips["Browsing History"].store_original_url(item, 1, url)   # Store full URL
                total_rows += 1
                browser_count += 1
                # Build time-series for Browsing History
                try:
                    day_key = when_dt.date()
                    ts_b = time_series_counts["Browsing History"].setdefault(browser, {})
                    ts_b[day_key] = ts_b.get(day_key, 0) + 1
                except Exception:
                    pass
            for file, source, when_dt, size_bytes in data["downloads"]:
                size_text = bytes_to_readable(size_bytes) if size_bytes is not None else ""
                # Show shortened source URL for better display
                short_source = shorten_url(source, 60) if source else ""
                item = self.tree_views["Downloads"].insert("", "end", values=(file, short_source, size_text, f"{datetime_to_ddmmyy(when_dt)} | {browser}"))
                # Store original URL for tooltip
                if "Downloads" in self.url_tooltips:
                    self.url_tooltips["Downloads"].store_original_url(item, 1, source)
                total_rows += 1
                browser_count += 1
                if size_bytes:
                    try:
                        browser_download_bytes += int(size_bytes)
                    except Exception:
                        pass
                # Build time-series for Downloads
                try:
                    day_key = when_dt.date()
                    ts_d = time_series_counts["Downloads"].setdefault(browser, {})
                    ts_d[day_key] = ts_d.get(day_key, 0) + 1
                except Exception:
                    pass
            for host, name, val, exp_dt in data["cookies"]:
                # Shorten cookie value and domain/name for display
                short_val = shorten_cookie_value(val, 30)  # Shorter for better UI display
                domain_name = f"{host} | {name}"
                short_domain_name = domain_name[:35] + "..." if len(domain_name) > 35 else domain_name
                item = self.tree_views["Cookies"].insert("", "end", values=(short_domain_name, short_val, datetime_to_ddmmyy(exp_dt), browser))
                # Store original values for tooltips
                if "Cookies" in self.url_tooltips:
                    self.url_tooltips["Cookies"].store_original_url(item, 0, domain_name)  # Store full domain/name
                    self.url_tooltips["Cookies"].store_original_url(item, 1, val)         # Store full cookie value
                total_rows += 1
                browser_count += 1
                # Build time-series for Cookies (by expiry date or today if missing)
                try:
                    day_key = (exp_dt or start_date).date() if hasattr(exp_dt, 'date') else start_date.date()
                    ts_c = time_series_counts["Cookies"].setdefault(browser, {})
                    ts_c[day_key] = ts_c.get(day_key, 0) + 1
                except Exception:
                    pass
            
            # Handle passwords
            for pwd_data in data.get("passwords", []):
                url = pwd_data.get('url', '')
                username = pwd_data.get('username', '')
                password = pwd_data.get('password_decrypted', '')
                encryption_type = pwd_data.get('encryption_type', '')
                date_created = datetime_to_ddmmyy(pwd_data.get('date_created'))
                date_last_used = datetime_to_ddmmyy(pwd_data.get('date_last_used'))
                
                # Show passwords by default (can be toggled)
                display_password = password if password else ""
                
                # Shorten URL for display
                short_url = shorten_url(url, 50)
                item = self.tree_views["Secrets"].insert("", "end", values=(
                    short_url, username, display_password, encryption_type,
                    date_created, browser
                ))
                # Store original URL for tooltip
                if "Secrets" in self.url_tooltips:
                    self.url_tooltips["Secrets"].store_original_url(item, 0, url)
                total_rows += 1
                browser_count += 1
                # Build time-series for Secrets (by date_created if present)
                try:
                    created_dt = pwd_data.get('date_created')
                    if created_dt and hasattr(created_dt, 'date'):
                        day_key = created_dt.date()
                        ts_s = time_series_counts["Secrets"].setdefault(browser, {})
                        ts_s[day_key] = ts_s.get(day_key, 0) + 1
                except Exception:
                    pass

            # Bookmarks -> build enhanced records
            for folder, name, url, added_dt in data.get("bookmarks", []):
                rec = {
                    "folder": folder,
                    "name": name,
                    "url": url,
                    "added": added_dt,
                    "domain": extract_domain(url),
                    "norm_url": normalize_url(url),
                    "dup": False,
                    "live": "?",  # unknown until checked
                    "browser": browser,
                }
                self.bookmarks_master.append(rec)
                # Build time-series for Bookmarks (by added date)
                try:
                    if added_dt and hasattr(added_dt, 'date'):
                        day_key = added_dt.date()
                        ts_bm = time_series_counts["Bookmarks"].setdefault(browser, {})
                        ts_bm[day_key] = ts_bm.get(day_key, 0) + 1
                except Exception:
                    pass

            per_browser_counts[browser] = browser_count
            per_browser_download_bytes[browser] = browser_download_bytes

        if selected == "All Browsers":
            for b in self.available_browsers.keys():
                per_browser_counts.setdefault(b, 0)
                per_browser_download_bytes.setdefault(b, 0)

        # Populate bookmarks table
        self._render_bookmarks_table()

        self.status_var.set(f"Loaded {total_rows} rows across tabs. Range: {date_only_ddmmyy(start_date)} - {date_only_ddmmyy(end_date)}.")
        self.draw_browser_data_chart(per_browser_counts, per_browser_download_bytes, time_series_counts)
        
    def toggle_graph_page_visibility(self):
        """Toggle entire graph page visibility"""
        self.graph_page_visible = not self.graph_page_visible
        
        if self.graph_page_visible:
            self.toggle_graph_page_btn.config(text="Hide Graph Page (Ctrl+G)")
            # Show the entire chart page
            self.chart_page_frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))
            # Adjust tabs to normal size when graph is visible
            self.tabs.pack_configure(expand=1, fill="both")
            # Hide full screen message
            self.full_screen_message.pack_forget()
            # Redraw the chart if we have data
            if hasattr(self, 'last_chart_data'):
                self.draw_browser_data_chart(*self.last_chart_data)
            self.status_var.set("Graph page is now visible")
        else:
            self.toggle_graph_page_btn.config(text="Show Graph Page (Ctrl+G)")
            # Hide the entire chart page
            self.chart_page_frame.pack_forget()
            # Make tabs expand to full screen when graph is hidden
            self.tabs.pack_configure(expand=True, fill="both")
            # Show full screen message
            self.full_screen_message.pack(fill="x", padx=12, pady=(0, 5))
            self.status_var.set("Graph page is now hidden; tabs expanded")
    
    def toggle_fullscreen_graph(self):
        """Toggle fullscreen graph mode"""
        if not self.graph_fullscreen:
            # Enter fullscreen mode
            self.graph_fullscreen = True
            self.fullscreen_btn.config(text="Exit Full Screen (ESC)")
            
            # Hide all other elements
            # Remember current visibility to restore exactly
            self._pre_fullscreen_graph_visible = getattr(self, 'graph_page_visible', True)
            self.tabs.pack_forget()
            
            # Make chart container fill entire window
            self.chart_page_frame.pack_configure(fill="both", expand=True, padx=0, pady=0)
            
            # Create fullscreen chart with larger size
            self.create_fullscreen_chart()
            
            self.status_var.set("Full Screen Graph Mode - Press ESC or click button to exit")
        else:
            # Exit fullscreen mode
            self.exit_fullscreen_graph()
    
    def exit_fullscreen_graph(self):
        """Exit fullscreen graph mode"""
        if self.graph_fullscreen:
            self.graph_fullscreen = False
            self.fullscreen_btn.config(text="Full Screen (F11)")
            
            # Restore normal layout
            # Ensure tabs are packed ABOVE the graph page by briefly un-packing the graph
            try:
                self.chart_page_frame.pack_forget()
            except Exception:
                pass
            self.tabs.pack(expand=1, fill="both", padx=8, pady=8)
            # Restore prior graph page visibility
            if getattr(self, '_pre_fullscreen_graph_visible', True):
                self.graph_page_visible = True
                self.chart_page_frame.pack_configure(fill="both", expand=True, padx=12, pady=(0, 8))
                self.full_screen_message.pack_forget()
            else:
                self.graph_page_visible = False
                self.chart_page_frame.pack_forget()
                self.full_screen_message.pack(fill="x", padx=12, pady=(0, 5))
            
            # Redraw normal chart
            if hasattr(self, 'last_chart_data'):
                self.draw_browser_data_chart(*self.last_chart_data)
            
            self.status_var.set("Exited fullscreen mode")
    
    def create_fullscreen_chart(self):
        """Create enhanced fullscreen chart"""
        # Clear existing chart
        for child in self.chart_container.winfo_children():
            child.destroy()
        
        if not hasattr(self, 'last_chart_data') or not self.last_chart_data[0]:
            lbl = ttk.Label(self.chart_container, text="No data available for fullscreen chart", 
                           font=("Arial", 16), anchor="center")
            lbl.pack(fill="both", expand=True)
            return
        
        # Create larger enhanced graph for fullscreen
        self.enhanced_graph = EnhancedGraph(self.chart_container)
        
        # Get current data
        counts_dict, download_bytes_dict, time_series_counts = self.last_chart_data
        current_ctx = getattr(self, 'current_graph_context', 'Browsing History')
        
        # Prepare fullscreen data based on current context
        history_data = []
        if hasattr(self, 'all_data'):
            for browser, data in self.all_data.items():
                if current_ctx == "Browsing History":
                    for title, url, when_dt in data.get("urls", []):
                        if when_dt:
                            history_data.append({
                                'visit_time': when_dt,
                                'url': url or 'Unknown URL',
                                'title': title or 'No Title',
                                'browser': browser
                            })
                elif current_ctx == "Downloads":
                    for filename, source, when_dt, size_bytes in data.get("downloads", []):
                        if when_dt:
                            size_text = bytes_to_readable(size_bytes) if size_bytes is not None else ""
                            history_data.append({
                                'visit_time': when_dt,
                                'url': source or 'Unknown Source',
                                'title': f"Downloaded {filename or '(unknown)'}{f' - {size_text}' if size_text else ''}",
                                'browser': browser,
                                'size_bytes': size_bytes
                            })
                elif current_ctx == "Cookies":
                    for host, name, val, exp_dt in data.get("cookies", []):
                        if exp_dt:
                            history_data.append({
                                'visit_time': exp_dt,
                                'url': host or 'Cookie',
                                'title': f"Cookie {name or '(no name)'}",
                                'browser': browser
                            })
                elif current_ctx == "Bookmarks":
                    for folder, name, url, added_dt in data.get("bookmarks", []):
                        if added_dt:
                            history_data.append({
                                'visit_time': added_dt,
                                'url': url or 'Unknown URL',
                                'title': f"{name or '(no title)'} — {folder or 'Bookmarks'}",
                                'browser': browser
                            })
                elif current_ctx == "Secrets":
                    for pwd_data in data.get("passwords", []):
                        dt = pwd_data.get('date_last_used') or pwd_data.get('date_created')
                        if not dt:
                            try:
                                dt = datetime.now()
                            except Exception:
                                dt = None
                        if dt:
                            history_data.append({
                                'visit_time': dt,
                                'url': pwd_data.get('url', '') or 'Unknown URL',
                                'title': f"Login: {pwd_data.get('username', '') or '(no username)'}",
                                'browser': browser
                            })
        
        # Plot fullscreen chart
        if current_ctx == "Installed Browsers":
            self.create_fullscreen_pie_chart(counts_dict)
        else:
            title = f"Full Screen: {current_ctx} Analysis"
            self.enhanced_graph.plot_daily_activity(history_data, title, current_ctx)
    
    def create_fullscreen_pie_chart(self, counts_dict):
        """Create fullscreen pie chart for browser distribution"""
        browsers = list(counts_dict.keys())
        counts = [counts_dict[b] for b in browsers]
        
        # Create fullscreen figure
        plt.style.use('default')
        fig, ax = plt.subplots(figsize=(16, 10))
        
        # Apply styling
        fig.patch.set_facecolor('white')
        ax.set_facecolor('white')
        
        total = sum(counts)
        if total <= 0:
            ax.text(0.5, 0.5, "No browser data available", ha='center', va='center', 
                   transform=ax.transAxes, fontsize=24, 
                   bbox=dict(boxstyle="round,pad=1", facecolor='lightgray', alpha=0.7))
            ax.axis('off')
        else:
            # Enhanced pie chart with better colors and labels
            colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8']
            wedges, texts, autotexts = ax.pie(counts, labels=browsers, autopct='%1.1f%%', 
                                            colors=colors[:len(browsers)], startangle=90,
                                            textprops={'fontsize': 14, 'fontweight': 'bold'},
                                            explode=[0.05] * len(browsers))
            
            # Enhance text appearance
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontsize(16)
                autotext.set_fontweight('bold')
            
            ax.set_title("Browser Distribution Analysis", fontsize=24, fontweight='bold', 
                        pad=30, color='#2c3e50')
        
        fig.tight_layout()
        
        # Add to container
        canvas = FigureCanvasTkAgg(fig, master=self.chart_container)
        canvas.draw()
        widget = canvas.get_tk_widget()
        widget.pack(fill="both", expand=True)
        self.chart_canvas_widget = canvas
        self.current_fig = fig
            
    def redraw_chart(self):
        """Redraw chart with current settings"""
        if self.graph_page_visible and hasattr(self, 'last_chart_data'):
            self.draw_browser_data_chart(*self.last_chart_data)
    
    def clear_enhanced_graph(self):
        """Clear the enhanced graph"""
        if hasattr(self, 'enhanced_graph'):
            self.enhanced_graph.clear_plot()
    
    def _on_tab_changed(self, event):
        """Adjust graph style and labels based on selected tab."""
        try:
            selected_tab = self.tabs.select()
            tab_text = self.tabs.tab(selected_tab, "text")
            self.current_graph_context = tab_text
            if hasattr(self, 'graph_context_label'):
                self.graph_context_label.config(text=tab_text)
            
            # Installed Browsers => force pie chart and disable switching
            if tab_text == "Installed Browsers":
                self.graph_type.set("pie")
                try:
                    self.graph_type_combo.set("pie")
                    self.graph_type_combo.configure(state="disabled")
                except Exception:
                    pass
                # Show pie chart immediately for Installed Browsers
                if hasattr(self, 'last_chart_data') and self.last_chart_data[0]:
                    self.draw_pie_chart_for_browsers(self.last_chart_data[0])
            else:
                # Other tabs => enable line chart
                self.graph_type.set("line")
                try:
                    self.graph_type_combo.configure(state="readonly")
                except Exception:
                    pass
                # Redraw as line chart for other tabs
                self.redraw_chart()
        except Exception:
            pass
    
    def draw_pie_chart_for_browsers(self, counts_dict):
        """Draw pie chart specifically for browser distribution"""
        if not self.graph_page_visible:
            return
            
        # Clear existing chart
        for child in self.chart_container.winfo_children():
            child.destroy()
            
        browsers = list(counts_dict.keys())
        counts = [counts_dict[b] for b in browsers]
        
        # Create pie chart
        plt.style.use('default')
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Apply custom styling
        fig.patch.set_facecolor(self.chart_style['background_color'])
        ax.set_facecolor(self.chart_style['background_color'])
        
        total = sum(counts)
        if total <= 0:
            ax.text(0.5, 0.5, "💻 No browser data available\nकोई ब्राउज़र डेटा उपलब्ध नहीं", 
                   ha='center', va='center', transform=ax.transAxes,
                   fontsize=14, bbox=dict(boxstyle="round,pad=0.5", facecolor='lightgray', alpha=0.7))
            ax.axis('off')
        else:
            # Enhanced pie chart with better colors
            colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD']
            wedges, texts, autotexts = ax.pie(counts, labels=browsers, autopct='%1.1f%%', 
                                            colors=colors[:len(browsers)], startangle=90,
                                            textprops={'fontsize': 11, 'fontweight': 'bold'},
                                            explode=[0.02] * len(browsers))
            
            # Enhance percentage text
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontweight('bold')
            
            ax.set_title("Browser Distribution", 
                        color=self.chart_style['title_color'], 
                        fontsize=self.chart_style['font_size'] + 4, fontweight='bold')
        
        fig.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, master=self.chart_container)
        canvas.draw()
        widget = canvas.get_tk_widget()
        widget.pack(fill="both", expand=True)
        self.chart_canvas_widget = canvas
        self.current_fig = fig
            
    def apply_quick_chart_style(self, style_key, value):
        """Apply quick style changes"""
        try:
            if style_key == 'font_size':
                self.chart_style[style_key] = int(value)
            else:
                self.chart_style[style_key] = value
                
            self.redraw_chart()
            self.status_var.set(f"Applied {style_key}: {value}")
        except Exception as e:
            messagebox.showerror("Error", f"Invalid value for {style_key}: {e}")
            
    def open_chart_style_dialog(self):
        """Open advanced chart style customization dialog"""
        style_dialog = tk.Toplevel(self.root)
        style_dialog.title("Advanced Chart Style")
        style_dialog.geometry("500x600")
        style_dialog.transient(self.root)
        style_dialog.grab_set()
        
        # Create style editor
        notebook = ttk.Notebook(style_dialog)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Colors tab
        colors_frame = ttk.Frame(notebook)
        notebook.add(colors_frame, text="Colors")
        
        color_options = [
            ("Background Color", "background_color"),
            ("Grid Color", "grid_color"),
            ("Bar Color", "bar_color"),
            ("Line Color", "line_color"),
            ("Title Color", "title_color"),
            ("Label Color", "label_color")
        ]
        
        for i, (label, key) in enumerate(color_options):
            ttk.Label(colors_frame, text=label).grid(row=i, column=0, padx=5, pady=5, sticky="w")
            color_var = tk.StringVar(value=self.chart_style[key])
            color_entry = ttk.Entry(colors_frame, textvariable=color_var, width=15)
            color_entry.grid(row=i, column=1, padx=5, pady=5)
            
            def apply_color(k=key, v=color_var):
                self.chart_style[k] = v.get()
                self.redraw_chart()
                
            ttk.Button(colors_frame, text="Apply", command=apply_color).grid(row=i, column=2, padx=5, pady=5)
        
        # Appearance tab
        appearance_frame = ttk.Frame(notebook)
        notebook.add(appearance_frame, text="Appearance")
        
        # Font size
        ttk.Label(appearance_frame, text="Font Size:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        font_size_var = tk.IntVar(value=self.chart_style['font_size'])
        font_spin = ttk.Spinbox(appearance_frame, from_=8, to=24, textvariable=font_size_var, width=10)
        font_spin.grid(row=0, column=1, padx=5, pady=5)
        
        # Grid alpha
        ttk.Label(appearance_frame, text="Grid Alpha:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        grid_alpha_var = tk.DoubleVar(value=self.chart_style['grid_alpha'])
        grid_alpha_spin = ttk.Spinbox(appearance_frame, from_=0.0, to=1.0, increment=0.1, textvariable=grid_alpha_var, width=10)
        grid_alpha_spin.grid(row=1, column=1, padx=5, pady=5)
        
        def apply_appearance():
            self.chart_style['font_size'] = font_size_var.get()
            self.chart_style['grid_alpha'] = grid_alpha_var.get()
            self.redraw_chart()
            
        ttk.Button(appearance_frame, text="Apply All", command=apply_appearance).grid(row=2, column=0, columnspan=2, pady=20)
        
        # Presets tab
        presets_frame = ttk.Frame(notebook)
        notebook.add(presets_frame, text="Presets")
        
        preset_styles = {
            "Default": {
                'background_color': '#ffffff',
                'grid_color': '#e0e0e0',
                'bar_color': '#1f77b4',
                'line_color': '#ff7f0e',
                'title_color': '#2c3e50',
                'label_color': '#34495e',
                'font_size': 12,
                'grid_alpha': 0.3
            },
            "Dark Theme": {
                'background_color': '#2c3e50',
                'grid_color': '#34495e',
                'bar_color': '#3498db',
                'line_color': '#e74c3c',
                'title_color': '#ecf0f1',
                'label_color': '#bdc3c7',
                'font_size': 12,
                'grid_alpha': 0.4
            },
            "Professional": {
                'background_color': '#f8f9fa',
                'grid_color': '#dee2e6',
                'bar_color': '#495057',
                'line_color': '#6c757d',
                'title_color': '#212529',
                'label_color': '#495057',
                'font_size': 14,
                'grid_alpha': 0.2
            },
            "Colorful": {
                'background_color': '#f0f8ff',
                'grid_color': '#d4edda',
                'bar_color': '#dc3545',
                'line_color': '#28a745',
                'title_color': '#6f42c1',
                'label_color': '#fd7e14',
                'font_size': 13,
                'grid_alpha': 0.3
            }
        }
        
        for i, (preset_name, preset_style) in enumerate(preset_styles.items()):
            def apply_preset(style=preset_style):
                self.chart_style.update(style)
                self.redraw_chart()
                self.status_var.set(f"Applied {preset_name} preset")
                
            ttk.Button(
                presets_frame, 
                text=preset_name, 
                command=apply_preset
            ).grid(row=i, column=0, padx=20, pady=10, sticky="ew")
            
        # Close button
        ttk.Button(style_dialog, text="Close", command=style_dialog.destroy).pack(pady=10)
        
    def export_chart(self):
        """Export the current chart"""
        if not self.current_fig:
            messagebox.showwarning("Warning", "No chart to export")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[
                ("PNG files", "*.png"),
                ("PDF files", "*.pdf"),
                ("SVG files", "*.svg"),
                ("All files", "*.*")
            ],
            title="Export Chart"
        )
        
        if file_path:
            try:
                self.current_fig.savefig(file_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("Success", f"Chart exported to:\n{file_path}")
                self.status_var.set(f"Chart exported to: {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export chart: {e}")

    def draw_browser_data_chart(self, counts_dict, download_bytes_dict, time_series_counts=None):
        if not self.graph_page_visible:
            return
            
        # Clear existing chart
        for child in self.chart_container.winfo_children():
            child.destroy()
            
        # Determine desired graph type early
        graph_type = self.graph_type.get() if hasattr(self, 'graph_type') else 'line'
        
        # If pie chart requested but no counts, show message and return.
        if graph_type != 'line' and not counts_dict:
            lbl = ttk.Label(self.chart_container, text="No browsers detected to chart.", anchor="center")
            lbl.pack(fill="both", expand=True, pady=8)
            return

        # Reinitialize enhanced graph after clearing
        self.enhanced_graph = EnhancedGraph(self.chart_container)
        
        # Prepare data for enhanced graph
        history_data = []
        
        # Convert time series data to format expected by enhanced graph
        current_ctx = getattr(self, 'current_graph_context', 'Browsing History')
        
        if time_series_counts and current_ctx in time_series_counts:
            series_for_ctx = time_series_counts[current_ctx]
            
            # Convert to list of dictionaries with visit_time
            for browser, daily_data in series_for_ctx.items():
                for date, count in daily_data.items():
                    # Create multiple entries for each count to show activity level
                    for i in range(count):
                        history_data.append({
                            'visit_time': datetime.combine(date, datetime.min.time()),
                            'url': f'Activity from {browser}',
                            'title': f'{browser} Activity {i+1}',
                            'browser': browser
                        })
        
        # If we have actual data, build entries based on active tab context
        if hasattr(self, 'all_data'):
            actual_history = []
            if current_ctx == "Browsing History":
                for browser, data in self.all_data.items():
                    for title, url, when_dt in data.get("urls", []):
                        if when_dt:
                            actual_history.append({
                                'visit_time': when_dt,
                                'url': url or 'Unknown URL',
                                'title': title or 'No Title',
                                'browser': browser
                            })
            elif current_ctx == "Downloads":
                for browser, data in self.all_data.items():
                    for filename, source, when_dt, size_bytes in data.get("downloads", []):
                        if when_dt:
                            size_text = bytes_to_readable(size_bytes) if size_bytes is not None else ""
                            actual_history.append({
                                'visit_time': when_dt,
                                'url': source or 'Unknown Source',
                                'title': f"Downloaded {filename or '(unknown)'}{f' - {size_text}' if size_text else ''}",
                                'browser': browser
                            })
            elif current_ctx == "Cookies":
                for browser, data in self.all_data.items():
                    for host, name, val, exp_dt in data.get("cookies", []):
                        if exp_dt:
                            actual_history.append({
                                'visit_time': exp_dt,
                                'url': host or 'Cookie',
                                'title': f"Cookie {name or '(no name)'}",
                                'browser': browser
                            })
            elif current_ctx == "Bookmarks":
                for browser, data in self.all_data.items():
                    for folder, name, url, added_dt in data.get("bookmarks", []):
                        if added_dt:
                            actual_history.append({
                                'visit_time': added_dt,
                                'url': url or 'Unknown URL',
                                'title': f"{name or '(no title)'} — {folder or 'Bookmarks'}",
                                'browser': browser
                            })
            elif current_ctx == "Secrets":
                for browser, data in self.all_data.items():
                    for pwd_data in data.get("passwords", []):
                        # Prefer last used when present
                        dt = pwd_data.get('date_last_used') or pwd_data.get('date_created')
                        if not dt:
                            try:
                                dt = datetime.now()
                            except Exception:
                                dt = None
                        if dt:
                            actual_history.append({
                                'visit_time': dt,
                                'url': pwd_data.get('url', '') or 'Unknown URL',
                                'title': f"Login: {pwd_data.get('username', '') or '(no username)'}",
                                'browser': browser
                            })

            if actual_history:
                history_data = actual_history
        
        # Plot using enhanced graph
        
        if graph_type == "line" and history_data:
            # Use enhanced graph for line charts with hover details
            title = f"Daily Activity - {current_ctx}"
            self.enhanced_graph.plot_daily_activity(history_data, title, current_ctx)
        else:
            # Fallback to original pie chart implementation
            browsers = list(counts_dict.keys())
            counts = [counts_dict[b] for b in browsers]
            
            # Apply custom styling
            plt.style.use('default')
            fig, ax = plt.subplots(figsize=(10.5, 5.5))
            
            # Apply custom colors and styling
            fig.patch.set_facecolor(self.chart_style['background_color'])
            ax.set_facecolor(self.chart_style['background_color'])
            
            # For pie chart, show only counts; handle zero-total safely
            total = sum(counts)
            if total <= 0:
                ax.text(0.5, 0.5, "No data to plot", ha='center', va='center', transform=ax.transAxes,
                        color=self.chart_style['label_color'], fontsize=self.chart_style['font_size'])
                ax.axis('off')
            else:
                ax.pie(counts, labels=browsers, autopct='%1.1f%%', colors=plt.cm.Set3.colors, startangle=90)
                ax.set_title("Items per Browser", color=self.chart_style['title_color'], fontsize=self.chart_style['font_size'] + 2)
            
            fig.tight_layout()
            
            canvas = FigureCanvasTkAgg(fig, master=self.chart_container)
            canvas.draw()
            widget = canvas.get_tk_widget()
            widget.pack(fill="both", expand=True)
            self.chart_canvas_widget = canvas
            self.current_fig = fig
        
        # Store chart data for redrawing
        self.last_chart_data = (counts_dict, download_bytes_dict, time_series_counts)

    # ---------- Sorting helper ----------
    def _sort_tree(self, tree, col, reverse):
        try:
            data = [(tree.set(k, col), k) for k in tree.get_children("")]
            # try numeric sort
            try:
                data.sort(key=lambda t: float(t[0]), reverse=reverse)
            except Exception:
                data.sort(key=lambda t: t[0], reverse=reverse)
            for index, (_, k) in enumerate(data):
                tree.move(k, "", index)
            tree.heading(col, command=lambda: self._sort_tree(tree, col, not reverse))
        except Exception:
            pass

    # ---------- Bookmark features ----------


    def _render_bookmarks_table(self):
        tree = self.tree_views["Bookmarks"]
        for item in tree.get_children():
            tree.delete(item)
        for rec in self.bookmarks_master:
            added_text = date_only_ddmmyy(rec["added"]) if rec["added"] else ""
            # Shorten URL for display
            short_url = shorten_url(rec["url"], 50)
            item = tree.insert("", "end", values=(rec["folder"], rec["name"], short_url, added_text, rec["domain"], rec["browser"]))
            # Store original URL for tooltip
            if "Bookmarks" in self.url_tooltips:
                self.url_tooltips["Bookmarks"].store_original_url(item, 2, rec["url"])
        self.status_var.set(f"Bookmarks shown: {len(self.bookmarks_master)}")



    def _export_bookmarks_dialog(self):
        path = filedialog.asksaveasfilename(defaultextension=",csv", filetypes=[("CSV", "*.csv"), ("HTML", "*.html"), ("All", "*.*")], initialfile="bookmarks_export")
        if not path:
            return
        try:
            if path.lower().endswith(".html"):
                html = self._bookmarks_to_html(self.bookmarks_master)
                with open(path, "w", encoding="utf-8") as f:
                    f.write(html)
            else:
                # CSV
                import csv
                rows = self._bookmarks_rows(self.bookmarks_master)
                with open(path, "w", newline="", encoding="utf-8") as f:
                    w = csv.writer(f)
                    w.writerow(["Folder", "Name", "URL", "Added", "Domain", "Browser"])
                    for r in rows:
                        w.writerow(r)
            messagebox.showinfo("Saved", f"Bookmarks exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export: {e}")

    def _bookmarks_rows(self, recs):
        rows = []
        for rec in recs:
            rows.append([
                rec["folder"], rec["name"], rec["url"], date_only_ddmmyy(rec["added"]) if rec["added"] else "",
                rec["domain"], rec["browser"]
            ])
        return rows

    def _bookmarks_to_html(self, recs):
        rows = self._bookmarks_rows(recs)
        # quick stats
        total = len(recs)
        dupes = sum(1 for r in recs if r["dup"] if isinstance(r, dict)) if recs and isinstance(recs[0], dict) else sum(1 for r in recs if (r[5] == "Yes"))
        domains = {}
        for rec in recs:
            d = rec["domain"] if isinstance(rec, dict) else rec[4]
            if d:
                domains[d] = domains.get(d, 0) + 1
        top_domains = sorted(domains.items(), key=lambda x: x[1], reverse=True)[:15]

        def esc(s):
            return (str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;').replace("'",'&#39;'))

        parts = [
            "<!doctype html><html><head><meta charset='utf-8'><title>Bookmarks Export</title>",
            "<style>body{font-family:Segoe UI,Arial;padding:20px} table{border-collapse:collapse;width:100%} th,td{border:1px solid #ddd;padding:6px;text-align:left} th{background:#f3f4f6}</style>",
            "</head><body>",
            f"<h1>Bookmarks Export</h1><p><strong>Total:</strong> {total} &nbsp; <strong>Duplicates:</strong> {dupes}</p>",
            "<h2>Top Domains</h2><ol>",
        ]
        for d, c in top_domains:
            parts.append(f"<li>{esc(d)} — {c}</li>")
        parts.append("</ol><hr/><h2>Bookmarks</h2><table><thead><tr>")
        headers = ["Folder", "Name", "URL", "Added", "Domain", "Browser"]
        for h in headers:
            parts.append(f"<th>{esc(h)}</th>")
        parts.append("</tr></thead><tbody>")
        for r in self._bookmarks_rows(recs):
            parts.append("<tr>" + "".join(f"<td>{esc(c)}</td>" for c in r) + "</tr>")
        parts.append("</tbody></table></body></html>")
        return "".join(parts)

    def _get_top_visited_urls(self, limit=10):
        """Get top visited URLs from all browser data"""
        try:
            url_counts = {}
            url_details = {}
            
            # Aggregate data from all browsers
            for browser, data in getattr(self, 'all_data', {}).items():
                for title, url, when_dt in data.get("urls", []):
                    if url:
                        # Normalize URL for counting
                        normalized_url = normalize_url(url)
                        if normalized_url:
                            url_counts[normalized_url] = url_counts.get(normalized_url, 0) + 1
                            # Keep the most recent title and visit time
                            if normalized_url not in url_details or when_dt > url_details[normalized_url][2]:
                                url_details[normalized_url] = (title or "No Title", url, when_dt)
            
            # Sort by visit count and get top URLs
            top_urls = sorted(url_counts.items(), key=lambda x: x[1], reverse=True)[:limit]
            
            # Format for table display
            result = []
            for url, count in top_urls:
                if url in url_details:
                    title, original_url, last_visit = url_details[url]
                    result.append([
                        shorten_url(original_url, 50),
                        title[:50] + "..." if len(title) > 50 else title,
                        str(count),
                        datetime_to_ddmmyy(last_visit) if last_visit else ""
                    ])
            
            return result
        except Exception as e:
            print(f"Error getting top visited URLs: {e}")
            return []

    # Export helpers for full HTML
    def _get_tree_data(self, tree, use_full_urls=False):
        headers = []
        col_ids = tree["columns"]
        for cid in col_ids:
            hdr = tree.heading(cid).get("text", cid)
            headers.append(hdr)
        rows = []
        
        # Determine which tree this is for proper data restoration
        tree_type = None
        for tab_name, tab_tree in self.tree_views.items():
            if tab_tree == tree:
                tree_type = tab_name
                break
        
        for iid in tree.get_children():
            vals = list(tree.item(iid, "values"))
            
            # Restore original values when exporting full data
            if use_full_urls and tree_type in self.url_tooltips:
                tooltip_handler = self.url_tooltips[tree_type]
                
                if tree_type == "Downloads":
                    # Restore original URL (column 1)
                    original_url = tooltip_handler.original_values.get((iid, 1))
                    if original_url and len(vals) > 1:
                        vals[1] = original_url
                        
                elif tree_type == "Secrets":
                    # Restore original URL (column 0)
                    original_url = tooltip_handler.original_values.get((iid, 0))
                    if original_url and len(vals) > 0:
                        vals[0] = original_url
                        
                elif tree_type == "Browsing History":
                    # Restore original title (column 0) and URL (column 1)
                    original_title = tooltip_handler.original_values.get((iid, 0))
                    original_url = tooltip_handler.original_values.get((iid, 1))
                    if original_title and len(vals) > 0:
                        vals[0] = original_title
                    if original_url and len(vals) > 1:
                        vals[1] = original_url
                        
                elif tree_type == "Bookmarks":
                    # Restore original URL (column 2)
                    original_url = tooltip_handler.original_values.get((iid, 2))
                    if original_url and len(vals) > 2:
                        vals[2] = original_url
                        
                elif tree_type == "Cookies":
                    # Restore original domain/name (column 0) and cookie value (column 1)
                    original_domain = tooltip_handler.original_values.get((iid, 0))
                    original_value = tooltip_handler.original_values.get((iid, 1))
                    if original_domain and len(vals) > 0:
                        vals[0] = original_domain
                    if original_value and len(vals) > 1:
                        vals[1] = original_value
            
            rows.append(vals)
        return headers, rows

    def _get_tree_data_with_tooltips(self, tree, export_type="html"):
        """Get tree data with shortened URLs and full URLs stored for tooltips"""
        headers = []
        col_ids = tree["columns"]
        for cid in col_ids:
            hdr = tree.heading(cid).get("text", cid)
            headers.append(hdr)
        rows = []
        tooltip_data = []  # Store full URLs for tooltips
        
        # Determine which tree this is for proper data restoration
        tree_type = None
        for tab_name, tab_tree in self.tree_views.items():
            if tab_tree == tree:
                tree_type = tab_name
                break
        
        for iid in tree.get_children():
            vals = list(tree.item(iid, "values"))
            tooltip_row = []  # Store full URLs for this row
            
            # Handle URL shortening and store full URLs for tooltips
            if tree_type in self.url_tooltips:
                tooltip_handler = self.url_tooltips[tree_type]
                
                if tree_type == "Downloads" and len(vals) > 1:
                    original_url = tooltip_handler.original_values.get((iid, 1))
                    if original_url:
                        vals[1] = shorten_url_for_export(original_url, export_type)
                        tooltip_row.append((1, original_url))  # Column 1, full URL
                        
                elif tree_type == "Secrets" and len(vals) > 0:
                    original_url = tooltip_handler.original_values.get((iid, 0))
                    if original_url:
                        vals[0] = shorten_url_for_export(original_url, export_type)
                        tooltip_row.append((0, original_url))  # Column 0, full URL
                        
                elif tree_type == "Browsing History":
                    if len(vals) > 0:
                        original_title = tooltip_handler.original_values.get((iid, 0))
                        if original_title and len(original_title) > 50:
                            vals[0] = original_title[:47] + "..."
                            tooltip_row.append((0, original_title))  # Column 0, full title
                    if len(vals) > 1:
                        original_url = tooltip_handler.original_values.get((iid, 1))
                        if original_url:
                            vals[1] = shorten_url_for_export(original_url, export_type)
                            tooltip_row.append((1, original_url))  # Column 1, full URL
                            
                elif tree_type == "Bookmarks" and len(vals) > 2:
                    original_url = tooltip_handler.original_values.get((iid, 2))
                    if original_url:
                        vals[2] = shorten_url_for_export(original_url, export_type)
                        tooltip_row.append((2, original_url))  # Column 2, full URL
            
            rows.append(vals)
            tooltip_data.append(tooltip_row)
        
        return headers, rows, tooltip_data

    def _rows_to_html_table(self, headers, rows):
        parts = []
        parts.append('<table>')
        parts.append('<thead><tr>')
        for h in headers:
            parts.append(f"<th>{self._html_escape(str(h))}</th>")
        parts.append('</tr></thead>')
        parts.append('<tbody>')
        for r in rows:
            parts.append('<tr>')
            for c in r:
                parts.append(f"<td>{self._html_escape(str(c))}</td>")
            parts.append('</tr>')
        parts.append('</tbody>')
        parts.append('</table>')
        return "\n".join(parts)

    def _rows_to_html_table_with_tooltips(self, headers, rows, tooltip_data):
        """Create HTML table with tooltips for shortened URLs"""
        parts = []
        parts.append('<table>')
        parts.append('<thead><tr>')
        for h in headers:
            parts.append(f"<th>{self._html_escape(str(h))}</th>")
        parts.append('</tr></thead>')
        parts.append('<tbody>')
        
        for row_idx, r in enumerate(rows):
            parts.append('<tr>')
            tooltip_row = tooltip_data[row_idx] if row_idx < len(tooltip_data) else []
            tooltip_dict = {col_idx: full_value for col_idx, full_value in tooltip_row}
            
            for col_idx, c in enumerate(r):
                cell_content = self._html_escape(str(c))
                
                # Add tooltip if this cell has a full URL/value
                if col_idx in tooltip_dict:
                    full_value = self._html_escape(str(tooltip_dict[col_idx]))
                    parts.append(f'<td title="{full_value}" style="cursor: help;">{cell_content}</td>')
                else:
                    parts.append(f"<td>{cell_content}</td>")
            parts.append('</tr>')
        parts.append('</tbody>')
        parts.append('</table>')
        return "\n".join(parts)

    def _html_escape(self, s):
        return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                .replace('"', '&quot;').replace("'", '&#39;'))

    def save_all_as_single_html(self):
        # Saves everything (Overview, Browsing History, Downloads, Cookies, Bookmarks)
        start_date, end_date = self.get_selected_date_range()
        default_name = f"forensic_full_export_{date_only_ddmmyy(start_date).replace('/', '-')}_to_{date_only_ddmmyy(end_date).replace('/', '-')}.html"
        path = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("HTML files", "*.html")], initialfile=default_name)
        if not path:
            return

        try:
            sections = []  # tuples of (id, title, html)

            # Overview
            hdrs, rows = self._get_tree_data(self.overview_tree)
            overview_html = '<h2 id="overview">Installed Browsers</h2>' + self._rows_to_html_table(hdrs, rows)
            sections.append(("overview", "Installed Browsers", overview_html))

            # For each data tab
            for tab_name, tree in self.tree_views.items():
                hdrs, rows, tooltip_data = self._get_tree_data_with_tooltips(tree, export_type="html")
                section_id = tab_name.lower().replace(' ', '-')
                title_html = f'<h2 id="{section_id}">{self._html_escape(tab_name)}</h2>'
                if rows:
                    table_html = self._rows_to_html_table_with_tooltips(hdrs, rows, tooltip_data)
                else:
                    table_html = '<p><em>No items in this view.</em></p>'
                # For Bookmarks, add quick stats above table
                if tab_name == "Bookmarks":
                    # compute stats from master
                    total = len(self.bookmarks_master)
                    dupes = sum(1 for r in self.bookmarks_master if r.get('dup'))
                    domains = {}
                    for r in self.bookmarks_master:
                        d = r.get('domain')
                        if d:
                            domains[d] = domains.get(d, 0) + 1
                    top_domains = sorted(domains.items(), key=lambda x: x[1], reverse=True)[:15]
                    stats = [f"<p><strong>Total:</strong> {total} &nbsp; <strong>Duplicates:</strong> {dupes}</p>"]
                    stats.append("<h3>Top Domains</h3><ol>")
                    for d, c in top_domains:
                        stats.append(f"<li>{self._html_escape(d)} — {c}</li>")
                    stats.append("</ol>")
                    table_html = "".join(stats) + table_html
                sections.append((section_id, tab_name, title_html + table_html))

            # Build Table of Contents
            toc_parts = ['<nav><h2>Table of Contents</h2><ul>']
            for sid, title, _ in sections:
                toc_parts.append(f'<li><a href="#${sid}">{self._html_escape(title)}</a></li>')
            toc_parts.append('</ul></nav><hr/>')
            toc_html = '\n'.join(toc_parts)

            # NOTE: anchor href uses #id — but to ensure compatibility with some viewers, also include a top anchor
            full_parts = []
            full_parts.append('<a id="top"></a>')
            full_parts.append(f'<h1>Browser Forensic Full Export - Indian Cybercrime Coordination Centre (I4C)</h1>')
            full_parts.append(f'<p><strong>Start:</strong> {date_only_ddmmyy(start_date)} &nbsp;&nbsp; <strong>End:</strong> {date_only_ddmmyy(end_date)}</p>')
            full_parts.append('<div style="background-color: #f0f8ff; border: 1px solid #0066cc; padding: 10px; margin: 10px 0; border-radius: 5px;">')
            full_parts.append('<p><strong>📝 Note:</strong> URLs in this report have been shortened for better readability. Hover over any shortened URL to see the complete URL in a tooltip.</p>')
            full_parts.append('</div>')
            full_parts.append(toc_html)

            for sid, title, html in sections:
                full_parts.append(html)
                full_parts.append('<p><a href="#top">Back to top</a></p>')
                full_parts.append('<hr/>')

            # Add IFSO-NCFL watermark
            ifso_watermark = """
            <div style="position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%) rotate(-45deg); 
                        font-size: 72px; color: rgba(255, 0, 0, 0.1); font-weight: bold; z-index: -1; 
                        pointer-events: none; user-select: none;">
                IFSO-NCFL
            </div>
            """
            
            full_html = """
            <!doctype html>
            <html lang="en">
            <head>
            <meta charset="utf-8" />
            <meta name="viewport" content="width=device-width,initial-scale=1" />
            <title>Forensic Full Export - IFSO-NCFL</title>
            <style>
                body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial; padding: 20px; position: relative; }
                nav ul { list-style: none; padding-left: 0; }
                nav li { margin: 6px 0; }
                table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
                th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
                th { background: #f3f4f6; }
                h1, h2 { color: #0f172a; }
                a { color: #0b4f6c; }
                .ifso-watermark {
                    position: fixed;
                    top: 50%;
                    left: 50%;
                    transform: translate(-50%, -50%) rotate(-45deg);
                    font-size: 72px;
                    color: rgba(255, 0, 0, 0.1);
                    font-weight: bold;
                    z-index: -1;
                    pointer-events: none;
                    user-select: none;
                }
                .ncfl-trademark {
                    position: fixed;
                    bottom: 10px;
                    right: 10px;
                    font-size: 24px;
                    color: rgba(0, 0, 0, 0.8);
                    font-weight: bold;
                    z-index: 1000;
                    pointer-events: none;
                    user-select: none;
                    background: rgba(255, 255, 255, 0.9);
                    padding: 5px 10px;
                    border-radius: 5px;
                    border: 2px solid #000;
                }
                @media print {
                    .ifso-watermark {
                        position: absolute;
                        top: 50%;
                        left: 50%;
                        transform: translate(-50%, -50%) rotate(-45deg);
                        font-size: 72px;
                        color: rgba(255, 0, 0, 0.15);
                        font-weight: bold;
                        z-index: -1;
                    }
                    .ncfl-trademark {
                        position: fixed;
                        bottom: 10px;
                        right: 10px;
                        font-size: 24px;
                        color: rgba(0, 0, 0, 0.8);
                        font-weight: bold;
                        z-index: 1000;
                        background: rgba(255, 255, 255, 0.9);
                        padding: 5px 10px;
                        border-radius: 5px;
                        border: 2px solid #000;
                    }
                }
            </style>
            </head>
            <body>
            <div class="ifso-watermark">IFSO-NCFL</div>
            <div class="ncfl-trademark">Indian Cybercrime Coordination Centre (I4C)</div>
            """ + "\n".join(full_parts) + "\n</body>\n</html>"

            # Fix TOC hrefs
            full_html = full_html.replace('#$', '#')

            with open(path, "w", encoding="utf-8") as f:
                f.write(full_html)

            messagebox.showinfo("Saved", f"Full report saved to:\n{path}")
            self.status_var.set(f"Saved full export as HTML: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save full HTML: {e}")

    def save_all_as_csv(self):
        """Export Overview and all tabs as CSV files into a chosen folder"""
        try:
            folder = filedialog.askdirectory(title="Select folder to save CSVs")
            if not folder:
                return
            # Overview
            hdrs, rows = self._get_tree_data(self.overview_tree)
            self._write_csv(os.path.join(folder, "overview.csv"), hdrs, rows)
            # Tabs
            for tab_name, tree in self.tree_views.items():
                hdrs, rows = self._get_tree_data(tree, use_full_urls=True)
                fname = tab_name.lower().replace(' ', '_') + ".csv"
                self._write_csv(os.path.join(folder, fname), hdrs, rows)
            messagebox.showinfo("Saved", f"CSV files saved to:\n{folder}")
            self.status_var.set(f"Saved CSVs to: {folder}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save CSVs: {e}")

    def _write_csv(self, path, headers, rows):
        import csv
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(headers)
            for r in rows:
                w.writerow(r)

    def save_all_as_pdf(self):
        """Export a comprehensive PDF report with clickable Table of Contents and complete navigation - Enhanced to match HTML save functionality"""
        try:
            from reportlab.lib.pagesizes import A4, landscape, A3
            from reportlab.lib import colors
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
            from reportlab.platypus.tableofcontents import TableOfContents
            from reportlab.lib.units import inch, cm
            from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
        except Exception as e:
            messagebox.showwarning("Import Error", f"Failed to import reportlab modules:\n{str(e)}\n\nPlease ensure reportlab is properly installed:\npip install reportlab")
            return

        # Enhanced file dialog with same naming convention as HTML export
        start_date, end_date = self.get_selected_date_range()
        default_name = f"forensic_full_export_{date_only_ddmmyy(start_date).replace('/', '-')}_to_{date_only_ddmmyy(end_date).replace('/', '-')}.pdf"
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf", 
            filetypes=[("PDF files", "*.pdf")], 
            initialfile=default_name
        )
        if not path:
            return

        try:
            from reportlab.platypus import PageTemplate, BaseDocTemplate, Frame
            from reportlab.lib.colors import Color
            
            # Custom page template with Indian Cybercrime Coordination Centre (I4C) branding
            class NCFLPageTemplate(PageTemplate):
                def __init__(self, id, frames, **kwargs):
                    PageTemplate.__init__(self, id, frames, **kwargs)
                
                def beforeDrawPage(self, canvas, doc):
                    # Get page dimensions
                    page_width, page_height = landscape(A3)
                    
                    # Add IFSO-NCFL watermark
                    canvas.saveState()
                    canvas.setFillColor(Color(1, 0, 0, alpha=0.1))  # Red with transparency
                    canvas.setFont("Helvetica-Bold", 120)
                    
                    # Center the watermark and rotate
                    canvas.translate(page_width/2, page_height/2)
                    canvas.rotate(45)
                    
                    # Draw IFSO-NCFL text
                    text_width = canvas.stringWidth("IFSO-NCFL", "Helvetica-Bold", 120)
                    canvas.drawString(-text_width/2, -60, "IFSO-NCFL")
                    
                    canvas.restoreState()
                    
                    # Add Indian Cybercrime Coordination Centre (I4C) trademark in bottom right corner
                    canvas.saveState()
                    canvas.setFillColor(Color(0, 0, 0, alpha=0.8))  # Black with transparency
                    canvas.setFont("Helvetica-Bold", 12)
                    
                    # Draw border box for trademark
                    trademark_text = "Indian Cybercrime Coordination Centre (I4C)"
                    text_width = canvas.stringWidth(trademark_text, "Helvetica-Bold", 12)
                    box_width = text_width + 20
                    box_height = 40
                    
                    # Position in bottom right
                    x_pos = page_width - box_width - 20
                    y_pos = 20
                    
                    # Draw white background box with border
                    canvas.setFillColor(Color(1, 1, 1, alpha=0.9))  # White background
                    canvas.setStrokeColor(Color(0, 0, 0))  # Black border
                    canvas.setLineWidth(2)
                    canvas.roundRect(x_pos, y_pos, box_width, box_height, 5, fill=1, stroke=1)
                    
                    # Draw trademark text
                    canvas.setFillColor(Color(0, 0, 0, alpha=0.8))  # Black text
                    canvas.drawString(x_pos + 10, y_pos + 10, trademark_text)
                    
                    canvas.restoreState()
            
            # Use custom BaseDocTemplate with Indian Cybercrime Coordination Centre (I4C) branding
            doc = BaseDocTemplate(
                path, 
                pagesize=landscape(A3), 
                title="Browser Forensic Report with Clickable TOC - Indian Cybercrime Coordination Centre (I4C)",
                leftMargin=0.3*inch,
                rightMargin=0.3*inch,
                topMargin=0.4*inch,
                bottomMargin=0.4*inch
            )
            
            # Create frame and add custom page template
            frame = Frame(0.3*inch, 0.4*inch, 
                         landscape(A3)[0] - 0.6*inch, 
                         landscape(A3)[1] - 0.8*inch)
            
            template = NCFLPageTemplate('normal', [frame])
            doc.addPageTemplates([template])
            
            styles = getSampleStyleSheet()
            
            # Enhanced custom styles for better visibility
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Title'],
                fontSize=22,
                spaceAfter=30,
                alignment=TA_CENTER,
                textColor=colors.HexColor('#2c3e50'),
                fontName='Helvetica-Bold'
            )
            
            heading_style = ParagraphStyle(
                'CustomHeading',
                parent=styles['Heading2'],
                fontSize=18,
                spaceAfter=15,
                spaceBefore=25,
                textColor=colors.HexColor('#34495e'),
                borderWidth=2,
                borderColor=colors.HexColor('#3498db'),
                borderPadding=10,
                backColor=colors.HexColor('#ecf0f1'),
                fontName='Helvetica-Bold'
            )
            
            normal_style = ParagraphStyle(
                'CustomNormal',
                parent=styles['Normal'],
                fontSize=11,
                spaceAfter=8,
                alignment=TA_JUSTIFY
            )
            
            # Enhanced TOC style
            toc_style = ParagraphStyle( 
                'TOCHeading',
                parent=styles['Normal'],
                fontSize=16,
                spaceAfter=15,
                spaceBefore=25,
                textColor=colors.HexColor('#2c3e50'),
                alignment=TA_CENTER,
                fontName='Helvetica-Bold',
                borderWidth=2,
                borderColor=colors.HexColor('#e74c3c'),
                borderPadding=10,
                backColor=colors.HexColor('#ecf0f1')
            )
            
            story = []
            
            # Title page with enhanced information
            story.append(Paragraph("Browser Forensic Analysis Report - Indian Cybercrime Coordination Centre (I4C)", title_style))
            story.append(Spacer(1, 20))
            
            start_date, end_date = self.get_selected_date_range()
            metadata_info = f"""
            <b>Analysis Period:</b> {date_only_ddmmyy(start_date)} to {date_only_ddmmyy(end_date)}<br/>
            <b>Report Generated on:</b> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}<br/>
            <b>System Platform:</b> {get_platform().upper()}<br/>
            <b>Intellectual Property:</b> National Cyber Forensic Lab I4C, MHA<br/>
            <b>Report Type:</b> PDF<br/>
            """
            story.append(Paragraph(metadata_info, normal_style))
            story.append(Spacer(1, 40))
            story.append(PageBreak())
             
            # Create fully clickable Table of Contents with enhanced navigation
            story.append(Paragraph('<a name="table_of_contents"/>', normal_style))
            story.append(Paragraph("Table of Contents", toc_style))
            story.append(Spacer(1, 15))
            
            # Create manual clickable TOC with better navigation
            toc_style_main = ParagraphStyle(
                'TOCMain',
                fontSize=14,
                spaceAfter=8,
                textColor=colors.HexColor('#2c3e50'),
                fontName='Helvetica-Bold',
                leftIndent=20
            )
            
            # Collect section information for manual TOC
            toc_sections = []
            
            # Add Executive Summary to TOC list
            toc_sections.append(("Executive Summary", "executive_summary"))
            
            # Add Top Visited URLs to TOC if data exists
            if hasattr(self, 'tree_views') and 'Browsing History' in self.tree_views:
                top_urls_data = self._get_top_visited_urls()
                if top_urls_data:
                    toc_sections.append(("Top Visited URLs", "top_visited_urls"))

            # Add main sections to TOC list
            for tab_name in ["Browsing History", "Downloads", "Cookies", "Bookmarks", "Secrets"]:
                if tab_name in self.tree_views and self.tree_views[tab_name].get_children():
                    clean_name = tab_name.lower().replace(" ", "_").replace("&", "and").replace("-", "_")
                    toc_sections.append((tab_name, clean_name))
            
            # Add Browsers Overview if available
            if hasattr(self, 'overview_tree') and self.overview_tree.get_children():
                toc_sections.append(("Installed Browsers", "installed_browsers"))
            
            # Add Annexure
            toc_sections.append(("Annexure", "report_information"))
            
            # Create clickable TOC entries with enhanced styling
            for i, (section_title, anchor) in enumerate(toc_sections):
                # Add some visual separation
                if i > 0:
                    story.append(Spacer(1, 4))
                
                # Create clickable link with enhanced styling
                toc_link = f'<a href="#{anchor}" color="#2980b9"><u>{section_title}</u></a>'
                story.append(Paragraph(toc_link, toc_style_main))
            
            # Add a separator line after TOC
            story.append(Spacer(1, 15))
            separator_style = ParagraphStyle(
                'Separator',
                fontSize=12,
                textColor=colors.HexColor('#bdc3c7'),
                alignment=TA_CENTER
            )
            story.append(Paragraph("─" * 80, separator_style))
            
            story.append(Spacer(1, 25))
            
            # Add note about URL shortening
            note_style = ParagraphStyle(
                'NoteStyle',
                parent=normal_style,
                fontSize=12,
                spaceAfter=15,
                spaceBefore=10,
                textColor=colors.HexColor('#2c3e50'),
                borderWidth=1,
                borderColor=colors.HexColor('#3498db'),
                borderPadding=10,
                backColor=colors.HexColor('#f0f8ff'),
                fontName='Helvetica'
            )
            
            story.append(Paragraph("📝 <b>Note:</b> URLs in this report have been shortened for better readability and consistent formatting. This matches the display format shown in the GUI interface.", note_style))
            story.append(Spacer(1, 15))
            
            story.append(PageBreak())

            def smart_text_wrap(text, max_length=120):
                """Smart text wrapping that preserves readability"""
                if not text:
                    return ""
                text_str = str(text).strip()
                
                # Don't truncate short text
                if len(text_str) <= max_length:
                    return text_str
                
                # For URLs, show beginning and end
                if any(keyword in text_str.lower() for keyword in ['http', 'www', '.com', '.org', '.net']):
                    if len(text_str) > max_length:
                        mid_point = max_length // 2 - 5
                        return f"{text_str[:mid_point]}...{text_str[-mid_point:]}"
                
                # For other text, try to break at word boundaries
                if ' ' in text_str and len(text_str) > max_length:
                    words = text_str.split()
                    result = ""
                    for word in words:
                        if len(result + " " + word) <= max_length - 3:
                            result += (" " if result else "") + word
                        else:
                            result += "..."
                            break
                    return result
                
                # Fallback truncation
                return text_str[:max_length-3] + "..."

            def add_enhanced_table(title, headers, rows, tooltip_data=None):
                # Create clean anchor name for navigation
                clean_title = title.replace("🌐 ", "").replace("📥 ", "").replace("🍪 ", "").replace("🔖 ", "").replace("🔐 ", "").replace("📊 ", "").replace("📋 ", "").replace("📄 ", "").strip()
                anchor_name = clean_title.lower().replace(" ", "_").replace("&", "and").replace("-", "_")
                
                # Add invisible anchor for TOC navigation
                story.append(Paragraph(f'<a name="{anchor_name}"/>', normal_style))
                
                # Anchor is already added above for navigation
                
                # Create section-specific styling based on content type
                section_colors = {
                    "Browser History": colors.HexColor('#3498db'),
                    "Browsing History": colors.HexColor('#3498db'),
                    "Downloads": colors.HexColor('#e67e22'),
                    "Cookies": colors.HexColor('#f39c12'),
                    "Bookmarks": colors.HexColor('#9b59b6'),
                    "Secrets": colors.HexColor('#e74c3c'),
                    "Executive Summary": colors.HexColor('#27ae60'),
                    "Installed Browsers": colors.HexColor('#34495e')
                }
                
                # Get color for this section
                section_color = section_colors.get(clean_title, colors.HexColor('#3498db'))
                
                # Enhanced section heading with color coding and back-to-TOC link
                section_heading = ParagraphStyle(
                    'SectionHeading',
                    parent=heading_style,
                    fontSize=20,
                    spaceAfter=20,
                    spaceBefore=30,
                    textColor=colors.HexColor('#2c3e50'),
                    borderWidth=3,
                    borderColor=section_color,
                    borderPadding=12,
                    backColor=colors.HexColor('#f8f9fa'),
                    fontName='Helvetica-Bold'
                )
                
                # Add section title with back-to-TOC link
                title_with_nav = f'{title} &nbsp;&nbsp;&nbsp; <a href="#table_of_contents" color="blue"><font size="10">↑ Back to Contents</font></a>'
                story.append(Paragraph(title_with_nav, section_heading))
                
                if not rows:
                    story.append(Paragraph("ℹ️ No data available for this section", normal_style))
                    story.append(Spacer(1, 20))
                    return
                
                # Show more rows for complete visibility
                max_rows = 2000
                view = rows[:max_rows]
                
                # Enhanced text processing for complete visibility
                processed_rows = []
                for row in view:
                    processed_row = []
                    for i, cell in enumerate(row):
                        cell_str = str(cell) if cell is not None else ""
                        
                        # Different handling based on column type
                        if i == 0 or 'url' in headers[i].lower() if i < len(headers) else False:
                            # URLs and first columns get more space
                            processed_row.append(smart_text_wrap(cell_str, 150))
                        elif 'password' in headers[i].lower() if i < len(headers) else False:
                            # Passwords get special handling
                            if cell_str and cell_str != "*" * len(cell_str):
                                processed_row.append(smart_text_wrap(cell_str, 100))
                            else:
                                processed_row.append(cell_str)
                        else:
                            # Other columns
                            processed_row.append(smart_text_wrap(cell_str, 80))
                    processed_rows.append(processed_row)
                
                # Prepare table data with enhanced headers
                enhanced_headers = headers
                table_data = [enhanced_headers] + processed_rows
                
                # Calculate optimal column widths for A3 landscape
                page_width = landscape(A3)[0] - 0.6*inch  # Account for margins
                num_cols = len(headers)
                
                if num_cols > 0:
                    # Optimized column widths for complete text visibility
                    col_widths = []
                    if title == "Installed Browsers":
                        col_widths = [page_width * 0.25, page_width * 0.25, page_width * 0.25, page_width * 0.25]
                    elif title == "Browsing History":
                        col_widths = [page_width * 0.45, page_width * 0.45, page_width * 0.1]
                    elif title == "Downloads":
                        col_widths = [page_width * 0.3, page_width * 0.5, page_width * 0.1, page_width * 0.1]
                    elif title == "Cookies":
                        col_widths = [page_width * 0.35, page_width * 0.35, page_width * 0.15, page_width * 0.15]
                    elif title == "Bookmarks":
                        col_widths = [page_width * 0.12, page_width * 0.25, page_width * 0.35, page_width * 0.08, 
                                    page_width * 0.08, page_width * 0.04, page_width * 0.04, page_width * 0.04]
                    elif title == "Secrets":
                        col_widths = [page_width * 0.25, page_width * 0.15, page_width * 0.25, page_width * 0.15,
                                    page_width * 0.08, page_width * 0.08, page_width * 0.04]
                    else:
                        # Default optimized widths
                        col_widths = [page_width / num_cols] * num_cols
                    
                    # Ensure proper column count
                    while len(col_widths) < num_cols:
                        col_widths.append(page_width / num_cols)
                    col_widths = col_widths[:num_cols]
                    
                    tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
                else:
                    tbl = Table(table_data, repeatRows=1)
                
                # Professional table styling with complete visibility
                tbl.setStyle(TableStyle([
                    # Enhanced header styling
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#34495e')),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0,0), (-1,0), 11),
                    ('ALIGN', (0,0), (-1,0), 'CENTER'),
                    
                    # Enhanced data rows styling
                    ('BACKGROUND', (0,1), (-1,-1), colors.white),
                    ('TEXTCOLOR', (0,1), (-1,-1), colors.black),
                    ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
                    ('FONTSIZE', (0,1), (-1,-1), 9),
                    ('ALIGN', (0,1), (-1,-1), 'LEFT'),
                    
                    # Professional grid and borders
                    ('GRID', (0,0), (-1,-1), 0.8, colors.HexColor('#95a5a6')),
                    ('LINEBELOW', (0,0), (-1,0), 3, colors.HexColor('#34495e')),
                    ('LINEBEFORE', (0,0), (0,-1), 2, colors.HexColor('#34495e')),
                    ('LINEAFTER', (-1,0), (-1,-1), 2, colors.HexColor('#34495e')),
                    
                    # Enhanced alternating row colors
                    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f8f9fa')]),
                    
                    # Optimized padding for readability
                    ('LEFTPADDING', (0,0), (-1,-1), 6),
                    ('RIGHTPADDING', (0,0), (-1,-1), 6),
                    ('TOPPADDING', (0,0), (-1,-1), 8),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 8),
                    
                    # Text alignment and wrapping
                    ('VALIGN', (0,0), (-1,-1), 'TOP'),
                    ('WORDWRAP', (0,0), (-1,-1), 'LTR'),
                ]))
                
                story.append(tbl)
                
                # Add full URLs section if tooltip data is available
                if tooltip_data:
                    full_urls = []
                    row_num = 1
                    for row_tooltips in tooltip_data:
                        if row_tooltips:  # If this row has tooltip data
                            for col_idx, full_url in row_tooltips:
                                if full_url and len(full_url) > 60:  # Only show if URL was actually shortened
                                    col_name = headers[col_idx] if col_idx < len(headers) else f"Column {col_idx+1}"
                                    full_urls.append(f"Row {row_num}, {col_name}: {full_url}")
                        row_num += 1
                    
                    if full_urls:
                        story.append(Spacer(1, 15))
                        story.append(Paragraph("📋 Full URLs (for shortened entries above):", 
                                             ParagraphStyle('FullURLsHeading', parent=styles['Heading3'], 
                                                          fontSize=12, textColor=colors.HexColor('#2c3e50'),
                                                          spaceAfter=8)))
                        
                        # Create a simple table for full URLs
                        url_table_data = [["Row/Column", "Full URL"]]
                        for i, url_info in enumerate(full_urls[:50]):  # Limit to 50 URLs to avoid overwhelming
                            parts = url_info.split(": ", 1)
                            if len(parts) == 2:
                                url_table_data.append([parts[0], parts[1]])
                        
                        if len(url_table_data) > 1:
                            url_table = Table(url_table_data, colWidths=[page_width * 0.25, page_width * 0.75])
                            url_table.setStyle(TableStyle([
                                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#34495e')),
                                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                                ('FONTSIZE', (0,0), (-1,0), 10),
                                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                                ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
                                ('FONTSIZE', (0,1), (-1,-1), 8),
                                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#bdc3c7')),
                                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f8f9fa')]),
                                ('LEFTPADDING', (0,0), (-1,-1), 4),
                                ('RIGHTPADDING', (0,0), (-1,-1), 4),
                                ('TOPPADDING', (0,0), (-1,-1), 4),
                                ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                                ('WORDWRAP', (0,0), (-1,-1), 'LTR'),
                            ]))
                            story.append(url_table)
                            
                            if len(full_urls) > 50:
                                story.append(Spacer(1, 5))
                                story.append(Paragraph(f"Note: Showing first 50 of {len(full_urls)} full URLs", 
                                                     ParagraphStyle('URLNote', parent=styles['Italic'], 
                                                                  fontSize=8, textColor=colors.HexColor('#7f8c8d'))))
                
                # Enhanced statistics
                if len(rows) > max_rows:
                    story.append(Spacer(1, 10))
                    story.append(Paragraph(f"📈 Statistics: Displaying first {max_rows:,} of {len(rows):,} total entries ({(max_rows/len(rows)*100):.1f}% shown)", 
                                         ParagraphStyle('Stats', parent=styles['Italic'], fontSize=10, textColor=colors.HexColor('#7f8c8d'))))
                else:
                    story.append(Spacer(1, 10))
                    story.append(Paragraph(f"📈 Complete Dataset: {len(rows):,} entries displayed", 
                                         ParagraphStyle('Stats', parent=styles['Italic'], fontSize=10, textColor=colors.HexColor('#27ae60'))))
                
                story.append(Spacer(1, 25))

            # Enhanced summary information with proper TOC entry
            story.append(Paragraph('<a name="executive_summary"/>', normal_style))
            # Executive Summary anchor already added above
            
            summary_heading = ParagraphStyle(
                'SummaryHeading',
                parent=heading_style,
                fontSize=18,
                spaceAfter=15,
                spaceBefore=25,
                textColor=colors.HexColor('#2c3e50'),
                borderWidth=2,
                borderColor=colors.HexColor('#e74c3c'),
                borderPadding=10,
                backColor=colors.HexColor('#ecf0f1')
            )
            story.append(Paragraph("Executive Summary", summary_heading))
            total_entries = 0
            section_counts = {}
            for tab_name, tree in self.tree_views.items():
                count = len(tree.get_children())
                total_entries += count
                section_counts[tab_name] = count
            
            summary_info = f"""
            <b>Total Data Entries:</b> {total_entries:,}<br/>
            <b>Analysis Period:</b> {date_only_ddmmyy(start_date)} to {date_only_ddmmyy(end_date)}<br/>
            <b>Report Generated:</b> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}<br/>
            """
            
            story.append(Paragraph(summary_info, normal_style))
            story.append(Spacer(1, 30))

            # Add all sections with enhanced formatting
            # Overview section
            hdrs, rows = self._get_tree_data(self.overview_tree)
            add_enhanced_table("Installed Browsers", hdrs, rows, None)

            # Add Top Visited URLs section before browsing history
            if hasattr(self, 'tree_views') and 'Browsing History' in self.tree_views:
                top_urls_data = self._get_top_visited_urls()
                if top_urls_data:
                    add_enhanced_table("Top Visited URLs", ["URL", "Title", "Visit Count", "Last Visit"], top_urls_data, None)

            # Each tab data with page breaks and enhanced organization
            section_config = {
                "Browsing History": {"icon": "", "description": "Web browsing activity and visited URLs"},
                "Downloads": {"icon": "", "description": "Downloaded files and their sources"}, 
                "Cookies": {"icon": "", "description": "Browser cookies and tracking data"},
                "Bookmarks": {"icon": "", "description": "Saved bookmarks and favorites"},
                "Secrets": {"icon": "", "description": "Stored passwords and authentication data"}
            }
            
            for tab_name, tree in self.tree_views.items():
                story.append(PageBreak())
                
                # Get section configuration
                config = section_config.get(tab_name, {"icon": "", "description": "Data analysis section"})
                icon = config["icon"]
                description = config["description"]
                
                # Create enhanced title with icon
                enhanced_title = f"{icon} {tab_name}"
                
                # Add section description before the table
                story.append(Paragraph(f"<i>{description}</i>", 
                    ParagraphStyle('SectionDesc', parent=normal_style, fontSize=10, 
                                 textColor=colors.HexColor('#7f8c8d'), alignment=TA_CENTER)))
                story.append(Spacer(1, 10))
                
                # Special handling for Bookmarks section - add statistics like HTML version
                if tab_name == "Bookmarks" and hasattr(self, 'bookmarks_master'):
                    total = len(self.bookmarks_master)
                    dupes = sum(1 for r in self.bookmarks_master if r.get('dup'))
                    domains = {}
                    for r in self.bookmarks_master:
                        d = r.get('domain')
                        if d:
                            domains[d] = domains.get(d, 0) + 1
                    top_domains = sorted(domains.items(), key=lambda x: x[1], reverse=True)[:15]
                    
                    # Add bookmark statistics
                    stats_text = f"<b>Total Bookmarks:</b> {total} &nbsp;&nbsp; <b>Duplicates:</b> {dupes}<br/><br/><b>Top Domains:</b><br/>"
                    for i, (domain, count) in enumerate(top_domains, 1):
                        stats_text += f"{i}. {domain} — {count}<br/>"
                    
                    story.append(Paragraph(stats_text, normal_style))
                    story.append(Spacer(1, 15))
                
                hdrs, rows, tooltip_data = self._get_tree_data_with_tooltips(tree, export_type="pdf")
                add_enhanced_table(enhanced_title, hdrs, rows, tooltip_data)

            # Add footer information with proper TOC entry
            story.append(PageBreak())
            story.append(Paragraph('<a name="report_information"/>', normal_style))
            # Report Information anchor already added above
            
            footer_heading = ParagraphStyle(
                'FooterHeading',
                parent=heading_style,
                fontSize=18,
                spaceAfter=15,
                spaceBefore=25,
                textColor=colors.HexColor('#2c3e50'),
                borderWidth=2,
                borderColor=colors.HexColor('#27ae60'),
                borderPadding=10,
                backColor=colors.HexColor('#ecf0f1')
            )
            story.append(Paragraph("Annexure", footer_heading))
            
            # Provider Statement Section
            provider_statement = """
            <b>Provider Statement and User Acknowledgment:</b><br/><br/>
            The Provider expressly states and the User acknowledges that:<br/><br/>
            
            <b>a)</b> The Browser Forensic Tool is not listed under the Government e-Marketplace (GeM) repository of forensic tools maintained by the Government of India.<br/><br/>
            
            <b>b)</b> The Provider is not an empaneled or accredited digital forensic laboratory under Section 79A of the Information Technology Act, 2000, or any related notification issued thereunder.<br/><br/>
            """
            story.append(Paragraph(provider_statement, normal_style))
            
            # User Agreement Section
            user_agreement = """
            <b>User Agreement and Undertakings:</b><br/><br/>
            The User hereby agrees and undertakes:<br/><br/>
            
            <b>1.</b> To utilize the Browser Forensic Tool at its own discretion, risk, and responsibility.<br/><br/>
            
            <b>2.</b> To indemnify, defend, and hold harmless the Provider, its directors, employees, representatives, and affiliates from and against any and all claims, liabilities, damages, penalties, costs, and expenses (including legal fees) arising out of or in connection with:<br/><br/>
            
            &nbsp;&nbsp;&nbsp;&nbsp;<b>(a)</b> The use of the Browser Forensic Tool in any investigation or judicial proceeding;<br/><br/>
            
            &nbsp;&nbsp;&nbsp;&nbsp;<b>(b)</b> Any reliance placed on the results, reports, or outputs generated by the tool;<br/><br/>
            
            &nbsp;&nbsp;&nbsp;&nbsp;<b>(c)</b> Any third-party challenge, objection, or legal scrutiny concerning admissibility of evidence derived therefrom.<br/><br/>
            
            <b>3.</b> Report exported from this tool shall only be used for investigative purposes and is not admissible in the court of law.<br/><br/>
            """
            story.append(Paragraph(user_agreement, normal_style))
            
            # Important Notice Section
            important_notice = """
            <b>IMPORTANT NOTICE:</b><br/><br/>
            This tool is provided for educational and investigative purposes only. Users are advised to consult with legal experts and follow proper chain of custody procedures when handling digital evidence. The Provider disclaims all warranties and shall not be liable for any consequences arising from the use of this tool.
            """
            story.append(Paragraph(important_notice, normal_style))

            # Build the enhanced PDF
            doc.build(story)
            
            # Enhanced success message matching HTML save functionality
            messagebox.showinfo("Saved", f"Full PDF report saved to:\n{path}")
            self.status_var.set(f"Saved full export as PDF: {os.path.basename(path)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save full PDF: {e}")

    def show_credits(self):
        """Show About/Help menu for i button."""
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="About", command=self.show_about)
        menu.add_command(label="Help", command=self.show_help)
        # Place menu near mouse pointer
        try:
            x = self.root.winfo_pointerx() - self.root.winfo_rootx()
            y = self.root.winfo_pointery() - self.root.winfo_rooty()
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def show_about(self):
        about_text = (
            "Developer:\n  Mr. Anand Raj Chopra\n  Email: anandrajchopra12@gmail.com\n\n"
        )
        messagebox.showinfo("About", about_text)

    def show_help(self):
        # Open Documentation.txt in default editor
        import subprocess, sys
        doc_path = os.path.join(os.path.dirname(__file__) if '__file__' in globals() else os.getcwd(), "Browser_Forensic_Tool_Documentation.docx")
        try:
            if sys.platform.startswith('win'):
                os.startfile(doc_path)
            elif sys.platform.startswith('darwin'):
                subprocess.Popen(['open', doc_path])
            else:
                subprocess.Popen(['xdg-open', doc_path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open help file:\n{e}")

    # Password-related methods - REMOVED
    
    # Password display refresh - REMOVED
    
    # Password filters - REMOVED
    
    def _show_password_menu(self, event, tree):
        """Show context menu for password entries"""
        try:
            self.pwd_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.pwd_menu.grab_release()
    
    def _password_copy(self, tree, which="username"):
        """Copy password data to clipboard"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a password entry first.")
            return
        
        item = selection[0]
        values = tree.item(item, "values")
        
        if which == "username" and len(values) > 1:
            self.root.clipboard_clear()
            self.root.clipboard_append(values[1])
            messagebox.showinfo("Copied", "Username copied to clipboard")
        elif which == "password" and len(values) > 2:
            # Get actual password, not masked version
            url = values[0]
            username = values[1]
            actual_password = self._get_actual_password(url, username)
            if actual_password:
                self.root.clipboard_clear()
                self.root.clipboard_append(actual_password)
                messagebox.showinfo("Copied", "Password copied to clipboard")
            else:
                messagebox.showwarning("No Password", "No password found for this entry")
        elif which == "url" and len(values) > 0:
            self.root.clipboard_clear()
            self.root.clipboard_append(values[0])
            messagebox.showinfo("Copied", "URL copied to clipboard")
    
    def _get_actual_password(self, url, username):
        """Get the actual decrypted password for a given URL and username"""
        if hasattr(self, 'all_data'):
            for browser_name, browser_data in self.all_data.items():
                if 'passwords' in browser_data:
                    for pwd_data in browser_data['passwords']:
                        if (pwd_data.get('url', '') == url and 
                            pwd_data.get('username', '') == username):
                            return pwd_data.get('password_decrypted', '')
        return ""
    
    def _show_encryption_details(self, tree):
        """Show detailed encryption information for selected password"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a password entry first.")
            return
        
        item = selection[0]
        values = tree.item(item, "values")
        url = values[0]
        username = values[1]
        
        # Find the password data
        pwd_data = None
        if hasattr(self, 'all_data'):
            for browser_name, browser_data in self.all_data.items():
                if 'passwords' in browser_data:
                    for pd in browser_data['passwords']:
                        if (pd.get('url', '') == url and 
                            pd.get('username', '') == username):
                            pwd_data = pd
                            break
                if pwd_data:
                    break
        
        if not pwd_data:
            messagebox.showwarning("Not Found", "Password data not found.")
            return
        
        # Create details window
        details_window = tk.Toplevel(self.root)
        details_window.title("Password Encryption Details")
        details_window.geometry("600x400")
        details_window.resizable(True, True)
        
        # Create text widget with scrollbar
        text_frame = ttk.Frame(details_window)
        text_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        text_widget = tk.Text(text_frame, wrap="word", font=("Consolas", 10))
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        # Add details
        details = f"""Password Encryption Analysis
{'='*50}

Website: {pwd_data.get('url', 'N/A')}
Username: {pwd_data.get('username', 'N/A')}

Encryption Information:
- Type: {pwd_data.get('encryption_type', 'Unknown')}
- Decryption Status: {'Success' if pwd_data.get('password_decrypted') else 'Failed'}

Decrypted Password: {pwd_data.get('password_decrypted', 'Failed to decrypt')}

Raw Encrypted Data:
{pwd_data.get('raw_password_string', 'No raw data available')}

Timestamps:
- Created: {datetime_to_ddmmyy(pwd_data.get('date_created')) if pwd_data.get('date_created') else 'N/A'}

Technical Details:
- Encrypted data length: {len(pwd_data.get('password_encrypted', b'')) if pwd_data.get('password_encrypted') else 0} bytes
- Platform: {get_platform()}
- Available crypto libraries: {'pycryptodome' if CRYPTO_AVAILABLE else 'None'}, {'win32crypt' if WIN32CRYPT_AVAILABLE else 'No win32crypt'}
"""
        
        text_widget.insert("1.0", details)
        text_widget.config(state="disabled")
        
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Add close button
        ttk.Button(details_window, text="Close", command=details_window.destroy).pack(pady=5)
    
    # Copy password data - REMOVED
    
    # Export passwords dialog - REMOVED

    # ===== COMPREHENSIVE COPY SYSTEM =====
    
    def _show_context_menu(self, event, tree, menu):
        """Show context menu for any tree view"""
        try:
            # Check if there's an item at the click position
            item = tree.identify_row(event.y)
            if item:
                # Select the item if not already selected
                if item not in tree.selection():
                    tree.selection_set(item)
                menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
    
    def _copy_data(self, tree, column_index, data_type):
        """Copy specific column data from selected row to clipboard"""
        try:
            selection = tree.selection()
            if not selection:
                messagebox.showwarning("No Selection", f"Please select a row to copy {data_type}.")
                return
            
            item = selection[0]  # Get first selected item
            values = tree.item(item, "values")
            
            if column_index < len(values):
                data = str(values[column_index])
                self.root.clipboard_clear()
                self.root.clipboard_append(data)
                self.status_var.set(f"Copied {data_type} to clipboard")
                # Optional: Show brief confirmation
                # messagebox.showinfo("Copied", f"{data_type} copied to clipboard")
            else:
                messagebox.showwarning("Error", f"No {data_type} data available")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to copy {data_type}: {str(e)}")
    
    def _copy_password_special(self, tree):
        """Special copy function for passwords that gets the actual decrypted password"""
        try:
            selection = tree.selection()
            if not selection:
                messagebox.showwarning("No Selection", "Please select a password entry first.")
                return
            
            item = selection[0]
            values = tree.item(item, "values")
            
            if len(values) >= 2:
                url = values[0]
                username = values[1]
                # Get actual decrypted password
                actual_password = self._get_actual_password(url, username)
                
                if actual_password:
                    self.root.clipboard_clear()
                    self.root.clipboard_append(actual_password)
                    self.status_var.set("Password copied to clipboard")
                    # messagebox.showinfo("Copied", "Password copied to clipboard")
                else:
                    messagebox.showwarning("No Password", "No password found for this entry")
            else:
                messagebox.showwarning("Error", "Invalid password entry")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to copy password: {str(e)}")
    
    def _copy_selected_row(self, tree):
        """Copy entire selected row as tab-separated values"""
        try:
            selection = tree.selection()
            if not selection:
                messagebox.showwarning("No Selection", "Please select a row to copy.")
                return
            
            item = selection[0]
            values = tree.item(item, "values")
            
            # Join all values with tabs for easy pasting into spreadsheets
            row_data = "\t".join(str(v) for v in values)
            
            self.root.clipboard_clear()
            self.root.clipboard_append(row_data)
            self.status_var.set("Row copied to clipboard")
            # messagebox.showinfo("Copied", "Row copied to clipboard")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to copy row: {str(e)}")
    
    def _copy_all_selected_rows(self, tree):
        """Copy all selected rows as tab-separated values with headers"""
        try:
            selection = tree.selection()
            if not selection:
                messagebox.showwarning("No Selection", "Please select rows to copy.")
                return
            
            # Get column headers
            headers = []
            for col in tree["columns"]:
                headers.append(tree.heading(col, "text"))
            
            # Prepare data
            data_lines = ["\t".join(headers)]  # Header row
            
            for item in selection:
                values = tree.item(item, "values")
                data_lines.append("\t".join(str(v) for v in values))
            
            # Join all lines
            full_data = "\n".join(data_lines)
            
            self.root.clipboard_clear()
            self.root.clipboard_append(full_data)
            self.status_var.set(f"Copied {len(selection)} rows to clipboard")
            # messagebox.showinfo("Copied", f"Copied {len(selection)} rows with headers to clipboard")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to copy rows: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ForensicBrowserApp(root)
    root.mainloop()
