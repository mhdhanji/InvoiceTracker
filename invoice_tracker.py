import customtkinter as ctk
import pandas as pd
from tkinter import ttk, messagebox
import os
import imaplib
import email
from email.header import decode_header
from datetime import datetime, timedelta
import time
from calendar import monthrange
import csv
import email.utils

# After your imports
def fixed_map(style, option):
    return [elm for elm in style.map('Treeview', query_opt=option)
            if elm[:2] != ('!disabled', '!selected')]

# Email configuration
EMAIL_ADDRESS = "mapreportphones@gmail.com"
EMAIL_PASSWORD = "rxekijbdisdjksin"
CHECK_TIME_HOUR = 22
CHECK_TIME_MINUTE = 38
REFRESH_INTERVAL = 30000  # 30 seconds in milliseconds

class EmailProcessor:
    def __init__(self):
        self.email_address = EMAIL_ADDRESS
        self.password = EMAIL_PASSWORD
        self.imap_server = "imap.gmail.com"
        self.mail = None
        self.processed_files = set()
        self.processed_reports = set()  # Track processed report filenames
        self.current_month = datetime.now().month
        self.load_processed_files()

    def load_processed_files(self):
        """Load list of previously processed files"""
        try:
            if os.path.exists('processed_files.txt'):
                with open('processed_files.txt', 'r') as f:
                    self.processed_files = set(f.read().splitlines())
            if os.path.exists('processed_reports.txt'):
                with open('processed_reports.txt', 'r') as f:
                    self.processed_reports = set(f.read().splitlines())
        except Exception as e:
            print(f"Error loading processed files: {e}")
            self.processed_files = set()
            self.processed_reports = set()

    def save_processed_files(self):
        """Save list of processed files"""
        try:
            with open('processed_files.txt', 'w') as f:
                f.write('\n'.join(self.processed_files))
            with open('processed_reports.txt', 'w') as f:
                f.write('\n'.join(self.processed_reports))
        except Exception as e:
            print(f"Error saving processed files: {e}")

    def connect(self):
        try:
            if self.mail is not None:
                try:
                    self.mail.close()
                    self.mail.logout()
                except:
                    pass
            
            self.mail = imaplib.IMAP4_SSL(self.imap_server)
            self.mail.login(self.email_address, self.password)
            return True
        except Exception as e:
            print(f"Failed to connect to email: {str(e)}")
            self.mail = None
            return False

    def get_report_date(self, filename):
        """Extract date from DOPER report filename"""
        try:
            # Extract date from format: 504DOPER_YYYYMMDD
            date_str = filename.split('_')[1][:8]
            return datetime.strptime(date_str, '%Y%m%d')
        except Exception as e:
            print(f"Error extracting date from filename: {e}")
            return None

    def get_latest_report(self):
        if self.mail is None:
            if not self.connect():
                return None
            
        try:
            # Select inbox
            self.mail.select("INBOX")

            # Search for emails from the last few days
            date_str = (datetime.now() - timedelta(days=7)).strftime("%d-%b-%Y")
            search_criteria = f'(SINCE "{date_str}")'
        
            _, message_numbers = self.mail.search(None, search_criteria)

            if not message_numbers[0]:
                print("No messages found")
                return None

            # Check all messages from newest to oldest
            for num in message_numbers[0].split()[::-1]:
                _, msg_data = self.mail.fetch(num, "(RFC822)")
                email_body = msg_data[0][1]
                msg = email.message_from_bytes(email_body)

                # Get message ID or create unique identifier
                message_id = msg.get('Message-ID', '')
                email_date_str = msg.get('Date', '')
                unique_id = f"{message_id}_{email_date_str}"

                # Check attachments
                for part in msg.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue

                    filename = part.get_filename()
                    if filename and "504DOPER" in filename:
                        # Check if we've already processed this report
                        if filename in self.processed_reports:
                            print(f"Report already processed: {filename}")
                            return None
                            
                        # Check report date
                        report_date = self.get_report_date(filename)
                        if report_date:
                            current_month = datetime.now().month
                            if report_date.month < current_month:
                                print(f"Skipping old report from month {report_date.month}")
                                continue
                            
                        # Save the attachment
                        with open(filename, 'wb') as f:
                            f.write(part.get_payload(decode=True))
                        
                        # Mark as processed
                        self.processed_files.add(unique_id)
                        self.processed_reports.add(filename)
                        self.save_processed_files()
                        return filename

            print("No new matching attachments found")
            return None

        except Exception as e:
            print(f"Failed to get report: {str(e)}")
            return None
        finally:
            try:
                self.mail.close()
                self.mail.logout()
            except:
                pass

    def check_for_month_change(self):
        """Check if we've entered a new month"""
        current_month = datetime.now().month
        if current_month != self.current_month:
            print(f"New month detected: {self.current_month} -> {current_month}")
            self.current_month = current_month
            self.processed_files.clear()
            self.processed_reports.clear()
            self.save_processed_files()
            return True
        return False

    def reset_processed_files(self):
        """Reset processed files tracking for new month"""
        self.processed_files.clear()
        self.processed_reports.clear()
        self.save_processed_files()
        self.current_month = datetime.now().month
class SalesTracker:
    def __init__(self):
        self.current_month = datetime.now().month
        self.data = pd.DataFrame(columns=['Name', 'Number'])
        self.load_or_reset_data()

    def is_new_month(self):
        """Check if we've entered a new month"""
        return datetime.now().month != self.current_month

    def load_or_reset_data(self):
        try:
            # Check if we're in a new month
            if self.is_new_month():
                # Archive the previous month's data
                self.archive_monthly_data()
                
                # Reset data for new month
                self.data = pd.DataFrame(columns=['Name', 'Number'])
                self.current_month = datetime.now().month
                self.data.to_csv('sales_data.csv', index=False)
                print("New month started - data reset")
                return

            # Normal operation - load existing data
            if os.path.exists('sales_data.csv'):
                self.data = pd.read_csv('sales_data.csv')
                self.data['Number'] = pd.to_numeric(self.data['Number'], errors='coerce').fillna(0)
        except Exception as e:
            print(f"Error loading data: {e}")
            self.data = pd.DataFrame(columns=['Name', 'Number'])

    def archive_monthly_data(self):
        """Archive the previous month's data"""
        try:
            # Create archives directory if it doesn't exist
            archive_dir = "archives"
            if not os.path.exists(archive_dir):
                os.makedirs(archive_dir)
            
            # Get previous month and year
            today = datetime.now()
            first_day = today.replace(day=1)
            last_month = first_day - timedelta(days=1)
            month_year = last_month.strftime("%B_%Y")
            
            # Save current data to archive
            if not self.data.empty:
                archive_file = f"{archive_dir}/sales_data_{month_year}.csv"
                self.data.to_csv(archive_file, index=False)
                print(f"Data archived to {archive_file}")
                
                # Create monthly summary
                summary_file = f"{archive_dir}/summary_{month_year}.txt"
                with open(summary_file, 'w') as f:
                    f.write(f"Monthly Summary for {month_year}\n")
                    f.write("=" * 50 + "\n\n")
                    f.write(f"Total Invoices: {int(self.data['Number'].sum()):,}\n")
                    f.write(f"Total Staff: {len(self.data)}\n")
                    if not self.data.empty:
                        f.write(f"Top Performer: {self.data.loc[self.data['Number'].idxmax(), 'Name']}\n")
                    
                return True
        except Exception as e:
            print(f"Error archiving data: {e}")
            return False

    def process_daily_report(self, filename):
        try:
            # Extract date from filename
            try:
                date_str = filename.split('_')[1][:8]  # Get YYYYMMDD part
                report_date = datetime.strptime(date_str, '%Y%m%d')
                report_month = report_date.month
            
                # If report is from a previous month, ignore it
                if report_month < self.current_month:
                    print(f"Ignoring report from previous month: {report_month}")
                    return False
            
                # If report is from a new month, reset data first
                if report_month != self.current_month:
                    print(f"New month detected in report: {report_month}")
                    self.archive_monthly_data()
                    self.data = pd.DataFrame(columns=['Name', 'Number'])
                    self.current_month = report_month
            except Exception as e:
                print(f"Error processing report date: {e}")
                return False

            # Read the daily report
            daily_data = pd.read_csv(filename)
        
            # Process only Name and Number columns
            daily_data = daily_data[['Name', 'Number']]
        
            # Convert Number to numeric
            daily_data['Number'] = pd.to_numeric(daily_data['Number'], errors='coerce').fillna(0)
        
            # Remove rows with empty names
            daily_data = daily_data[daily_data['Name'].notna() & (daily_data['Name'] != "")]
        
            # Update data based on the month
            if report_month == self.current_month:
            # For the first report of the month, just set the data
                if self.data.empty:
                    self.data = daily_data
                else:
                    # For subsequent reports, add to existing entries
                    for _, row in daily_data.iterrows():
                        if row['Name'] in self.data['Name'].values:
                            # Add to existing entry
                            self.data.loc[self.data['Name'] == row['Name'], 'Number'] += row['Number']
                        else:
                            # Add new entry
                            self.data = pd.concat([self.data, pd.DataFrame([row])], ignore_index=True)
        
            # Sort by number of invoices
            self.data = self.data.sort_values('Number', ascending=False)
        
            # Save updated data
            self.data.to_csv('sales_data.csv', index=False)
        
            # Clean up the report file
            try:
                os.remove(filename)
            except:
                pass
        
            return True
        except Exception as e:
            print(f"Failed to process report: {str(e)}")
            return False
class InvoiceTrackerApp:
    def __init__(self):
        # Initialize sales tracker
        self.sales_tracker = SalesTracker()
        
        # Initialize email processor
        self.email_processor = EmailProcessor()

        # Initialize auto-update job tracker
        self.auto_update_job = None

        # Set up the window
        self.window = ctk.CTk()
        self.window.title("Sales Dashboard")
        self.window.attributes('-fullscreen', True)
        
        # Set light theme
        self.is_dark_theme = False
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        # Create main container
        self.main_container = ctk.CTkFrame(self.window)
        self.main_container.pack(fill="both", expand=True, padx=20, pady=20)

        # Create header
        self.create_header()
        
        # Create main table
        self.create_main_table()
        
        # Bind keyboard shortcuts
        self.window.bind('<Escape>', lambda e: self.window.quit())
        self.window.bind('<Control-t>', lambda e: self.toggle_theme())
        self.window.bind('<Configure>', self.on_resize)
        
        # Initialize loading state
        self.is_loading = False
        
        # Load initial data
        if hasattr(self, 'sales_tracker'):
            self.update_table(self.sales_tracker.data)
        
        # Start auto-update
        self.auto_update()

    def create_header(self):
        self.header_frame = ctk.CTkFrame(self.main_container)
        self.header_frame.pack(fill="x", padx=10, pady=(0, 20))

        # Left side - Title with buttons underneath
        self.left_section = ctk.CTkFrame(self.header_frame)
        self.left_section.pack(side="left", padx=20, pady=20)
        
        # Title
        self.title_label = ctk.CTkLabel(
            self.left_section,
            text="SALES LEADERBOARD",
            font=("Arial Black", 48, "bold")
        )
        self.title_label.pack()

        # Buttons frame under title
        self.buttons_frame = ctk.CTkFrame(self.left_section)
        self.buttons_frame.pack(fill="x", pady=(10, 0))

        # Reset Button
        self.reset_button = ctk.CTkButton(
            self.buttons_frame,
            text="🗑️ Reset Table",
            command=self.reset_table,
            width=120,
            height=40,
            font=("Arial", 16),
            fg_color="red",
            hover_color="#AA0000"
        )
        self.reset_button.pack(side="left", padx=5)

        # Manual Update Button
        self.update_button = ctk.CTkButton(
            self.buttons_frame,
            text="🔄 Update Now",
            command=self.manual_update,
            width=120,
            height=40,
            font=("Arial", 16)
        )
        self.update_button.pack(side="left", padx=5)

        # Theme toggle button
        self.theme_button = ctk.CTkButton(
            self.buttons_frame,
            text="🌙 Dark Mode",
            command=self.toggle_theme,
            width=120,
            height=40,
            font=("Arial", 16)
        )
        self.theme_button.pack(side="left", padx=5)

        # Compare Months button
        self.compare_button = ctk.CTkButton(
            self.buttons_frame,
            text="📊 Compare Months",
            command=self.create_comparison_window,
            width=120,
            height=40,
            font=("Arial", 16)
        )
        self.compare_button.pack(side="left", padx=5)

        # Middle section - Check DOPER button and status
        self.middle_frame = ctk.CTkFrame(self.header_frame)
        self.middle_frame.pack(side="left", expand=True, padx=20, pady=20)

        # 504DOPER Check Button
        self.doper_button = ctk.CTkButton(
            self.middle_frame,
            text="📋 Check 504DOPER",
            command=self.check_504_report,
            width=200,
            height=40,
            font=("Arial", 16)
        )
        self.doper_button.pack(pady=5)

        # Status label under the button
        self.status_label = ctk.CTkLabel(
            self.middle_frame,
            text="Ready",
            font=("Arial", 12)
        )
        self.status_label.pack(pady=5)

        # Last Updated Time
        self.update_label = ctk.CTkLabel(
            self.middle_frame,
            text="Last Updated: Never",
            font=("Arial", 14)
        )
        self.update_label.pack(pady=5)

        # Exit button (right side)
        self.exit_button = ctk.CTkButton(
            self.header_frame,
            text="×",
            command=self.window.quit,
            width=40,
            height=40,
            font=("Arial", 20)
        )
        self.exit_button.pack(side="right", padx=20, pady=20)

    def check_504_report(self):
        """Check for 504DOPER report"""
        self.status_label.configure(text="Checking for new 504DOPER reports...")
        self.window.update()
        
        try:
            if self.email_processor.connect():
                filename = self.email_processor.get_latest_report()
                if filename and "504DOPER" in filename:
                    try:
                        # Extract date from filename (504DOPER_YYYYMMDD)
                        date_str = filename.split('_')[1][:8]  # Get YYYYMMDD part
                        report_date = datetime.strptime(date_str, '%Y%m%d')
                        report_month = report_date.month
                        current_month = datetime.now().month
                        
                        # If report is from a previous month, ignore it
                        if report_month < current_month:
                            self.status_label.configure(text="Ignoring DOPER report from previous month")
                            try:
                                os.remove(filename)  # Clean up the downloaded file
                            except:
                                pass
                            return
                        
                        # Process the report
                        if self.sales_tracker.process_daily_report(filename):
                            self.update_table(self.sales_tracker.data)
                            self.status_label.configure(text=f"Processed DOPER report for {report_date.strftime('%B %Y')}")
                        else:
                            self.status_label.configure(text="Failed to process DOPER report")
                            
                    except Exception as e:
                        print(f"Error processing report date: {e}")
                        self.status_label.configure(text="Error processing report date")
                else:
                    self.status_label.configure(text="No new 504DOPER report found")
            else:
                self.status_label.configure(text="Failed to connect to email")
        except Exception as e:
            self.status_label.configure(text=f"Error: {str(e)}")

    def create_main_table(self):
        # Create frame for the table
        self.table_frame = ctk.CTkFrame(self.main_container)
        self.table_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Create style for the treeview
        self.style = ttk.Style()
        self.style.theme_use('default')

        # Configure Treeview style
        self.style.configure("Treeview",
            background="#2b2b2b" if self.is_dark_theme else "white",
            foreground="white" if self.is_dark_theme else "black",
            fieldbackground="#2b2b2b" if self.is_dark_theme else "white",
            rowheight=50,
            font=("Arial", 24)
        )
    
        self.style.configure("Treeview.Heading",
            background="#1f538d" if self.is_dark_theme else "#0078D7",
            foreground="white",
            relief="flat",
            font=("Arial", 26, "bold")
        )

        # Create scrollbar
        self.tree_scroll = ttk.Scrollbar(self.table_frame)
        self.tree_scroll.pack(side="right", fill="y")

        # Create Treeview with proper column proportions
        self.tree = ttk.Treeview(
            self.table_frame,
            columns=("Name", "Invoices"),
            show='headings',
            style="Treeview",
            selectmode="none",
            yscrollcommand=self.tree_scroll.set
        )

        self.tree_scroll.config(command=self.tree.yview)

        # Configure columns with proportional widths
        self.tree.heading("Name", text="NAME", anchor="w")
        self.tree.heading("Invoices", text="INVOICES", anchor="e")

        # Calculate initial widths
        total_width = self.window.winfo_screenwidth() - 40  # Account for padding
        name_width = int(total_width * 0.8)  # 80% for name
        invoices_width = int(total_width * 0.2)  # 20% for invoices

        self.tree.column("Name", anchor="w", width=name_width, stretch=True)
        self.tree.column("Invoices", anchor="e", width=invoices_width, stretch=True)

        # Configure row colors
        self.tree.tag_configure('oddrow', 
            background='#333333' if self.is_dark_theme else '#F0F0F0',
            foreground='white' if self.is_dark_theme else 'black'
        )
        self.tree.tag_configure('evenrow', 
            background='#2b2b2b' if self.is_dark_theme else 'white',
            foreground='white' if self.is_dark_theme else 'black'
        )
        self.tree.tag_configure('toprow', 
            background='#1f538d' if self.is_dark_theme else '#0078D7',
            foreground='white'
        )

        self.tree.pack(fill="both", expand=True, padx=(0, 10))  # Add padding only on right for scrollbar

    def update_table(self, df):
        """Update the display table with new data"""
        self.set_loading(True)
        self.tree.delete(*self.tree.get_children())

        try:
            df_sorted = df.sort_values('Number', ascending=False)
            for index, row in df_sorted.iterrows():
                tags = ('toprow',) if index == 0 else ('evenrow',) if index % 2 == 0 else ('oddrow',)
                self.tree.insert("", "end", values=(
                    row['Name'],
                    f"{int(row['Number']):,}"
                ), tags=tags)

            current_time = datetime.now().strftime('%d %B %Y, %H:%M')
            self.update_label.configure(text=f"Last Updated: {current_time}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update table: {str(e)}")
            
        finally:
            self.set_loading(False)

    def set_loading(self, is_loading):
        """Update loading state"""
        self.is_loading = is_loading
        if is_loading:
            self.status_label.configure(text="Loading...")
        else:
            self.status_label.configure(text="Ready")

    def manual_update(self):
        """Manually trigger an update"""
        self.check_504_report()

    def auto_update(self):
        """Automatic update function"""
        try:
            current_time = datetime.now()
            
            # Regular update check
            if (current_time.hour == CHECK_TIME_HOUR and 
                current_time.minute >= CHECK_TIME_MINUTE and 
                current_time.minute < CHECK_TIME_MINUTE + 1):
                
                self.check_504_report()

        except Exception as e:
            self.status_label.configure(text=f"Update error: {str(e)}")
        
        # Schedule next update
        self.auto_update_job = self.window.after(REFRESH_INTERVAL, self.auto_update)

    def reset_table(self):
        """Reset the table to zero"""
        result = messagebox.askyesno("Confirm Reset", 
            "Are you sure you want to reset all data to zero?\nThis action cannot be undone!")
        
        if result:
            try:
                # Archive current data before reset
                if self.sales_tracker.archive_monthly_data():
                    # Reset the data
                    self.sales_tracker.data = pd.DataFrame(columns=['Name', 'Number'])
                    # Save empty data
                    self.sales_tracker.data.to_csv('sales_data.csv', index=False)
                    # Update the display
                    self.update_table(self.sales_tracker.data)
                    
                    # Reset email processor's processed files tracking
                    self.email_processor.reset_processed_files()
                    
                    self.status_label.configure(text="Table reset successfully")
                else:
                    messagebox.showwarning("Warning", "Failed to archive data")
                    
            except Exception as e:
                self.status_label.configure(text="Failed to reset table")
                messagebox.showerror("Error", f"Failed to reset table: {str(e)}")

    def create_comparison_window(self):
        """Create window to compare monthly data"""
        comp_window = ctk.CTkToplevel()
        comp_window.title("Monthly Comparison")
        comp_window.geometry("1200x800")

        # Create control frame
        control_frame = ctk.CTkFrame(comp_window)
        control_frame.pack(fill="x", padx=10, pady=5)

        # Get available months from archives
        archive_dir = "archives"
        if not os.path.exists(archive_dir):
            os.makedirs(archive_dir)
        
        archives = [f.replace("sales_data_", "").replace(".csv", "") 
                   for f in os.listdir(archive_dir) 
                   if f.startswith("sales_data_")]
        archives.sort(reverse=True)
        archives.insert(0, "Current Month")  # Add current month as an option

        # Create month selection dropdowns
        month1_var = ctk.StringVar(value=archives[0] if archives else "")
        month2_var = ctk.StringVar(value=archives[1] if len(archives) > 1 else "")

        ctk.CTkLabel(control_frame, text="Compare:").pack(side="left", padx=5)
        month1_dropdown = ctk.CTkOptionMenu(
            control_frame, 
            variable=month1_var,
            values=archives,
            command=lambda x: update_comparison()
        )
        month1_dropdown.pack(side="left", padx=5)
    
        ctk.CTkLabel(control_frame, text="with:").pack(side="left", padx=5)
        month2_dropdown = ctk.CTkOptionMenu(
            control_frame,
            variable=month2_var,
            values=archives,
            command=lambda x: update_comparison()
        )
        month2_dropdown.pack(side="left", padx=5)

        # Create comparison table
        table_frame = ctk.CTkFrame(comp_window)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("Name", "Month1", "Month2", "Difference", "Change%")
        comp_tree = ttk.Treeview(table_frame, columns=columns, show='headings')

        # Configure columns
        for col in columns:
            comp_tree.heading(col, text=col.upper())
            comp_tree.column(col, width=200, anchor="center")

        comp_tree.pack(fill="both", expand=True)

        def update_comparison():
            comp_tree.delete(*comp_tree.get_children())
            
            # Load data for both months
            data1 = None
            data2 = None
            
            if month1_var.get() == "Current Month":
                data1 = self.sales_tracker.data
            else:
                file1 = f"{archive_dir}/sales_data_{month1_var.get()}.csv"
                if os.path.exists(file1):
                    data1 = pd.read_csv(file1)

            if month2_var.get() == "Current Month":
                data2 = self.sales_tracker.data
            else:
                file2 = f"{archive_dir}/sales_data_{month2_var.get()}.csv"
                if os.path.exists(file2):
                    data2 = pd.read_csv(file2)

            if data1 is not None and data2 is not None:
                # Merge the datasets
                merged = pd.merge(data1, data2, on='Name', how='outer', suffixes=('_1', '_2'))
                merged = merged.fillna(0)

                # Calculate differences and percentages
                for _, row in merged.iterrows():
                    val1 = row['Number_1']
                    val2 = row['Number_2']
                    diff = val2 - val1
                    pct = (diff / val1 * 100) if val1 != 0 else 0

                    # Format values
                    val1_fmt = f"{int(val1):,}"
                    val2_fmt = f"{int(val2):,}"
                    diff_fmt = f"{int(diff):,}"
                    pct_fmt = f"{pct:.1f}%"

                    # Add row with color coding
                    tags = ('positive',) if diff > 0 else ('negative',) if diff < 0 else ('neutral',)
                    comp_tree.insert("", "end", values=(
                        row['Name'],
                        val1_fmt,
                        val2_fmt,
                        diff_fmt,
                        pct_fmt
                    ), tags=tags)

                # Configure colors for changes
                comp_tree.tag_configure('positive', foreground='green')
                comp_tree.tag_configure('negative', foreground='red')
                comp_tree.tag_configure('neutral', foreground='black')

        # Add export button
        def export_comparison():
            try:
                export_file = f"comparison_{month1_var.get()}_vs_{month2_var.get()}.csv"
                with open(export_file, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow([col.upper() for col in columns])
                    for item in comp_tree.get_children():
                        writer.writerow(comp_tree.item(item)['values'])
                messagebox.showinfo("Success", f"Comparison exported to {export_file}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export: {str(e)}")

        export_btn = ctk.CTkButton(
            control_frame,
            text="Export Comparison",
            command=export_comparison
        )
        export_btn.pack(side="right", padx=5)

        # Initial comparison update
        update_comparison()

    def on_resize(self, event):
        """Handle window resize"""
        if hasattr(self, 'tree'):
            # Calculate new widths based on window size
            total_width = event.width - 40  # Account for padding
            name_width = int(total_width * 0.8)
            invoices_width = int(total_width * 0.2)
        
            # Update column widths
            self.tree.column("Name", width=name_width)
            self.tree.column("Invoices", width=invoices_width)

    def toggle_theme(self):
        """Toggle between light and dark theme"""
        self.is_dark_theme = not self.is_dark_theme
        if self.is_dark_theme:
            ctk.set_appearance_mode("dark")
            self.style.configure("Treeview",
                background="#2b2b2b",
                foreground="white",
                fieldbackground="#2b2b2b"
            )
            self.style.configure("Treeview.Heading",
                background="#1f538d",
                foreground="white"
            )
            self.tree.tag_configure('oddrow', background='#333333', foreground='white')
            self.tree.tag_configure('evenrow', background='#2b2b2b', foreground='white')
            self.tree.tag_configure('toprow', background='#1f538d', foreground='white')
        else:
            ctk.set_appearance_mode("light")
            self.style.configure("Treeview",
                background="white",
                foreground="black",
                fieldbackground="white"
            )
            self.style.configure("Treeview.Heading",
                background="#0078D7",
                foreground="white"
            )
            self.tree.tag_configure('oddrow', background='#F0F0F0', foreground='black')
            self.tree.tag_configure('evenrow', background='white', foreground='black')
            self.tree.tag_configure('toprow', background='#0078D7', foreground='white')

        if hasattr(self, 'sales_tracker'):
            self.update_table(self.sales_tracker.data)

    def run(self):
        """Start the application"""
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        self.window.geometry(f"{screen_width}x{screen_height}+0+0")
        self.window.mainloop()

if __name__ == "__main__":
    app = InvoiceTrackerApp()
    app.run()