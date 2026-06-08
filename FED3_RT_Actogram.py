import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import os
import re
import argparse
import urllib.request
import sys
import subprocess
from datetime import datetime

# Run this script
# From a terminal, run this command: 
# python FED3_RT_Actogram.py "https://docs.google.com/spreadsheets/d/1p5B2r8i9jiza19YGHup5oXVZIpM5BFaedN9zGxVCzOQ/edit?gid=103919384#gid=103919384"

def download_google_sheet_all_tabs(sheet_url):
    """
    Downloads the entire Google Sheet as an .xlsx file to capture all tabs,
    and saves it in the exact same folder as this Python script.
    """
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', sheet_url)
    
    if not match:
        print("❌ Error: Could not extract the Spreadsheet ID from the URL.")
        return None
        
    sheet_id = match.group(1)
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    script_directory = os.path.dirname(os.path.abspath(__file__))
    xlsx_path = os.path.join(script_directory, f"fed3_multifeeder_{sheet_id[:8]}.xlsx")
    
    print(f"Downloading multi-tab workbook to: '{script_directory}'...")
    try:
        urllib.request.urlretrieve(export_url, xlsx_path)
        print(f"✅ Workbook saved successfully to: '{xlsx_path}'\n")
        return xlsx_path
    except Exception as e:
        print(f"❌ Failed to download sheet: {e}")
        print("Note: Ensure your Google Sheet sharing setting is set to 'Anyone with the link'.")
        return None

def open_pdf_automatically(filepath):
    """Detects the operating system and opens the PDF in the default viewer."""
    print("Opening PDF viewer...")
    try:
        if sys.platform == "win32": # Windows
            os.startfile(filepath)
        elif sys.platform == "darwin": # macOS
            subprocess.call(['open', filepath])
        else: # Linux
            subprocess.call(['xdg-open', filepath])
    except Exception as e:
        print(f"⚠️ Could not automatically open the PDF. You can open it manually. Error: {e}")

def generate_actograms_pdf(xlsx_file, bin_minutes=10, light_start=11, dark_start=23):
    """Reads every tab and complies the generated actograms into a single PDF."""
    print("Reading workbook tabs...")
    try:
        all_tabs = pd.read_excel(xlsx_file, sheet_name=None)
    except Exception as e:
        print(f"❌ Error reading the Excel file: {e}")
        return

    time_col = 'MM/DD/YYYY hh:mm:ss.SSS'
    script_directory = os.path.dirname(os.path.abspath(__file__))
    
    # Create the PDF filename with a timestamp based on the downloaded Excel file
    current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    base_name = os.path.splitext(os.path.basename(xlsx_file))[0]
    pdf_path = os.path.join(script_directory, f"{base_name}_Actograms_{current_time}.pdf")

    print(f"Creating multi-page PDF binder...")
    
    # Initialize the PDF file
    with PdfPages(pdf_path) as pdf:
        # Loop through every tab found in the Google Sheet
        for tab_name, df in all_tabs.items():
            print(f"--- Processing Tab: '{tab_name}' ---")
            
            if time_col not in df.columns:
                print(f"⚠️ Skipping '{tab_name}': Doesn't look like FED3 data.\n")
                continue

            df[time_col] = pd.to_datetime(df[time_col], errors='coerce')
            df_pellets = df[df['Event'].astype(str).str.contains('PelletInWell', case=False, na=False)].copy()
            
            if df_pellets.empty:
                print(f"⚠️ Skipping '{tab_name}': No 'PelletInWell' events found.\n")
                continue

            df_pellets['Event_Count'] = 1
            df_pellets.set_index(time_col, inplace=True)
            binned_data = df_pellets['Event_Count'].resample(f'{bin_minutes}min').sum().reset_index()
            
            binned_data['Date'] = binned_data[time_col].dt.date
            binned_data['Time_of_Day'] = binned_data[time_col].dt.hour + (binned_data[time_col].dt.minute / 60.0)
            
            unique_dates = binned_data['Date'].unique()
            num_days = len(unique_dates)
            
            # Find the absolute maximum value across ALL days to fix the Y-axis scale
            max_events = binned_data['Event_Count'].max()
            y_max = (max_events * 1.1) if pd.notna(max_events) and max_events > 0 else 5
            
            # Using sharex=True and sharey=True
            fig, axes = plt.subplots(nrows=num_days, ncols=1, figsize=(10, 1.2 * num_days), sharex=True, sharey=True)
            if num_days == 1:
                axes = [axes]
                
            fig.suptitle(f'FED3 Actogram: {tab_name}', fontsize=14, fontweight='bold', y=1.05)
            bar_width_hours = bin_minutes / 60.0

            for i, (ax, date) in enumerate(zip(axes, unique_dates)):
                
                ax.set_zorder(num_days - i)
                ax.set_facecolor('none') 
                
                day_data = binned_data[binned_data['Date'] == date]
                ax.axvspan(0, light_start, color='lightgray', alpha=0.4, lw=0)
                ax.axvspan(light_start, dark_start, color='lemonchiffon', alpha=0.5, lw=0)
                ax.axvspan(dark_start, 24, color='lightgray', alpha=0.4, lw=0)
                
                ax.bar(day_data['Time_of_Day'], day_data['Event_Count'], 
                       width=bar_width_hours, color='black', align='edge', edgecolor='none')
                
                ax.set_xlim(0, 24)
                
                # Apply the fixed Y-limit
                ax.set_ylim(0, y_max) 
                
                # Limit to a few y-ticks so it isn't completely cluttered with numbers
                ax.locator_params(axis='y', nbins=3) 
                ax.tick_params(axis='y', labelsize=8)

                # Date label pad increased to 45 so it doesn't overlap the new Y-axis numbers
                ax.set_ylabel(date.strftime('%Y-%m-%d'), rotation=0, labelpad=45, ha='right', va='center', fontsize=9)
                
                # Keep top and right spines hidden, but keep the LEFT spine visible
                for spine in ['top', 'right']:
                    ax.spines[spine].set_visible(False)
                
                ax.spines['left'].set_visible(True)
                ax.spines['left'].set_color('black')
                ax.spines['left'].set_linewidth(1.0)
                
                ax.spines['bottom'].set_visible(True)
                ax.spines['bottom'].set_color('black')
                ax.spines['bottom'].set_linewidth(1.0) 
                
                # Turn off tick marks for all middle days to keep the lines crisp
                if ax != axes[-1]:
                    ax.tick_params(bottom=False)

            axes[-1].set_xlabel('Time of Day (h)', fontsize=11, labelpad=10)
            axes[-1].set_xticks(range(0, 25, 2))
            
            # Reverted to 0.0 for seamless stacked actogram styling
            plt.subplots_adjust(hspace=0.0)
            
            # Save the current figure as a new page in the PDF binder
            pdf.savefig(fig, bbox_inches='tight', facecolor='white')
            
            # Close the figure to free up memory before moving to the next tab
            plt.close(fig) 
            
            print(f"✅ Added '{tab_name}' to PDF.\n")

    print(f"🎉 Success! Multi-page PDF saved at: '{pdf_path}'")
    
    # Trigger the PDF to pop up on the screen
    open_pdf_automatically(pdf_path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Download multi-tab FED3 Google Sheet and Generate Actogram PDF")
    
    parser.add_argument('url', type=str, help='The full URL of the Google Sheet')
    parser.add_argument('--bins', type=int, default=10, help='Bin size in minutes')
    
    args = parser.parse_args()
    
    # 1. Download the sheet
    local_xlsx_path = download_google_sheet_all_tabs(args.url)
    
    # 2. Process all tabs and create the PDF
    if local_xlsx_path:
        generate_actograms_pdf(local_xlsx_path, bin_minutes=args.bins)