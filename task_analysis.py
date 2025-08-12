import pandas as pd
from tkinter import Tk, filedialog, messagebox

# Function to process the uploaded file
def analyze_file(file_path):
    try:
        # Read the CSV file
        data = pd.read_csv(file_path)

        # Process the data (example operations)
        data_filtered = data[['Feature_ID', 'Feature_Name', 'PBI_Id', 'PBI_Title',
                              'Last_Done_by_Resource_Name', 'Working_Hours']]
        data_filtered.rename(columns={
            'Feature_ID': 'FeatureID',
            'Feature_Name': 'FeatureName',
            'PBI_Id': 'TaskID',
            'PBI_Title': 'TaskTitle',
            'Last_Done_by_Resource_Name': 'Resource',
            'Working_Hours': 'Effort'
        }, inplace=True)

        data_filtered['Days'] = data_filtered['Effort'] / 8

        # Generate summaries
        feature_summary = data_filtered.groupby(['FeatureID', 'FeatureName']).agg(
            TotalHours=('Effort', 'sum'),
            TotalDays=('Days', 'sum')
        ).reset_index()

        task_summary = data_filtered.groupby(['FeatureID', 'FeatureName', 'TaskID', 'TaskTitle']).agg(
            TotalHours=('Effort', 'sum'),
            TotalDays=('Days', 'sum')
        ).reset_index()

        resource_summary = data_filtered.groupby(['FeatureID', 'FeatureName', 'TaskID', 'TaskTitle', 'Resource']).agg(
            TotalHours=('Effort', 'sum'),
            TotalDays=('Days', 'sum')
        ).reset_index()

        # Save results to an Excel file
        output_file = "Analysis_Result.xlsx"
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            feature_summary.to_excel(writer, sheet_name="Feature Summary", index=False)
            task_summary.to_excel(writer, sheet_name="Task Summary", index=False)
            resource_summary.to_excel(writer, sheet_name="Resource Summary", index=False)

        messagebox.showinfo("Success", f"Analysis completed! Results saved to {output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# GUI for file selection
def select_file():
    # Open a file dialog
    file_path = filedialog.askopenfilename(
        title="Select a CSV file",
        filetypes=[("CSV Files", "*.csv")]
    )
    if file_path:
        analyze_file(file_path)

# Main script
if __name__ == "__main__":
    # Create the main GUI window
    root = Tk()
    root.withdraw()  # Hide the root window
    messagebox.showinfo("Welcome", "Select a CSV file for analysis.")
    select_file()
