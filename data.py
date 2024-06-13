import pandas as pd

# Define file path and sheet names
file_path = '/Users/sumukha.jayanth1/Downloads/Analysis Results __ Qualitative Eval __ 30 Prompt __ Spectre_v2, HF, vLLM, E+R_v1 (1).xlsx'  # Replace with your file path
sheet_languages = {
    'English': 'English',
    'Hindi': 'Hindi',
    'Kannada': 'Kannada',
    'Tamil': 'Tamil',
    'Telugu': 'Telugu',
    'Hinglish': 'Hinglish',
    'Malayalam': 'Malayalam',
    'Gujarati': 'Gujarati',
    'Marathi': 'Marathi',
    'Bengali': 'Bengali'
}

# Dictionary to map models to their corresponding remarks columns
models_remarks = {
    'Spectre_v2, HF': ['Remark 1 - Response Spectre_v2, HF', 'Remark 2 - Response Spectre_v2, HF', 'Remark 3 - Response Spectre_v2, HF'],
    'Spectre_v2, vLLM': ['Remark 1 - Response Spectre_v2, vLLM', 'Remark 2 - Response Spectre_v2, vLLM', 'Remark 3 - Response Spectre_v2, vLLM'],
    'E+R_v1': ['Remark 1 - Response E+R_v1', 'Remark 2 - Response E+R_v1', 'Remark 3 - Response E+R_v1']
}

# Create a new Excel writer object
with pd.ExcelWriter('Statistics.xlsx', engine='openpyxl') as writer:
    for language, sheet_name in sheet_languages.items():
        # Initialize a dictionary to store DataFrames for each model
        model_dfs = {}

        # Load data for the current language
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # Iterate over models and their corresponding remarks columns
        for model, remarks_columns in models_remarks.items():
            # Initialize a dictionary to store remarks, frequencies, and percentages
            model_data = {
                'Remark 1': [], 'Frequency 1': [], 'Percentage 1': [],
                'Remark 2': [], 'Frequency 2': [], 'Percentage 2': [],
                'Remark 3': [], 'Frequency 3': [], 'Percentage 3': []
            }

            for idx, remark_col in enumerate(remarks_columns):
                # Extract remarks from the current remark column
                remarks = df[remark_col].dropna()

                # Compute frequency and percentage of remarks
                freq = remarks.value_counts()
                percentage = (freq / freq.sum()) * 100

                # Sort values by frequency descending
                freq_sorted = freq.sort_values(ascending=False)
                
                # Fill model_data dictionary
                model_data[f'Remark {idx + 1}'] = freq_sorted.index.tolist()
                model_data[f'Frequency {idx + 1}'] = freq_sorted.values.tolist()
                model_data[f'Percentage {idx + 1}'] = percentage.loc[freq_sorted.index].values.tolist()

            # Determine the maximum length of the lists
            max_length = max(
                len(model_data['Remark 1']), len(model_data['Remark 2']), len(model_data['Remark 3'])
            )

            # Ensure all lists have the same length by padding with NaNs
            for key in model_data:
                while len(model_data[key]) < max_length:
                    model_data[key].append(float('nan'))

            # Create a DataFrame for the current model
            model_df = pd.DataFrame(model_data)
            model_df.columns = [
                f'{model} - Remark 1', f'{model} - Frequency 1', f'{model} - Percentage 1',
                f'{model} - Remark 2', f'{model} - Frequency 2', f'{model} - Percentage 2',
                f'{model} - Remark 3', f'{model} - Frequency 3', f'{model} - Percentage 3'
            ]

            # Store the model_df in model_dfs
            model_dfs[model] = model_df

        # Concatenate model DataFrames into a single DataFrame
        merged_df = pd.concat(model_dfs.values(), axis=1)

        # Check if DataFrame is empty
        if merged_df.empty:
            raise ValueError(f"DataFrame for language '{language}' is empty, cannot write to Excel.")

        # Write merged_df to a new sheet in the Excel file
        merged_df.to_excel(writer, sheet_name=language, index=False)
