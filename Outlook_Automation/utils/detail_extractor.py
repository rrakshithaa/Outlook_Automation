import pandas as pd


class DetailExtractor:
    """Handles updating extracted details into Excel or CSV files."""

    @staticmethod
    def update_file(file_path, data):
        try:
            df = pd.read_excel(file_path) if file_path.endswith(".xlsx") else pd.read_csv(file_path)
        except FileNotFoundError:
            # Create a new DataFrame if the file does not exist
            df = pd.DataFrame(columns=["Candidate Name", "Experience", "Phone Number", "Email ID"])

        # Append the extracted details to the DataFrame
        df = df.append(data, ignore_index=True)

        # Save the DataFrame back to the file
        if file_path.endswith(".xlsx"):
            df.to_excel(file_path, index=False)
        else:
            df.to_csv(file_path, index=False)
