import os
import pandas as pd


def get_common_columns(dataframes):
    """Get columns that exist in all dataframes"""
    if not dataframes:
        return []
    common_cols = set(dataframes[0].columns)
    for df in dataframes[1:]:
        common_cols &= set(df.columns)
    return list(common_cols)


def translate_columns(df):
    """Translate Swedish column names to English"""
    translations = {
        "Valuta": "Currency",
        "Plats": "Location",
        "Kortinnehavare": "Cardholder",
        "Fakturabelopp": "Invoice Amount",
        "Detaljer": "Details",
        "Transaktionsbelopp": "Transaction Amount",
        "Source File": "Source File",
        "Datum": "Date",
        "Köpdatum": "Purchase Date",
        "Bokföringsdatum": "Posting Date",
    }
    df.columns = [translations.get(col, col) for col in df.columns]
    return df


def read_excel_file(file):
    """Read Excel file with appropriate engine"""
    try:
        # Read with second row as headers
        if file.endswith(".xls"):
            df = pd.read_excel(file, dtype=str, engine="xlrd", header=1)
        else:
            df = pd.read_excel(file, dtype=str, engine="openpyxl", header=1)

        # Clean column names first
        df.columns = df.columns.str.strip()

        # Translate columns
        df = translate_columns(df)

        # Convert date column after translation (now called "Date")
        if "Date" in df.columns:
            df["Date"] = pd.to_datetime(df["Date"]).dt.strftime("%Y/%m/%d")

        # Convert numeric columns
        if "Transaction Amount" in df.columns:
            df["Transaction Amount"] = pd.to_numeric(
                df["Transaction Amount"], errors="coerce"
            )
        if "Invoice Amount" in df.columns:
            df["Invoice Amount"] = pd.to_numeric(df["Invoice Amount"], errors="coerce")

        return df
    except Exception as e:
        raise Exception(f"Failed to read {file}: {str(e)}")


def merge_excel_files(output_file="Combined_Transactions.xlsx"):
    """Merge all Excel files (both .xls and .xlsx) into one"""
    # Skip temporary Excel files
    files = [
        f
        for f in os.listdir()
        if f.endswith((".xls", ".xlsx"))
        and not f.startswith("~$")
        and not f == output_file  # Skip the output file itself
    ]
    all_dataframes = []

    for file in files:
        try:
            df = read_excel_file(file)
            df["Source File"] = file
            all_dataframes.append(df)
        except Exception as e:
            print(f"Error reading {file}: {str(e)}")
            continue

    if all_dataframes:
        # Get common columns across all dataframes
        common_cols = get_common_columns(all_dataframes)

        # Remove unwanted columns
        cols_to_remove = ["Source File", "Date", "Purchase Date"]
        common_cols = [col for col in common_cols if col not in cols_to_remove]

        # Only keep common columns in each dataframe
        aligned_dfs = [df[common_cols] for df in all_dataframes]

        combined_df = pd.concat(aligned_dfs, ignore_index=True)
        combined_df.to_excel(output_file, index=False, engine="openpyxl", header=False)
        print(f"✅ Combined file saved as: {output_file}")
        print(f"Columns in combined file: {', '.join(combined_df.columns)}")
    else:
        print("❌ No Excel files found or no valid data to combine.")


if __name__ == "__main__":
    merge_excel_files()
