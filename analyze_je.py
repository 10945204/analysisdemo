import pandas as pd
import matplotlib.pyplot as plt
import sys

def analyze_je():
    try:
        # Load the data
        print("Loading data...")
        df = pd.read_excel('je_samples.xlsx')

        # Clean data
        # Debit seems to be object type, convert to numeric, coerce errors to NaN
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce').fillna(0)

        # Ensure date columns are datetime
        df['EffectiveDate'] = pd.to_datetime(df['EffectiveDate'])
        df['EntryDate'] = pd.to_datetime(df['EntryDate'])

        # Prepare report content
        report_lines = []
        report_lines.append("Analysis Report for je_samples.xlsx")
        report_lines.append("===================================")

        # Basic Statistics
        report_lines.append(f"\nTotal Rows: {len(df)}")
        report_lines.append(f"Total Columns: {len(df.columns)}")
        report_lines.append("\nColumn Descriptions:")
        for col in df.columns:
            report_lines.append(f"  - {col}: {df[col].dtype}")

        # Date Ranges
        report_lines.append("\nDate Ranges:")
        if not df['EffectiveDate'].isnull().all():
            report_lines.append(f"  - EffectiveDate: {df['EffectiveDate'].min()} to {df['EffectiveDate'].max()}")
        else:
             report_lines.append("  - EffectiveDate: No valid dates found")

        if not df['EntryDate'].isnull().all():
            report_lines.append(f"  - EntryDate: {df['EntryDate'].min()} to {df['EntryDate'].max()}")
        else:
             report_lines.append("  - EntryDate: No valid dates found")

        # Financial Analysis
        report_lines.append("\nFinancial Analysis:")
        report_lines.append(f"  - Total Debit: {df['Debit'].sum():,.2f}")
        report_lines.append(f"  - Total Credit: {df['Credit'].sum():,.2f}")
        report_lines.append(f"  - Total Amount: {df['Amount'].sum():,.2f}")
        report_lines.append(f"  - Average Amount: {df['Amount'].mean():,.2f}")

        # Categorical Analysis
        report_lines.append("\nCategorical Analysis (Top 5):")

        categorical_cols = ['BusinessUnit', 'Source', 'PreparerID', 'AccountType']
        for col in categorical_cols:
            if col in df.columns:
                report_lines.append(f"\n  {col}:")
                top_counts = df[col].value_counts().head(5)
                for val, count in top_counts.items():
                    report_lines.append(f"    - {val}: {count}")

        # Write report to file
        with open('analysis_report.txt', 'w') as f:
            f.write('\n'.join(report_lines))
        print("Report generated: analysis_report.txt")

        # Generate Charts
        print("Generating charts...")

        # 1. Entries over time (EffectiveDate)
        plt.figure(figsize=(10, 6))
        df['EffectiveDate'].hist(bins=50)
        plt.title('Distribution of Entries by Effective Date')
        plt.xlabel('Date')
        plt.ylabel('Count')
        plt.savefig('entries_over_time.png')
        plt.close()

        # 2. Amount by Business Unit
        if 'BusinessUnit' in df.columns:
            plt.figure(figsize=(10, 6))
            # Use AbsoluteAmount for magnitude
            df.groupby('BusinessUnit')['AbsoluteAmount'].sum().sort_values(ascending=False).head(10).plot(kind='bar')
            plt.title('Total Absolute Amount by Business Unit (Top 10)')
            plt.ylabel('Total Absolute Amount')
            plt.tight_layout()
            plt.savefig('amount_by_business_unit.png')
            plt.close()

        # 3. Amount by Account Type
        if 'AccountType' in df.columns:
            plt.figure(figsize=(10, 6))
            df.groupby('AccountType')['AbsoluteAmount'].sum().sort_values(ascending=False).plot(kind='bar')
            plt.title('Total Absolute Amount by Account Type')
            plt.ylabel('Total Absolute Amount')
            plt.tight_layout()
            plt.savefig('amount_by_account_type.png')
            plt.close()

        # 4. Amount Distribution
        plt.figure(figsize=(10, 6))
        df['Amount'].hist(bins=50)
        plt.title('Distribution of Amount')
        plt.xlabel('Amount')
        plt.ylabel('Frequency')
        plt.savefig('amount_distribution.png')
        plt.close()

        print("Charts generated.")

    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    analyze_je()
