import pandas as pd
import matplotlib.pyplot as plt
import sys
import numpy as np

def get_leading_digit(n):
    """Extracts the first non-zero digit (1-9) from a number."""
    s = str(n)
    for char in s:
        if char in '123456789':
            return int(char)
    return None

def analyze_je():
    try:
        # Load the data
        print("Loading data...")
        df = pd.read_excel('je_samples.xlsx')

        # Clean data
        # Debit seems to be object type, convert to numeric, coerce errors to NaN
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce').fillna(0)

        # Ensure AbsoluteAmount is populated correctly
        df['AbsoluteAmount'] = df['Amount'].abs()

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

        # Benford's Law Analysis
        report_lines.append("\nBenford's Law Analysis:")
        df['LeadingDigit'] = df['AbsoluteAmount'].apply(get_leading_digit)
        benford_df = df[df['LeadingDigit'].notna()].copy()
        total_count = len(benford_df)

        if total_count > 0:
            actual_counts = benford_df['LeadingDigit'].value_counts().sort_index()
            expected_probs = {1: 0.301, 2: 0.176, 3: 0.125, 4: 0.097, 5: 0.079, 6: 0.067, 7: 0.058, 8: 0.051, 9: 0.046}

            report_lines.append(f"  Total analyzed entries: {total_count}")
            report_lines.append(f"  {'Digit':<5} {'Actual %':<10} {'Expected %':<10} {'Diff %':<10} {'Flag'}")

            digits = range(1, 10)
            actual_pcts = []
            expected_pcts = []

            for d in digits:
                count = actual_counts.get(d, 0)
                actual_pct = count / total_count
                expected_pct = expected_probs[d]
                diff = actual_pct - expected_pct

                actual_pcts.append(actual_pct)
                expected_pcts.append(expected_pct)

                flag = ""
                if abs(diff) > 0.05:
                    flag = "SIGNIFICANT DEVIATION"

                report_lines.append(f"  {d:<5} {actual_pct:.1%}     {expected_pct:.1%}     {diff:+.1%}     {flag}")
        else:
             report_lines.append("  No valid amounts for Benford's Law analysis.")

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

        # 5. Benford's Law Analysis
        if total_count > 0:
            plt.figure(figsize=(10, 6))
            x = np.arange(1, 10)
            width = 0.35

            plt.bar(x - width/2, [p * 100 for p in actual_pcts], width, label='Actual')
            plt.bar(x + width/2, [p * 100 for p in expected_pcts], width, label='Expected', alpha=0.7)

            plt.xlabel('Leading Digit')
            plt.ylabel('Frequency (%)')
            plt.title("Benford's Law Analysis: Actual vs Expected")
            plt.xticks(x)
            plt.legend()
            plt.tight_layout()
            plt.savefig('benford_analysis.png')
            plt.close()

        print("Charts generated.")

    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    analyze_je()
