#!/usr/bin/env python3
import pandas as pd

def main():
    path = 'cab_changes_5_weeks.xlsx'
    try:
        df = pd.read_excel(path, dtype=str)
    except Exception as e:
        print('Error reading Excel:', e)
        return
    print('Columns:', list(df.columns))
    for col in ['Code de fermeture', 'Etat', 'État']:
        if col in df.columns:
            vals = sorted(set((str(x) if x is not None else '').strip() for x in df[col].dropna()))
            print(f'Unique {col}:', vals)

if __name__ == '__main__':
    main()

