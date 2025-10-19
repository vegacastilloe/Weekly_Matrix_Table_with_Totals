import pandas as pd
from tabulate import tabulate
df_raw = pd.read_excel(url, header=1, nrows=4)
df_raw.columns = df_raw.columns.str.strip()
df = df_raw.iloc[:, 1:3]

# Crear columnas para 6 semanas
for week in range(1, 7):
    df[f'Week {week}'] = df['Prices']

# Calcular Total 1 (semanas 1 a 3) y Grand Total (semanas 1 a 6)
df['Total 1'] = df[[f'Week {i}' for i in range(1, 4)]].sum(axis=1)
df['Grand Total'] = df[[f'Week {i}' for i in range(1, 7)]].sum(axis=1)

# Reordenar columnas: insertar Total 1 después de Week 3
cols = ['Item'] + \
       [f'Week {i}' for i in range(1, 4)] + \
       ['Total 1'] + \
       [f'Week {i}' for i in range(4, 7)] + \
       ['Grand Total']

df_final = df[cols]
print(tabulate(df_final, headers='keys', tablefmt='fancy_grid'))
