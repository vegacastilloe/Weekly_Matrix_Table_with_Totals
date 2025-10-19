# 📊 Weekly Matrix Table with Totals in pandas

![License: MIT](https://img.shields.io/badge/License-MIT-cyan.svg)
![Python](https://img.shields.io/badge/python-3.7%2B-blue)
![Last Updated](https://img.shields.io/github/last-commit/vegacastilloe/Weekly_Matrix_Table_with_Totals)
![Language](https://img.shields.io/badge/language-español-darkred)

#
---
- 🌟 Easy Sunday Excel Challenge 🌟
- 🌟 **Author**: Crispo Mwangi

    - ⭐Create a Matrix Table showing Total 1 and Grand Total.
    - ⭐Populate the prices across the 6 weeks.
    - ⭐Create Total 1 after 3 weeks and 6 weeks Grand Total.
    - ⭐Solution MUST be dynamic.

 🔰 Este script genera una tabla tipo matriz a partir de una lista de ítems y precios, replicando los valores a lo largo de 6 semanas y calculando totales intermedios y finales.

 🔗 Link to Excel file:
 👉 https://lnkd.in/dNzYRZaU

**My code in Python** 🐍 **for this challenge**

 🔗 https://github.com/vegacastilloe/Weekly_Matrix_Table_with_Totals/blob/main/weekly_matrix_table.py

---
---

## Weekly Matrix Table with Totals in pandas

Este script genera una tabla tipo matriz a partir de una lista de ítems y precios, replicando los valores a lo largo de 6 semanas y calculando totales intermedios y finales.



## 📦 Requisitos

- Python 3.9+
- Paquetes:
- pandas openpyxl (para leer .xlsx)
- tabulate (solo para impresión bonita)
- Archivo Excel con al menos:
    - La columna: `Item`.
    - La columna: `Prices`
    - En la columna 5 y siguientes: resultados esperados para comparación

---

## 🚀 Cómo funciona

1. Carga de archivo Excel con pandas. read_excel.
2. Limpieza de nombres de columnas (str.strip).
3. Replicar los precios en columnas `Week 1` a `Week 6`.
4. Calcular:
   - `Total 1` como suma de semanas 1 a 3
   - `Grand Total` como suma de semanas 1 a 6
5. Reordenar columnas para insertar `Total 1` después de `Week 3`.
6. Visualización con tabulate en formato fancy_grid.

---

## 📤 Salida

El script imprime un DataFrame con:

- `Item`, `Week 1`, `Week 2`, `Week 3`, `Total 1`, `Week 4`, `Week 5`, `Week 6`, `Grand Total`
- `Resultado esperado`

---

## 🧹 Output:


| Item   | Week 1 | Week 2 | Week 3 | Total 1 | Week 4 | Week 5 | Week 6 | Grand Total |
|--------|--------|--------|--------|---------|--------|--------|--------|--------------|
| Item 1 | 100    | 100    | 100    | 300     | 100    | 100    | 100    | 600          |
| Item 2 | 200    | 200    | 200    | 600     | 200    | 200    | 200    | 1200         |
| Item 3 | 300    | 300    | 300    | 900     | 300    | 300    | 300    | 1800         |
| Item 4 | 400    | 400    | 400    | 1200    | 400    | 400    | 400    | 2400         |


---

## 🛠️ Personalización

Puedes adaptar el script para:

- Aplicar reglas más complejas
- Exportar el resultado a Excel o CSV

---

## 🚀 Ejecución

```python
import pandas as pd
from tabulate import tabulate

df_raw = pd.read_excel(url, header=1, nrows=4)
df_raw.columns = df_raw.columns.str.strip()
df = df_raw.iloc[:, 1:3]

for week in range(1, 7):
    df[f'Week {week}'] = df['Prices']

df['Total 1'] = df[[f'Week {i}' for i in range(1, 4)]].sum(axis=1)
df['Grand Total'] = df[[f'Week {i}' for i in range(1, 7)]].sum(axis=1)

cols = ['Item'] + \
       [f'Week {i}' for i in range(1, 4)] + \
       ['Total 1'] + \
       [f'Week {i}' for i in range(4, 7)] + \
       ['Grand Total']

df_final = df[cols]
print(tabulate(df_final, headers='keys', tablefmt='fancy_grid'))
```


### 💾 Exportación opcional
```python
# df_input.to_excel("weekly_matrix_table_output.xlsx", index=False)
```
---
### 📄 Licencia
---
Este proyecto está bajo ![License: MIT](https://img.shields.io/badge/License-MIT-cyan.svg). Puedes usarlo, modificarlo y distribuirlo libremente.

---
