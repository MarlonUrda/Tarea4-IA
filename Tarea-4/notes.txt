import pandas as p
import openpyxl as op
import matplotlib.pyplot as plt
import twilio as tw

file = r"C:\Users\mgur2\OneDrive\Documentos\Ventas-Fundamentos.xlsx"
sheet = "VENTAS"

# Parametros para cada funciom.
# En un futuro esto se volvera dinamico.
sale_graph_params = {
  'date_column':'Fecha',
  'col':'Precio Venta Real', 
  'title':"Grafica de Ventas Totales", 
  'xlabel':'Fecha', 
  'ylabel':'Precio Venta Real', 
  'freq':'ME', 
  'plot_type':'line'
}

spy_graph_params = {
  'date_column':'Fecha',
  'col':'Precio Venta Real', 
  'title':"Ventas Anuales", 
  'xlabel':'Año', 
  'ylabel':'Ventas', 
  'freq':'YE', 
  'plot_type':'bar'
}

def readExcel(file, sheet):
  wb = op.load_workbook(file, data_only=True)
  page = wb[sheet]
  rows = list(page.rows)

  columns = [cell.value for cell in rows[0]]

  data = []

  for row in rows[1:]:
    row_data=([cell.value for cell in row][:len(columns)])
    data.append(row_data)
  
  df = p.DataFrame(data, columns=columns)
  df.dropna(axis=1, how='all', inplace=True)

  #Generacion del reporte financiero
  sales, spy = finantialReport(df)
  graph(df, **sale_graph_params)
  graph(df, **spy_graph_params, spy)

  return df

def finantialReport(df):

  #Total de ventas
  sales = df['Precio Venta Real'].sum()
  total_sales = round(sales, 2)
  print(f"Total Sales: {total_sales}")

  df['Fecha'] = p.to_datetime(df['Fecha'])
  spy = df.groupby(df['Fecha'].dt.year).size()

  print(f"Total Sales per Year: {spy}")

  return total_sales, spy

def graph(df, date_column, col, title, xlabel, ylabel, freq, plot_type, spy=None):

  #Nos aseguramos que la columna de fechas este en el formato correcto
  df[date_column] = p.to_datetime(df[date_column])

  #Agrupamos
  group = df.groupby(p.Grouper(key=date_column, freq=freq))[col].sum()

  #Grafica
  plt.figure(figsize=(10, 6))

  #Determinar tipo de grafica
  if plot_type == 'line':
    plt.plot(group.index, group.values, marker='o', linestyle='-', color='b')
  elif plot_type == 'bar':
    if spy is not None:
      plt.bar(group.index, spy.values, color='g')
    else:
      plt.bar(group.index, group.values, color='g')
  elif plot_type == 'hist':
    plt.hist(group.values, bins=20, color='r')
  elif plot_type == 'pie':
    plt.pie(group.values, labels=group.index, autopct='%1.1f%%')
  else:
    print(f"Tipo de grafica: {plot_type}. No soportado")
    return
  
  #Dibujar
  plt.title(title)
  plt.xlabel(xlabel)
  plt.ylabel(ylabel)
  plt.grid(True)

  if plot_type != 'hist':
    plt.xticks(rotation=45)

  plt.ticklabel_format(style="plain", axis="y")

  plt.tight_layout()
  plt.show()


readExcel(file, sheet)
