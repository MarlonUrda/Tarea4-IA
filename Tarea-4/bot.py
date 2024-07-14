import openpyxl
import numpy as n
import pandas as p
import matplotlib.pyplot as pl
from matplotlib.backends.backend_pdf import PdfPages

years = [2015, 2016, 2017]

def generate_report(sales_df, path):
  with PdfPages(path) as pdf:
    for year in years:
      
      sales_per_year = sales_df[sales_df['Año'] == year]

      total_cost_v = sales_per_year['Costo Vehículo'].sum()
      total_without_igv = sales_per_year['Precio Venta sin IGV'].sum()
      number_sales = sales_per_year.shape[0]
      real_price = sales_per_year['Precio Venta Real'].sum()
      income = real_price - total_cost_v

      pl.figure(figsize=(35, 12))
      pl.text(0.5, 0.9, f"Sales Review for {year}", fontsize=80, ha='center', va='center')

      pl.text(0.5, 0.65, f"Total cost in {year}: ${total_cost_v:,.2f}", fontsize=36, ha='center', va='center')
      pl.text(0.5, 0.59, f"Total cost without igv in {year}: ${total_without_igv:,.2f}", fontsize=36, ha='center', va='center')
      pl.text(0.5, 0.53, f"Number of sales in {year}: {number_sales}", fontsize=36, ha='center', va='center')
      pl.text(0.5, 0.30, f"Net sales in {year}: ${real_price:,.2f} with 18% added due to IGV", fontsize=36, ha='center', va='center')
      pl.text(0.5, 0.23, f"Income {year}: ${income:,.2f}", fontsize=46, ha='center', va='center')

      pl.axis('off')
      pdf.savefig()
      pl.close()

      fig, ax = pl.subplots(nrows=1, ncols=3, figsize=(35, 12))
      fig.suptitle(f"Finantials for {year}", fontsize=42, fontstyle="italic")
      fig.tight_layout(rect=[0, 0, 1, 0.975])

      def autopct_func(p, vals):
        absolute = int(round(p/100.*n.sum(vals)))
        return "{:.1f}%\n\n(Total: {})".format(p, absolute)
      
      def autopct_sale(p, vals):
        absolute = int(round(p/100.*n.sum(vals)))
        return "{:.1f}%\n\n${:,}".format(p, absolute)
      
      sums = sales_per_year.groupby('Canal')['Precio Venta Real'].sum()
      bars = ax[0].bar(sums.index, sums.values, color='red')
      ax[0].set_title(f"Sales by Channel in {year}")
      ax[0].set_xticks(range(len(sums.index)))
      ax[0].set_xticklabels(sums.index, rotation=45)

      for bar in bars:
        yval = bar.get_height()
        format_yval = "${:,.2}".format(yval)
        ax[0].text(bar.get_x() + bar.get_width()/2, yval, format_yval, ha="center", va="bottom")
      
      segment_sums = sales_per_year.groupby('Segmento')['Precio Venta Real'].sum()
      colors = ['blue', 'red']
      ax[1].pie(segment_sums.values, labels=segment_sums.index, autopct=lambda pct: autopct_sale(pct, segment_sums.values), startangle=90, colors=colors)
      ax[1].set_title(f"Sales by Segment in {year}")  

      sede_sums = sales_per_year.groupby('Sede')['Precio Venta Real'].sum()
      bars = ax[2].bar(sede_sums.index, sede_sums.values, color="blue")
      ax[2].set_title(f"Sales by Location in {year}")
      ax[2].set_xticks(range(len(sede_sums.index)))
      ax[2].set_xticklabels(sede_sums.index, rotation=45)
      
      for bar in bars:
        yval = bar.get_height()
        format_yval = "${:,.2}".format(yval)
        ax[2].text(bar.get_x() + bar.get_width()/2, yval, format_yval, ha="center", va="bottom")
      
      pl.tight_layout(rect=[0,0,1,0.975])
      pdf.savefig()
      pl.close(fig)

      fig, ax = pl.subplots(nrows=1, ncols=3, figsize=(35, 12))
      fig.suptitle(f"Statistics for {year}", fontsize=42, fontstyle="italic")
      fig.tight_layout(rect=[0, 0, 1, 0.975])

      channel_counts = sales_per_year["Canal"].value_counts()
      bars = ax[0].bar(channel_counts.index, channel_counts.values, color="green")
      ax[0].set_title(f"Number of Sales per Channel in {year}")
      ax[0].set_xticks(range(len(channel_counts.index)))
      ax[0].set_xticklabels(channel_counts.index, rotation=45)

      for bar in bars:
        yval = bar.get_height()
        ax[0].text(bar.get_x() + bar.get_width()/2, yval, yval, ha='center', va='bottom')
      
      segment_counts = sales_per_year["Segmento"].value_counts()
      colors2 = ['green', 'orange']
      ax[1].pie(segment_counts.values, labels=segment_counts.index, autopct=lambda pct: autopct_func(pct, segment_counts.values), startangle=90, colors=colors2)
      ax[1].set_title(f"Sales per Segment in {year}")  
      
      sede_counts = sales_per_year["Sede"].value_counts()
      bars = ax[2].bar(sede_counts.index, sede_counts.values, color="orange")
      ax[2].set_title(f"Sales per Location in {year}")
      ax[2].set_xticks(range(len(sede_counts.index)))
      ax[2].set_xticklabels(sede_counts.index, rotation=45)

      for bar in bars:
        yval = bar.get_height()
        ax[2].text(bar.get_x() + bar.get_width()/2, yval, yval, ha='center', va='bottom')
      
      pl.tight_layout(rect=[0,0,1,0.975])
      pdf.savefig()
      pl.close(fig)

file = r"C:\Users\mgur2\OneDrive\Documentos\Ventas-Fundamentos.xlsx"
wb = openpyxl.load_workbook(file)
sheet = 'VENTAS'

sales_df = p.read_excel(file, sheet_name=sheet, engine='openpyxl')
sales_df['Año'] = p.to_datetime(sales_df['Fecha']).dt.year

generate_report(sales_df, "Report.pdf")