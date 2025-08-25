import pandas
import numpy
import matplotlib.pyplot as plt
from PIL.Image import Image as ExcelImage
import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.drawing.image import Image

df = pandas.read_csv("ventas_ejemplo.csv")
df['fecha'] = pandas.to_datetime(df['fecha'])


def ventas_por_dia(df):
    """calcular ventas"""
    return df.groupby(df['fecha'].dt.date)['total'].sum()

def producto_mas_vendido(df):
    """producto mas vendido"""
    productos = df.groupby('nombre_producto')['cantidad'].sum().sort_values(ascending=False)
    return productos.index[0], productos.iloc[0]

def vendedor_con_mas_ingresos(df):
    """el vendedos mvp"""
    vendedores = df.groupby('vendedor')['total'].sum().sort_values(ascending=False)
    return vendedores.index[0], vendedores.iloc[0]

def generar_reporte_semanal(df):
    """genera reporte completo"""
    print("=" * 50)
    print("   Reporte Semanal de Ventas")

    #ventas por dia
    print("\n📅 VENTAS POR DÍA:")
    ventas_dia = ventas_por_dia(df)
    for fecha, total in ventas_dia.items():
        print(f" {fecha}: ${total:,.2f}")

    #producto mas vendido
    item, cantidad = producto_mas_vendido(df)
    print(f"\n🏆 PRODUCTO MÁS VENDIDO:")
    print(f"  {item}: {cantidad} unidades")

    vendedor, ingresos = vendedor_con_mas_ingresos(df)
    print(f"\n🏆 VENDEDOR MVP:")
    print(f"  {vendedor}: ${ingresos:,.2f}")

    #resumen total
    total_semana = df['total'].sum()
    print(f"\n 📊 TOTAL SEMANAL: ${total_semana:,.2f}")
    print("=" * 50)

def crear_grafica(df, nombre_grafica ="ventas_por_dia.png"):
    """cra la grafica y la guarda como imagen"""

    ventas_dia = ventas_por_dia(df)

    #grafica
    plt.figure(figsize=(10,6))
    plt.plot(ventas_dia.index, ventas_dia.values, marker='o', linewidth=2, markersize=8)
    plt.title('ventas por dia', fontsize=12, fontweight='bold')
    plt.xlabel('fecha', fontsize=12)
    plt.grid(True, alpha=0.3)


    plt.ticklabel_format(style='plain', axis='y')

    plt.xticks(rotation=45)
    plt.tight_layout()

    plt.savefig(nombre_grafica, dpi=300, bbox_inches='tight')
    plt.close()

    return nombre_grafica

def generar_excel_completo(df, nombre_archivo="reporte_Semanal.xlsx"):
    """generacion del excel"""

    # 1. Crear la gráfica primero usando TU función
    archivo_grafica = crear_grafica(df)

    #crear excel
    wb = Workbook()
    ws = wb.active
    ws.title = "reporte semanal"

    #titula
    ws['A1'] = "Reporte Semanal de ventas"
    ws['A1'].font = Font(size=18, bold=True)
    ws['A1'].alignment = Alignment(horizontal='center')
    ws.merge_cells('A1:F1')

    #insertar grafica
    img = Image(archivo_grafica)
    img.width = 600
    img.height = 400
    ws.add_image(img, 'A3')

    #tabla resumen
    fila_inicio = 25

    ws[f'A{fila_inicio}'] = "Resuman ejecutivo"
    ws[f'A{fila_inicio}'].font = Font(size=14, bold=True)
    fila_inicio += 2

    #usar las funciones

    producto, cantidad, = producto_mas_vendido(df)
    vendedor, ingresos = vendedor_con_mas_ingresos(df)
    total_semana = df['total'].sum()

    ws[f'A{fila_inicio}'] = "producto mas vendido"
    ws[f'B{fila_inicio}'] = f"{producto} ({cantidad} unidades)"

    ws[f'A{fila_inicio+1}'] = "MVP"
    ws[f'B{fila_inicio+1}'] = f"{vendedor} (${ingresos:,.2f})"

    ws[f'A{fila_inicio+2}'] = "total semana"
    ws[f'B{fila_inicio+2}'] = f"${total_semana:,.2f}"
    ws[f'B{fila_inicio+2}'].font = Font(bold=True)

    fila_tabla = fila_inicio + 5

    ws[f'A{fila_tabla}'] = "VENTAS DETALLADAS POR DÍA"
    ws[f'A{fila_tabla}'].font = Font(size=14, bold=True)
    fila_tabla += 2

    # Headers de la tabla
    ws[f'A{fila_tabla}'] = "Fecha"
    ws[f'B{fila_tabla}'] = "Total Ventas"
    ws[f'A{fila_tabla}'].font = Font(bold=True)
    ws[f'B{fila_tabla}'].font = Font(bold=True)
    fila_tabla += 1

    #llenar la tabla
    ventas_dia = ventas_por_dia(df)
    for fecha, total in ventas_dia.items():
        ws[f'A{fila_tabla}'] = str(fecha)
        ws[f'B{fila_tabla}'] = total
        fila_tabla += 1

    wb.save(nombre_archivo)
    print(f"EXcel generado: {nombre_archivo}")
    print(f"🎨 Gráfica creada: {archivo_grafica}")

    return nombre_archivo


generar_excel_completo(df)











