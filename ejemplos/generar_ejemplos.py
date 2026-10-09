"""Genera archivos de ejemplo con estructuras distintas para probar el generador."""
from pathlib import Path

import numpy as np
import pandas as pd

AQUI = Path(__file__).parent
rng = np.random.default_rng(7)


def ventas():
    n = 1500
    fechas = pd.to_datetime("2024-01-01") + pd.to_timedelta(rng.integers(0, 640, n), unit="D")
    region = rng.choice(["Norte", "Sur", "Centro", "Oriente", "Occidente"], n, p=[.3, .2, .25, .15, .1])
    cat = rng.choice(["Tecnología", "Hogar", "Ropa", "Deportes", "Juguetes", "Libros", "Belleza"], n)
    unidades = rng.integers(1, 40, n)
    precio = rng.choice([12.5, 25, 49.9, 89, 150, 320], n)
    estacional = 1 + 0.35 * np.sin((fechas.month.values - 3) / 12 * 2 * np.pi)
    df = pd.DataFrame({
        "ID Pedido": np.arange(10001, 10001 + n),
        "Fecha": fechas, "Región": region, "Categoría": cat,
        "Vendedor": rng.choice(["Ana", "Luis", "Marta", "Jorge", "Sofía", "Pedro"], n),
        "Unidades": unidades, "Precio unitario": precio,
        "Ventas": np.round(unidades * precio * estacional, 2),
    }).sort_values("Fecha")
    # Reporte "real": título arriba, filas vacías y una fila de total al final.
    with pd.ExcelWriter(AQUI / "ventas.xlsx") as w:
        pd.DataFrame([["Reporte de ventas 2024-2025"], [None]]).to_excel(w, sheet_name="Reporte", header=False, index=False)
        df.to_excel(w, sheet_name="Reporte", startrow=3, index=False)
        pd.DataFrame([["Total general", None, None, None, None, df.Unidades.sum(), None, df.Ventas.sum()]]).to_excel(
            w, sheet_name="Reporte", startrow=len(df) + 4, header=False, index=False)
        pd.DataFrame({"Nota": ["hoja auxiliar"]}).to_excel(w, sheet_name="Notas", index=False)


def encuesta_csv():
    """CSV con ';', decimales con coma y fechas dd/mm/aaaa, como lo exporta Excel en español."""
    n = 400
    fechas = pd.to_datetime("2025-03-01") + pd.to_timedelta(rng.integers(0, 45, n), unit="D")
    df = pd.DataFrame({
        "Fecha respuesta": fechas.strftime("%d/%m/%Y"),
        "Ciudad": rng.choice(["Bogotá", "Medellín", "Cali", "Barranquilla"], n),
        "Canal": rng.choice(["Web", "App", "Tienda"], n),
        "Satisfacción": rng.integers(1, 11, n),
        "Gasto": [f"$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") for x in rng.gamma(2, 80000, n)],
    })
    df.to_csv(AQUI / "encuesta.csv", sep=";", index=False, encoding="utf-8")


def inventario():
    """Sin fechas: solo categorías y números."""
    n = 120
    pd.DataFrame({
        "SKU": [f"SKU-{i:04d}" for i in range(n)],
        "Bodega": rng.choice(["Principal", "Norte", "Sur"], n),
        "Familia": rng.choice(["Tornillería", "Eléctrico", "Pintura", "Plomería", "Herramientas"], n),
        "Stock": rng.integers(0, 900, n),
        "Costo": np.round(rng.uniform(1, 250, n), 2),
    }).to_excel(AQUI / "inventario.xlsx", index=False)


if __name__ == "__main__":
    ventas(); encuesta_csv(); inventario()
    print("Ejemplos creados en", AQUI)
