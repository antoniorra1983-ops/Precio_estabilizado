"""Cálculo de referencia del Archivo 4 del CEN: SSCC + CO netos.

PAGA y RECIBE corresponden a Enel Generación en el SEN, NO a EFE.
CEN registra NETO = RECIBE - PAGA; para costo se usa PAGA - RECIBE.
Los saldos pueden ser negativos: no aplicar abs() ni forzar a cero.
"""
from __future__ import annotations

import math


def calcular_netos_empresa(empresa: dict) -> dict:
    """Obtiene los dos componentes netos y valida la convención del Archivo 4.

    Un resultado negativo es un ingreso neto de Enel y reduce el costo total.
    No es una liquidación oficial atribuible a EFE Valparaíso.
    """
    componentes = {}
    for tipo in ("SC_SSCC", "CO"):
        datos = empresa.get(tipo)
        if not isinstance(datos, dict):
            raise ValueError(f"Archivo 4: falta el componente {tipo}.")
        try:
            paga, recibe = float(datos["paga"]), float(datos["recibe"])
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(f"Archivo 4: importes de PAGA/RECIBE inválidos para {tipo}.") from exc
        if not math.isfinite(paga) or not math.isfinite(recibe):
            raise ValueError(f"Archivo 4: importes no finitos para {tipo}.")
        costo_neto = paga - recibe
        # La hoja del CEN publica D=RECIBE-PAGA. Respetar su signo inverso.
        if datos.get("neto") is not None:
            saldo_cen = float(datos["neto"])
            if not math.isfinite(saldo_cen) or abs(saldo_cen + costo_neto) > 0.05:
                raise ValueError(
                    f"Archivo 4: NETO de {tipo} no concuerda con RECIBE-PAGA; "
                    "revisar columnas B, C y D.")
        componentes[tipo] = {"paga": paga, "recibe": recibe,
                              "costo_neto": costo_neto, "neto_cen": -costo_neto}
    componentes["total_paga"] = sum(componentes[t]["paga"] for t in ("SC_SSCC", "CO"))
    componentes["total_recibe"] = sum(componentes[t]["recibe"] for t in ("SC_SSCC", "CO"))
    componentes["total_costo_neto"] = sum(componentes[t]["costo_neto"] for t in ("SC_SSCC", "CO"))
    return componentes
