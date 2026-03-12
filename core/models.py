"""
Modelos de datos para el extractor CORREAGRO
"""
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional


@dataclass
class RowOpVigente:
    operacion: str
    doc_cruce: int
    fecha_doc: str
    razon_social: str
    fecha_vencimiento: str
    saldo_final: float


@dataclass
class ClienteOpVigente:
    nit: str
    razon_social: str
    rows: List[RowOpVigente] = field(default_factory=list)

    @property
    def total_saldo(self) -> float:
        return sum(r.saldo_final for r in self.rows)


@dataclass
class RowFisicoCompra:
    fecha_docto: str
    fecha_vcto: str
    docto: str
    num: int
    notas: str
    debitos: float
    creditos: float


@dataclass
class ClienteFisicoCompra:
    nit: str
    razon_social: str
    saldo_inicial: float
    entradas: float
    salidas: float
    saldo_final: float
    rows: List[RowFisicoCompra] = field(default_factory=list)

    @property
    def total_debitos(self) -> float:
        return sum(r.debitos for r in self.rows)

    @property
    def total_creditos(self) -> float:
        return sum(r.creditos for r in self.rows)


@dataclass
class RowCartera:
    documento: str
    fecha_contabilizacion: str
    fecha_vencimiento: str
    dias: int
    corriente: float
    de_1_30: float
    de_31_60: float
    de_61_90: float
    de_91_9999: float
    total: float


@dataclass
class ClienteCartera:
    nit: str
    razon_social: str
    rows: List[RowCartera] = field(default_factory=list)

    @property
    def total_general(self) -> float:
        return sum(r.total for r in self.rows)


@dataclass
class CarteraData:
    excel_path: str
    corte: Optional[datetime] = None
    cuenta: str = ""
    clientes: List[ClienteCartera] = field(default_factory=list)


@dataclass
class InformeData:
    """Contenedor principal de todos los datos del informe"""
    excel_path: str
    desde: Optional[datetime] = None
    hasta: Optional[datetime] = None
    clientes_op_vigentes: List[ClienteOpVigente] = field(default_factory=list)
    clientes_fisicos_compras: List[ClienteFisicoCompra] = field(default_factory=list)
