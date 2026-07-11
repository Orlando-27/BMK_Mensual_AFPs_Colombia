"""
Pruebas unitarias de la logica determinista del pipeline (sin insumos pesados).
Ejecutar: python3 -m pytest tests/ -q
"""
from bmk.classify.classifier import key_R, NEMOS_ESPECIALES
from bmk.derivatives.forwards import _compra_venta, _par
from bmk.derivatives.swaps import clasificar_tipo
from bmk.prepare.normalize import normalizar_afp, _cod_portafolio


def test_key_R_isin_directo_para_tdpe_findi():
    # Clas SFC TDPE/FINDI -> usa ISIN (S)
    assert key_R("TDPE", "US1234567890", "NEMO1", "EMISOR") == "US1234567890"
    assert key_R("FINDI", "LU0000000001", "", "") == "LU0000000001"


def test_key_R_nemo_especial():
    # ISIN especial se mapea a su nemo (Diccionario SFC BD:BE)
    isin = "CL0002892397"  # -> NUAMCO
    assert key_R("BOEVS", isin, "", "") == NEMOS_ESPECIALES[isin]


def test_key_R_prioriza_nemo():
    # Caso general: usa nemo (AQ) cuando existe
    assert key_R("BOEVS", "US1234567890", "ECOPETROL", "EMISOR") == "ECOPETROL"


def test_normalizar_afp():
    assert normalizar_afp('"PORVENIR"') == "PORVENIR"
    assert normalizar_afp('"COLFONDOS S.A." Y "COLFONDOS"') == "COLFONDOS"
    assert normalizar_afp("SKANDIA AFP - ACCAI S.A.") == "SKANDIA"
    assert normalizar_afp("OLD MUTUAL PENSIONES") == "OLD MUTUAL"


def test_cod_portafolio():
    assert _cod_portafolio("FONDO DE PENSIONES OBLIGATORIAS MODERADO") == "PO"
    assert _cod_portafolio("FONDO DE PENSIONES OBLIGATORIAS MAYOR RIESGO") == "PM"
    assert _cod_portafolio("FONDO DE CESANTIAS LEY 50") == "CS"
    assert _cod_portafolio("FONDO ESPECIAL RETIRO PROGRAMADO") == "PR"


def test_forward_compra_venta():
    assert _compra_venta("USD", "COP") == "VENTA"   # obligacion COP -> venta
    assert _compra_venta("COP", "USD") == "COMPRA"  # derecho COP -> compra
    assert _compra_venta("EUR", "USD") == "COMPRA"  # der<>COP y obl=USD -> compra


def test_forward_par():
    assert _par("USD", "COP", "VENTA") == "USDCOP"
    assert _par("COP", "USD", "COMPRA") == "USDCOP"


def test_clasificar_tipo_swap():
    assert clasificar_tipo("COP", "COP") == "IRS IBRCOP"
    assert clasificar_tipo("COP", "USD") == "CCS COPUSD"
    assert clasificar_tipo("USD", "COP") == "CCS COPUSD"
    assert clasificar_tipo("COP", "COU") == "CCS COPUVR"  # COU -> UVR
    assert clasificar_tipo("USD", "USD") == "IRS SOFUSD"
