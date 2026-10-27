"""Estado oficial de publicación CEN/CNE para EFE Valparaíso.

No convierte reparticiones agregadas del SEN en cargos de EFE. Consulta
fuentes públicas únicamente y muestra enlaces para cargar archivos validados.
"""
from __future__ import annotations

from datetime import date, datetime
from html.parser import HTMLParser
from urllib.parse import urljoin, urlparse
from urllib.request import Request, urlopen
import re
import unicodedata

import streamlit as st

_ROOT = 'https://www.coordinador.cl'
_REPARTO = (_ROOT + '/mercados/documentos/pagos-de-transmision/'
            'reparticion-de-cargos-unicos-de-transmision/{y}-reparticion-de-cargos-unicos-de-transmision/')
_PEAJES = (_ROOT + '/mercados/documentos/pagos-de-transmision/pagos-por-transmision-nacional/'
           'pagos-mensuales-de-peajes/{y}-pagos-mensuales-de-peajes/')
_SSCC = 'https://plabacom.coordinador.cl/'
_SERVICIO = _ROOT + '/coordinador/cargo-por-servicio-publico/'
_CNE_TRANSMISION = 'https://www.cne.cl/tarificacion/electrica/cargos-de-transmision/'
_CNE_PUBLICO = 'https://www.cne.cl/tarificacion/electrica/cargo-por-servicio-publico/'
_FUENTES = [
    ('Nacional · peajes de transmisión', _PEAJES, 'Liquidaciones mensuales de peajes (preliminar/definitiva)'),
    ('Nacional y Zonal · cargos únicos', _REPARTO, 'Repartición de cargos únicos de transmisión'),
    ('SSCC · pagos y BESS', _SSCC, 'Balances y detalle de clientes vía PLABACOM'),
    ('Cargo por Servicio Público', _SERVICIO, 'Recaudación agregada SEN, no cobro individual de EFE'),
    ('Tarifas de transmisión · CNE', _CNE_TRANSMISION, 'Resoluciones que fijan cargos de transmisión'),
    ('Servicio Público · CNE', _CNE_PUBLICO, 'Fundamento y disposiciones del cargo público'),
]
_MONTHS = dict(zip(
    ('ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO',
     'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'), range(1, 13)))


def _normalize(s):
    return ' '.join(unicodedata.normalize('NFKD', str(s)).encode('ASCII', 'ignore')
                    .decode('ASCII').upper().split())


class _Tokens(HTMLParser):
    """Extrae texto ordenado con los vínculos reales, sin paquetes nuevos."""
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.items = []
        self.href = None
        self.skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in ('script', 'style', 'noscript'):
            self.skip += 1
        if tag == 'a':
            self.href = dict(attrs).get('href')

    def handle_endtag(self, tag):
        if tag in ('script', 'style', 'noscript') and self.skip > 0:
            self.skip -= 1
        if tag == 'a':
            self.href = None

    def handle_data(self, data):
        if not self.skip:
            t = ' '.join(data.split())
            if t:
                self.items.append((t, self.href))


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def _page(url):
    req = Request(url, headers={'User-Agent': 'Mozilla/5.0 (compatible; EFE-CEN-Revisiones/1.0)',
                                'Accept': 'text/html,application/xhtml+xml'})
    with urlopen(req, timeout=14) as res:
        raw = res.read(3_000_000)
    # WordPress del CEN se publica en UTF-8.
    parsed = _Tokens()
    parsed.feed(raw.decode('utf-8', 'replace'))
    if len(parsed.items) < 4:
        raise ValueError('La página no entregó documentos legibles')
    return parsed.items


def _fecha(texto):
    f = re.search(r'(?<!\d)(\d{1,2})[/-](\d{1,2})[/-](20\d{2})(?!\d)', texto)
    if not f:
        return ''
    try:
        return date(int(f.group(3)), int(f.group(2)), int(f.group(1))).isoformat()
    except ValueError:
        return ''


def _releases_reparto(tokens, anio, source_url):
    """Detecta mes, estado, fecha y ZIP de la página de cargos únicos."""
    found = []
    for i, (text, href) in enumerate(tokens):
        s = _normalize(text)
        m = re.match(r'^(%s)\s+(DEFINITIVO|PRELIMINAR)\b' % '|'.join(_MONTHS), s)
        if not m:
            continue
        month = _MONTHS[m.group(1)]
        stage = m.group(2).title()
        published = ''
        link = ''
        for t, h in tokens[i+1:i+12]:
            normalized = _normalize(t)
            if re.match(r'^(%s)\s+(DEFINITIVO|PRELIMINAR)\b' % '|'.join(_MONTHS), normalized):
                break
            if not published and ('PUBLICACION' in normalized or _fecha(t)):
                published = _fecha(t)
            if h and ('DESCARGAR' in normalized or h.lower().endswith('.zip')):
                link = urljoin(source_url, h)
                break
        found.append({'Concepto': 'Transmisión nacional y zonal · reparto',
                      'Mes': f'{anio}-{month:02}', 'Estado': stage,
                      'Publicado': published, 'Documento': link or source_url,
                      'Verificación': 'Publicación oficial; importe EFE sin validar'})
    return found


def _months_nacional(tokens, anio, source_url):
    """Devuelve vínculos a cada mes de los peajes nacionales."""
    found = []
    for i, (text, _) in enumerate(tokens):
        s = _normalize(text)
        if s not in _MONTHS:
            continue
        link = ''
        for t, h in tokens[i+1:i+6]:
            if h and 'VER DOCUMENTOS' in _normalize(t):
                link = urljoin(source_url, h)
                break
        if link:
            found.append({'Concepto': 'Transmisión nacional · peajes',
                          'Mes': f'{anio}-{_MONTHS[s]:02}', 'Estado': 'Ver documento',
                          'Publicado': '', 'Documento': link,
                          'Verificación': 'Revisar liquidación preliminar/definitiva'})
    return found


def _detail_nacional(tokens, period, source_url):
    """Enlace a liquidaciones reales de cada mes, con preferencia por definitivas."""
    found = []
    for i, (title, _) in enumerate(tokens):
        text = _normalize(title)
        if 'LIQUIDACION DE PEAJES' not in text:
            continue
        if 'ANTECEDENTES' in text:
            continue
        stage = 'Definitivo' if 'DEFINITIV' in text else ('Preliminar' if 'PRELIMINAR' in text else 'Sin confirmar')
        published, href = '', ''
        for t, h in tokens[i+1:i+12]:
            if _fecha(t):
                published = _fecha(t)
            if h and 'DESCARGAR' in _normalize(t):
                href = urljoin(source_url, h)
                break
            if 'LIQUIDACION DE PEAJES' in _normalize(t):
                break
        if href:
            found.append({'Concepto': 'Transmisión nacional · peajes', 'Mes': period,
                          'Estado': stage, 'Publicado': published,
                          'Documento': href, 'Verificación': 'Publicación oficial; importe EFE sin validar'})
    return found


def _priorizar(releases):
    """Una referencia por mes y tipo; definitivo prevalece sobre preliminar."""
    best = {}
    order = {'Definitivo': 2, 'Preliminar': 1, 'Ver documento': 0}
    for row in releases:
        key = (row['Concepto'], row['Mes'])
        if key not in best or (order.get(row['Estado'], 0), row['Publicado']) > (order.get(best[key]['Estado'], 0), best[key]['Publicado']):
            best[key] = row
    return sorted(best.values(), key=lambda r: (r['Mes'], r['Concepto']), reverse=True)


def _identificar_inventario(inventario):
    """Resume el inventario sin asumir que una carga equivale a pago validado."""
    out = []
    if not inventario:
        return out
    for key, label in [('sscc9','SSCC detalle por retiro'), ('sscc4', 'SSCC remuneraciones'),
                       ('balance','Balance de energía'), ('exenciaci','Exenciones CI')]:
        items = inventario.get(key, [])
        for a in items:
            out.append({'Categoría': label, 'Período detectado': getattr(a, 'anio', '') or '',
                        'Archivo': str(getattr(a, 'name', '')), 'Origen': str(getattr(a, 'ruta', 'Datos/'))})
    return out


def mostrar_monitoreo_cen(inventario=None):
    st.header('📡 Publicaciones y actualizaciones CEN')
    st.caption('Fuentes: CEN y CNE. Revisa cambios de documentos preliminares a definitivos. '
               'Este módulo no inventa importes de Quilpué/Limache ni reemplaza cargos ya validados.')
    st.markdown('#### 1 · Fuentes oficiales')
    for label, url, purpose in _FUENTES:
        if '{y}' in url:
            url = url.format(y=datetime.now().year)
        st.markdown(f'**{label}** · [Abrir fuente oficial]({url})')
        st.caption(purpose)
    st.divider()
    st.markdown('#### 2 · Consulta de publicaciones de transmisión')
    years = st.multiselect('Años a revisar en el CEN', list(range(2024, datetime.now().year+1)),
                            default=[datetime.now().year], key='mon_cen_years')
    if st.button('🔄 Reconsultar páginas oficiales', key='mon_cen_refresh'):
        _page.clear()
        st.rerun()
    if not years:
        st.info('Selecciona algún año para consultar.')
        return
    reports, errors = [], []
    with st.spinner('Verificando publicaciones oficiales del CEN…'):
        for y in years:
            reparto = _REPARTO.format(y=y)
            try:
                reports.extend(_releases_reparto(_page(reparto), y, reparto))
            except Exception as exc:
                errors.append(f'Repartición de cargos {y}: {exc}')
            peajes = _PEAJES.format(y=y)
            try:
                entries = _months_nacional(_page(peajes), y, peajes)
                # Leer solo el mes más reciente para verificar cuál es definitivo;
                # los demás meses conservan enlaces al índice oficial.
                if entries:
                    newest = max(entries, key=lambda r: r['Mes'])
                    reports.extend(_detail_nacional(_page(newest['Documento']), newest['Mes'], newest['Documento']))
                    reports.extend(r for r in entries if r['Mes'] != newest['Mes'])
                else:
                    errors.append(f'Peajes nacionales {y}: sin meses reconocibles en la página')
            except Exception as exc:
                errors.append(f'Peajes nacionales {y}: {exc}')
    reports = _priorizar(reports)
    if reports:
        st.success(f'{len(reports)} períodos/documentos detectados en el sitio público.')
        st.dataframe(reports, hide_index=True, use_container_width=True,
                     column_config={'Documento': st.column_config.LinkColumn('Documento oficial', display_text='Abrir')})
        import csv
        from io import StringIO
        out = StringIO()
        w = csv.DictWriter(out, fieldnames=list(reports[0]), delimiter=';')
        w.writeheader(); w.writerows(reports)
        st.download_button('⬇ Descargar índice CEN (CSV)', ('\ufeff'+out.getvalue()).encode('utf-8'),
                           file_name='EFE_indice_actualizaciones_CEN.csv', mime='text/csv', key='mon_cen_export')
    else:
        st.warning('No fue posible construir el listado automáticamente. Usa los enlaces oficiales de arriba.')
    if errors:
        with st.expander('Ver advertencias de consulta'):
            for msg in errors:
                st.write(msg)
    st.divider()
    st.markdown('#### 3 · Estado de SSCC y Cargo por Servicio Público')
    st.info('**SSCC:** los balances nuevos se distribuyen mediante las descargas de PLABACOM. '
            'Sin una interfaz pública documentada para descargar archivos individuales de manera directa, '
            'no se consideran automáticamente cargados ni se sustituyen los archivos 4 y 9. '
            'Se mantiene BESS dentro del total SSCC, sin duplicarlo.')
    st.info('**Cargo por Servicio Público:** la tabla del CEN informa recaudación del sistema completo, '
            'no el monto individual cobrable a EFE. Para obtenerlo debe identificarse el cargo aplicable '
            'en la norma y el mecanismo de facturación de Enel.')
    inv = _identificar_inventario(inventario)
    if inv:
        with st.expander('Archivos CEN detectados actualmente en la aplicación'):
            st.dataframe(inv, hide_index=True, use_container_width=True)
    st.warning('**Los informes disponibles no equivalen a un nuevo cargo validado.** '
               'Para incorporarlo a un total por Quilpué/Limache hay que confirmar clave de retiro, '
               'vigencia, unidad, fecha de pago, documentos definitivos y evitar cargos duplicados. '
               'La consulta a CEN se actualiza al abrir esta pestaña (caché de 6 horas).')
