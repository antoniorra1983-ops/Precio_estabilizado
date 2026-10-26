"""Divisor de archivos CEN para GitHub / Streamlit.

Aplicación independiente: desplegar Divisor_Excel/app.py como segunda app de
Streamlit Community Cloud. No requiere cargar ni modificar la calculadora EFE.
Los archivos se procesan únicamente durante la sesión y se descargan desde
el navegador; esta herramienta NO sube archivos automáticamente a GitHub.
"""
from __future__ import annotations

from datetime import datetime
from hashlib import sha256
from io import BytesIO
from pathlib import Path, PurePosixPath
import re
from typing import BinaryIO
from zipfile import ZipFile, ZIP_STORED

MIB = 1024 * 1024
DEFAULT_PART_MIB = 20
GITHUB_WEB_LIMIT_MIB = 25
READ_BLOCK = 1024 * 1024
FILE_TYPES = {".xlsx", ".xlsm", ".csv"}
CATEGORIES = (
    "Balances",
    "Precios_estabilizados",
    "Sscc_archivo4",
    "Sscc_archivo9",
    "Potencia",
    "Mt_pd",
    "Horarios",
)
PART_PATTERN = re.compile(r"^(.*\.(?:xlsx|xlsm|csv))\.part(\d{3})$", re.IGNORECASE)


def checked_name(name: str) -> str:
    """Conserva el nombre real del archivo, sin permitir rutas inseguras."""
    cleaned = str(name).replace("\\", "/").split("/")[-1].strip()
    if not cleaned or cleaned in {".", ".."} or "\x00" in cleaned:
        raise ValueError("Nombre de archivo no válido.")
    if Path(cleaned).suffix.casefold() not in FILE_TYPES:
        raise ValueError("Solo se admiten archivos .xlsx, .xlsm y .csv.")
    return cleaned


def make_parts_zip(
    input_file: BinaryIO,
    filename: str,
    category: str,
    year: int,
    part_size_mib: int = DEFAULT_PART_MIB,
    split_always: bool = False,
) -> dict:
    """Devuelve ZIP con Datos/<tipo>/<año>/archivo(.partNNN) y hash verificado.

    Lee por bloques y almacena las partes dentro de un ZIP sin comprimir los
    libros Excel de nuevo. La integridad se comprueba releyendo los miembros.
    """
    filename = checked_name(filename)
    if category not in CATEGORIES:
        raise ValueError("Selecciona una categoría válida.")
    if not (2000 <= int(year) <= 2100):
        raise ValueError("El año debe estar entre 2000 y 2100.")
    if not 1 <= int(part_size_mib) <= 24:
        raise ValueError("Selecciona partes entre 1 y 24 MiB.")
    part_size = int(part_size_mib) * MIB
    prefix = PurePosixPath("Datos", category, str(year))

    input_file.seek(0, 2)
    total_bytes = input_file.tell()
    input_file.seek(0)
    if total_bytes <= 0:
        raise ValueError("El archivo está vacío.")
    # Si cabe en GitHub web, preferimos el formato Excel original.
    must_split = split_always or total_bytes > GITHUB_WEB_LIMIT_MIB * MIB
    zip_buffer = BytesIO()
    filenames = []
    sizes = []
    original_hash = sha256()
    with ZipFile(zip_buffer, "w", compression=ZIP_STORED, allowZip64=True) as zout:
        if must_split:
            number = 1
            while True:
                first = input_file.read(min(READ_BLOCK, part_size))
                if not first:
                    break
                if number > 999:
                    raise ValueError("El archivo requiere más de 999 partes.")
                member = str(prefix / f"{filename}.part{number:03d}")
                part_written = 0
                with zout.open(member, "w", force_zip64=True) as part:
                    part.write(first)
                    original_hash.update(first)
                    part_written += len(first)
                    while part_written < part_size:
                        next_block = input_file.read(min(READ_BLOCK, part_size - part_written))
                        if not next_block:
                            break
                        part.write(next_block)
                        original_hash.update(next_block)
                        part_written += len(next_block)
                filenames.append(member)
                sizes.append(part_written)
                number += 1
        else:
            member = str(prefix / filename)
            written = 0
            with zout.open(member, "w", force_zip64=True) as full:
                while block := input_file.read(READ_BLOCK):
                    full.write(block)
                    original_hash.update(block)
                    written += len(block)
            filenames.append(member)
            sizes.append(written)

        instructions = (
            "DIVISOR DE ARCHIVOS CEN PARA GITHUB\n"
            "===================================\n"
            f"Archivo original: {filename}\n"
            f"Carpeta destino: {prefix}/\n"
            f"Número de archivos a subir: {len(filenames)}\n"
            f"Tamaño original (bytes): {total_bytes}\n"
            f"SHA256 original: {original_hash.hexdigest()}\n\n"
            "DESCOMPRIME ESTE ZIP. No subas el ZIP.\n"
            f"En GitHub, abre la carpeta {prefix}/ y sube solo los archivos \n"
            "que contiene esa carpeta, manteniendo TODOS los .part001, .part002, etc.\n"
            "No cambies nombres, no olvides partes ni agregues el libro original duplicado.\n"
            "El app.py de EFE que reconoce .part001 reconstruirá el Excel al leerlo.\n"
            "Los datos subidos a un repositorio público serán accesibles públicamente.\n"
        )
        zout.writestr("INSTRUCCIONES_NO_SUBIR_A_DATOS.txt", instructions)

    # Verificar el hash de la concatenación desde el ZIP creado.
    zip_buffer.seek(0)
    reconstructed_hash = sha256()
    reconstructed_size = 0
    with ZipFile(zip_buffer, "r") as zcheck:
        for member in filenames:
            with zcheck.open(member, "r") as part:
                while block := part.read(READ_BLOCK):
                    reconstructed_hash.update(block)
                    reconstructed_size += len(block)
    if reconstructed_size != total_bytes or reconstructed_hash.digest() != original_hash.digest():
        raise RuntimeError("Falló la comprobación de integridad. No subas las partes.")

    return {
        "zip_bytes": zip_buffer.getvalue(),
        "original_name": filename,
        "path": str(prefix),
        "parts": filenames,
        "sizes": sizes,
        "total_bytes": total_bytes,
        "sha256": original_hash.hexdigest(),
        "is_split": must_split,
        "valid": True,
        "year": year,
        "category": category,
    }


def reconstruct_parts(part_files: list[tuple[str, BinaryIO]], expected_hash: str = "") -> dict:
    """Reconstruye solo secuencias part001...partNNN sin saltos."""
    if not part_files:
        raise ValueError("Selecciona al menos una parte.")
    parsed: list[tuple[int, str, BinaryIO]] = []
    for name, input_file in part_files:
        safe = str(name).replace("\\", "/").split("/")[-1]
        match = PART_PATTERN.fullmatch(safe)
        if not match:
            raise ValueError(f"Nombre inválido: {safe}. Debe terminar en .part001, .part002, etc.")
        parsed.append((int(match.group(2)), match.group(1), input_file))
    names = {x[1] for x in parsed}
    if len(names) != 1:
        raise ValueError("Las partes corresponden a archivos distintos.")
    parsed.sort(key=lambda x: x[0])
    ids = [x[0] for x in parsed]
    if ids != list(range(1, len(ids) + 1)):
        raise ValueError(f"Faltan partes o hay repetidas. Recibidas: {ids}.")
    name = names.pop()
    out = BytesIO()
    digest = sha256()
    for _, _, part in parsed:
        part.seek(0)
        while block := part.read(READ_BLOCK):
            out.write(block)
            digest.update(block)
    hash_hex = digest.hexdigest()
    if expected_hash.strip() and hash_hex.lower() != expected_hash.strip().lower():
        raise ValueError("SHA-256 diferente al original. Revisa las partes seleccionadas.")
    return {"original_name": name, "bytes": out.getvalue(), "sha256": hash_hex, "num_parts": len(parsed)}


def main() -> None:
    import streamlit as st

    st.set_page_config(page_title="Divisor Excel | EFE Valparaíso", page_icon="📂", layout="centered")
    st.title("📂 Divisor de archivos CEN")
    st.caption("EFE Valparaíso · Archivos Excel grandes para GitHub")
    st.write(
        "Divide archivos de más de **25 MiB** en partes pequeñas que tu calculadora "
        "CEN puede reconstruir, sin alterar el contenido del Excel."
    )
    tab_split, tab_rebuild, tab_guide = st.tabs(["✂️ Dividir", "🔄 Reconstruir", "📖 Instrucciones"])

    with tab_split:
        st.subheader("1. Selecciona el archivo")
        with st.form("divide_form"):
            uploaded = st.file_uploader(
                "Excel del CEN", type=["xlsx", "xlsm", "csv"],
                help="Los datos se procesan temporalmente en el servidor de Streamlit; no se suben a GitHub automáticamente.",
            )
            category = st.selectbox("Carpeta de destino", CATEGORIES, index=CATEGORIES.index("Sscc_archivo9"))
            year = st.number_input("Año del balance", min_value=2000, max_value=2100, value=datetime.now().year, step=1)
            part_size = st.select_slider("Tamaño de cada parte", options=[5, 10, 15, 20, 24], value=20, format_func=lambda n: f"{n} MiB")
            split_always = st.checkbox("Dividir aunque el archivo mida menos de 25 MiB", value=False)
            submitted = st.form_submit_button("Preparar partes para GitHub", type="primary", use_container_width=True)
        if submitted:
            st.session_state.pop("divider_result", None)
            if uploaded is None:
                st.warning("Primero selecciona un Excel del CEN.")
            else:
                try:
                    with st.spinner("Dividiendo y comprobando la integridad del archivo…"):
                        st.session_state["divider_result"] = make_parts_zip(
                            uploaded, uploaded.name, category, int(year), int(part_size), split_always
                        )
                except Exception as exc:
                    st.error(f"No se pudo preparar el archivo: {exc}")
        result = st.session_state.get("divider_result")
        if result:
            st.success("Archivo dividido y verificado (SHA-256 coincide)." if result["is_split"] else "Archivo listo: no necesitó división.")
            c1, c2, c3 = st.columns(3)
            c1.metric("Tamaño original", f"{result['total_bytes']/MIB:,.1f} MiB")
            c2.metric("Archivos", str(len(result["parts"])))
            c3.metric("Integridad", "Correcta")
            st.write("**Destino en GitHub:**", f"`{result['path']}/`")
            st.dataframe(
                [{"Nombre del archivo": name.rsplit("/", 1)[-1], "Tamaño (MiB)": round(size / MIB, 2)}
                 for name, size in zip(result["parts"], result["sizes"])],
                hide_index=True, use_container_width=True,
            )
            st.download_button(
                "⬇️ Descargar ZIP con las partes", data=result["zip_bytes"],
                file_name=f"Partes_GitHub_{Path(result['original_name']).stem}_{result['year']}.zip",
                mime="application/zip", type="primary", use_container_width=True,
            )
            st.caption("Descomprime el ZIP y sube las partes a GitHub. El ZIP no se sube tal cual.")
            with st.expander("Comprobación SHA-256"):
                st.code(result["sha256"], language=None)
            if st.checkbox("Mostrar descargas individuales", key="individual_buttons"):
                with ZipFile(BytesIO(result["zip_bytes"]), "r") as zin:
                    for name in result["parts"]:
                        st.download_button(
                            f"Descargar {name.rsplit('/', 1)[-1]}",
                            data=zin.read(name),
                            file_name=name.rsplit("/", 1)[-1],
                            mime="application/octet-stream",
                            key=f"download_{name}", use_container_width=True,
                        )

    with tab_rebuild:
        st.subheader("Reconstruir un Excel desde sus partes")
        st.caption("Útil para comprobar las partes antes de subirlas, o recuperar el archivo original.")
        with st.form("rebuild_form"):
            pieces = st.file_uploader(
                "Selecciona todas las partes .part001, .part002, etc.",
                accept_multiple_files=True, key="rebuild_files",
            )
            hash_expected = st.text_input("SHA-256 original (opcional)")
            rebuild_submit = st.form_submit_button("Reconstruir y verificar", type="primary", use_container_width=True)
        if rebuild_submit:
            st.session_state.pop("rebuild_result", None)
            try:
                with st.spinner("Reconstruyendo el archivo…"):
                    st.session_state["rebuild_result"] = reconstruct_parts(
                        [(p.name, p) for p in pieces], hash_expected
                    )
            except Exception as exc:
                st.error(f"No se pudo reconstruir: {exc}")
        rebuilt = st.session_state.get("rebuild_result")
        if rebuilt:
            st.success(f"Reconstrucción completa: {rebuilt['num_parts']} partes.")
            st.download_button(
                "⬇️ Descargar Excel reconstruido", data=rebuilt["bytes"],
                file_name=rebuilt["original_name"],
                mime="application/octet-stream", use_container_width=True,
            )
            st.caption(f"SHA-256: {rebuilt['sha256']}")

    with tab_guide:
        st.subheader("Publicar los archivos en GitHub")
        st.markdown(
            """1. Abre **Dividir**, selecciona el Excel, la categoría y el año.
2. Pulsa **Preparar partes para GitHub** y descarga el ZIP.
3. En tu iPad o computador, **descomprime** el ZIP.
4. En [tu repositorio de GitHub](https://github.com/antoniorra1983-ops/Precio_estabilizado),
   abre la ruta `Datos/<categoría>/<año>/`.
5. Selecciona **Add file → Upload files**, sube **todas** las partes y confirma con **Commit changes**.
6. Abre tu calculadora CEN; su cargador de archivos divididos debe reconstruir el libro.

**No subas:** el ZIP, `INSTRUCCIONES_NO_SUBIR_A_DATOS.txt`, partes incompletas ni una segunda copia del Excel entero.
"""
        )
        st.info("El archivo se procesa en Streamlit Cloud y se descarga al navegador. No se almacena automáticamente en GitHub.")
        st.warning("Antes de publicar balances comerciales, confirma que puedes compartirlos: los archivos de un repositorio público quedan accesibles a terceros.")
        st.caption("Formatos reconocidos por la calculadora CEN: .xlsx, .xlsm y .csv; partes .part001, .part002…")


if __name__ == "__main__":
    main()
