import docx
import os
import json
import tempfile
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn

# --- COLORES EXACTOS DE TU IMAGEN Y DISEÑO ---
COLOR_TEAL_OSCURO = "005B96"  # Azul vibrante más moderno y llamativo
COLOR_GRIS_CLARO = "E8EEF2"   # Azul/Gris platino más limpio y luminoso
COLOR_BLANCO = "FFFFFF"       # Texto superior / Fondo principal derecho
COLOR_TEXTO_OSCURO = "222222" # Texto de lectura general
COLOR_MUTED = "555555"        # Fechas y subtítulos

# --- FUNCIONES DE INYECCIÓN XML NATIVA ---
def set_cell_background(cell, color_hex):
    """Aplica color de fondo sólido a una celda sin intermediarios."""
    tcPr = cell._tc.get_or_add_tcPr()
    for existing_shd in tcPr.xpath('w:shd'):
        tcPr.remove(existing_shd)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def set_cell_margins(cell, top=200, left=200, bottom=200, right=200):
    """Establece márgenes internos (padding) por celda para evitar textos pegados."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    margins = {'top': top, 'left': left, 'bottom': bottom, 'right': right}
    for m_type, m_val in margins.items():
        node = OxmlElement(f'w:{m_type}')
        node.set(qn('w:w'), str(m_val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_bottom_border(p, color_hex, size=12):
    """Inserta una línea divisoria horizontal debajo de un título."""
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(size))
    bottom.set(qn('w:space'), '4')
    bottom.set(qn('w:color'), color_hex)
    pBdr.append(bottom)
    pPr.append(pBdr)

def add_run(p, text, size=10, bold=False, color_hex=COLOR_TEXTO_OSCURO, italic=False):
    """Escribe texto formateado garantizando que Word aplique el color HEX."""
    run = p.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    r, g, b = tuple(int(color_hex[i:i+2], 16) for i in (0, 2, 4))
    run.font.color.rgb = RGBColor(r, g, b)
    return run

def allow_row_to_break(row):
    """Permite que una fila de tabla se divida limpiamente entre páginas."""
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    cantSplit = trPr.find(qn('w:cantSplit'))
    if cantSplit is not None:
        cantSplit.set(qn('w:val'), '0')

def add_hyperlink(p, url, text, size=9, color_hex="0000EE"):
    """
    Añade un hipervínculo a un párrafo.
    """
    part = p.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    run_element = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    
    # Estilo de hipervínculo (subrayado)
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)
    run_element.append(rPr)
    run_element.text = text
    hyperlink.append(run_element)
    p._p.append(hyperlink)
    
    # Se aplica el formato de fuente después de crear el elemento
    # El estilo 'Hyperlink' de Word se encarga del color, pero lo forzamos por si acaso
    run = p.runs[-1]
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    r, g, b = tuple(int(color_hex[i:i+2], 16) for i in (0, 2, 4))
    run.font.color.rgb = RGBColor(r, g, b)
    return run

def load_data(json_path):
    """Carga los datos del currículum desde un archivo JSON."""
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def crear_cv_completo():
    doc = docx.Document() 

    # Configuración A4 con un ligero margen superior/inferior para evitar que el texto
    # se superponga (traslape) con los bordes físicos al cambiar de página.
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.3)
        section.bottom_margin = Inches(0.3)
        section.left_margin = Inches(0)
        section.right_margin = Inches(0)

    # Función local para forzar a la tabla a ocupar el 100% del ancho de la hoja
    def set_table_full_width(tbl):
        tblPr = tbl._tbl.tblPr
        tblW = tblPr.find(qn('w:tblW'))
        if tblW is None:
            tblW = OxmlElement('w:tblW')
            tblPr.append(tblW)
        tblW.set(qn('w:type'), 'pct')
        tblW.set(qn('w:w'), '5000')

    # Función local para remover bordes de una tabla
    def remove_borders(tbl):
        tblBorders = OxmlElement('w:tblBorders')
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            bdr = OxmlElement(f'w:{border_name}')
            bdr.set(qn('w:val'), 'none')
            tblBorders.append(bdr)
        tbl._tbl.tblPr.append(tblBorders)

    # Matriz estructural principal de 1 fila x 2 columnas (evita fallos de paginación)
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    remove_borders(table)
    set_table_full_width(table)

    # Distribución simétrica de dimensiones
    table.rows[0].cells[0].width = Inches(2.75)
    table.rows[0].cells[1].width = Inches(5.52)

    # Permitir que la fila principal se divida entre páginas para evitar cortes
    allow_row_to_break(table.rows[0])

    cell_left = table.cell(0, 0)
    cell_right = table.cell(0, 1)

    # Configuración de la columna izquierda (balanceamos márgenes para dar espacio a la foto y evitar cortes)
    set_cell_background(cell_left, COLOR_GRIS_CLARO)
    set_cell_margins(cell_left, top=300, left=500, bottom=300, right=250)

    # --- CARGA DE DATOS ---
    data = load_data('cv_data.json')

    # =========================================================================
    # 1. BLOQUE SUPERIOR IZQUIERDO: FOTO NATIVA Y CONTACTO
    # =========================================================================
    
    p_foto = cell_left.paragraphs[0]
    p_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_foto.paragraph_format.space_before = Pt(10)
    
    # Detectamos la imagen proporcionada (img.jpg), si no existe, usamos fallback
    base_dir = os.path.dirname(os.path.abspath(__file__))
    img_path = os.path.join(base_dir, "img.jpg")
    usar_temp = not os.path.exists(img_path)
    
    if usar_temp:
        dummy_gif = b'GIF89a\x01\x00\x01\x00\x80\x00\x00\xcc\xcc\xcc\x00\x00\x00!\xf9\x04\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;'
        img_path = os.path.join(tempfile.gettempdir(), "placeholder_cv.gif")
        with open(img_path, "wb") as f:
            f.write(dummy_gif)

    run_foto = p_foto.add_run()
    run_foto.add_picture(img_path, width=Inches(1.7), height=Inches(1.7))
    
    if usar_temp:
        p_ayuda = cell_left.add_paragraph()
        p_ayuda.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_ayuda.paragraph_format.space_before = Pt(5)
        add_run(p_ayuda, "(Clic derecho en el cuadro gris -> Cambiar imagen)", size=8.5, italic=True, color_hex=COLOR_MUTED)
        try: os.remove(img_path)
        except: pass

    # --- Sub-sección: Contacto (Inmediatamente debajo de la foto) ---
    p_t_contacto = cell_left.add_paragraph()
    p_t_contacto.paragraph_format.space_before = Pt(25)
    add_run(p_t_contacto, data['contacto']['titulo'], size=13, bold=True)
    add_bottom_border(p_t_contacto, COLOR_TEXTO_OSCURO, size=8)

    # Renderizar datos de contacto, con manejo especial para el hyperlink
    for key, value in data['contacto']['datos'].items():
        p_d = cell_left.add_paragraph()
        p_d.paragraph_format.space_before = Pt(4)
        p_d.paragraph_format.space_after = Pt(0)
        p_d.paragraph_format.line_spacing = 1.0
        
        if key == 'linkedin':
            add_run(p_d, "🔗 ", size=9)
            add_hyperlink(p_d, value['url'], value['texto'], size=9)
        else:
            add_run(p_d, value, size=9)

    # =========================================================================
    # 2. BLOQUE SUPERIOR DERECHO: NOMBRE Y RESUMEN (Fondo Teal - Letras Blancas)
    # =========================================================================
    # Se crea una tabla para esta sección para controlar el fondo y los márgenes
    table_top_right = cell_right.add_table(rows=1, cols=1)
    table_top_right.autofit = False
    table_top_right.rows[0].cells[0].width = Inches(5.52)
    remove_borders(table_top_right)
    set_table_full_width(table_top_right)
    allow_row_to_break(table_top_right.rows[0])

    cell_top_right = table_top_right.cell(0, 0)
    set_cell_background(cell_top_right, COLOR_TEAL_OSCURO)
    set_cell_margins(cell_top_right, top=300, left=400, bottom=400, right=400)

    p_nombre = cell_top_right.paragraphs[0]
    add_run(p_nombre, data['nombre_completo'], size=26, bold=True, color_hex=COLOR_BLANCO)
    add_bottom_border(p_nombre, COLOR_BLANCO, size=18)

    for i, parrafo in enumerate(data['resumen_profesional']):
        p_resumen = cell_top_right.add_paragraph()
        p_resumen.paragraph_format.space_before = Pt(14) if i == 0 else Pt(8)
        p_resumen.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_resumen.paragraph_format.line_spacing = 1.15
        add_run(p_resumen, parrafo, size=9.5, color_hex=COLOR_BLANCO)

    p_fortalezas_t = cell_top_right.add_paragraph()
    p_fortalezas_t.paragraph_format.space_before = Pt(12)
    add_run(p_fortalezas_t, "FORTALEZAS PRINCIPALES", size=11, bold=True, color_hex=COLOR_BLANCO)

    for fort in data['fortalezas']:
        p_f = cell_top_right.add_paragraph()
        p_f.paragraph_format.space_after = Pt(2)
        add_run(p_f, "▪ ", size=9, color_hex=COLOR_BLANCO)
        add_run(p_f, fort, size=9.5, color_hex=COLOR_BLANCO)

    # =========================================================================
    # 3. BLOQUE IZQUIERDO: COMPETENCIAS, IDIOMAS Y OTROS ESTUDIOS
    # =========================================================================
    
    # --- Sub-sección: Tecnologías ---
    if 'tecnologias' in data:
        p_t_tech = cell_left.add_paragraph()
        p_t_tech.paragraph_format.space_before = Pt(25)
        add_run(p_t_tech, data['tecnologias']['titulo'], size=13, bold=True)
        add_bottom_border(p_t_tech, COLOR_TEXTO_OSCURO, size=8)

        for tech in data['tecnologias']['lista']:
            p_tc = cell_left.add_paragraph()
            p_tc.paragraph_format.space_after = Pt(0)
            p_tc.paragraph_format.line_spacing = 1.0
            p_tc.paragraph_format.left_indent = Inches(0.15)
            add_run(p_tc, "▪ ", size=8.5)
            add_run(p_tc, tech, size=9)
    
    # --- Sub-sección: Competencias (31 competencias completas extraídas del HTML) ---
    p_t_comp = cell_left.add_paragraph()
    p_t_comp.paragraph_format.space_before = Pt(25)
    add_run(p_t_comp, data['competencias']['titulo'], size=13, bold=True)
    add_bottom_border(p_t_comp, COLOR_TEXTO_OSCURO, size=8)

    for comp in data['competencias']['lista']:
        p_c = cell_left.add_paragraph()
        p_c.paragraph_format.space_after = Pt(0)
        p_c.paragraph_format.line_spacing = 1.0
        p_c.paragraph_format.left_indent = Inches(0.15)
        add_run(p_c, "▪ ", size=8.5)
        add_run(p_c, comp, size=9)

    # --- Sub-sección: Idiomas ---
    p_t_idiomas = cell_left.add_paragraph()
    p_t_idiomas.paragraph_format.space_before = Pt(20)
    add_run(p_t_idiomas, data['idiomas']['titulo'], size=13, bold=True)
    add_bottom_border(p_t_idiomas, COLOR_TEXTO_OSCURO, size=8)

    for item in data['idiomas']['datos']:
        p_l = cell_left.add_paragraph()
        p_l.paragraph_format.space_after = Pt(2)
        p_l.paragraph_format.left_indent = Inches(0.15)
        add_run(p_l, f"{item['idioma']}: ", size=9.5, bold=True)
        add_run(p_l, item['nivel'], size=9)

    # --- Sub-sección: Otros Estudios / Certificaciones ---
    p_t_cert = cell_left.add_paragraph()
    p_t_cert.paragraph_format.space_before = Pt(20)
    add_run(p_t_cert, data['otros_estudios']['titulo'], size=13, bold=True)
    add_bottom_border(p_t_cert, COLOR_TEXTO_OSCURO, size=8)

    for cert in data['otros_estudios']['lista']:
        p_ce = cell_left.add_paragraph()
        p_ce.paragraph_format.space_after = Pt(2)
        p_ce.paragraph_format.left_indent = Inches(0.15)
        add_run(p_ce, "▪ ", size=8.5)
        add_run(p_ce, cert, size=9)

    # =========================================================================
    # 4. BLOQUE INFERIOR DERECHO: FORMACIÓN Y EXPERIENCIA LABORAL (Fondo Blanco)
    # =========================================================================
    # Se crea una tabla para el contenido principal para controlar el fondo y los márgenes
    table_bottom_right = cell_right.add_table(rows=1, cols=1)
    table_bottom_right.autofit = False
    table_bottom_right.rows[0].cells[0].width = Inches(5.52)
    remove_borders(table_bottom_right)
    set_table_full_width(table_bottom_right)
    allow_row_to_break(table_bottom_right.rows[0])

    cell_bottom_right = table_bottom_right.cell(0, 0)
    set_cell_background(cell_bottom_right, COLOR_BLANCO)
    set_cell_margins(cell_bottom_right, top=400, left=400, bottom=600, right=400)

    # --- Sección: Formación ---
    p_t_form = cell_bottom_right.paragraphs[0]
    p_t_form.paragraph_format.keep_with_next = True
    add_run(p_t_form, data['formacion']['titulo'], size=15, bold=True, color_hex=COLOR_TEAL_OSCURO)
    add_bottom_border(p_t_form, COLOR_TEAL_OSCURO, size=10)
    p_t_form.paragraph_format.space_after = Pt(0)

    for i, formacion in enumerate(data['formacion']['datos']):
        p_edu = cell_bottom_right.add_paragraph()
        p_edu.paragraph_format.keep_with_next = True
        p_edu.paragraph_format.space_before = Pt(8) if i == 0 else Pt(10)
        p_edu.paragraph_format.space_after = Pt(2)
        add_run(p_edu, formacion['titulo'], size=10.5, bold=True)
        # Solo añadir subtítulo si es significativamente diferente y no es redundante
        if formacion.get('subtitulo') and formacion['subtitulo'].lower() not in formacion['titulo'].lower():
            add_run(p_edu, f" ({formacion['subtitulo']})", size=10, italic=True)
        p_edu.add_run('\n') # Salto de línea para la institución
        add_run(p_edu, formacion['institucion'], size=10, bold=True, color_hex=COLOR_MUTED)
        if formacion.get('ubicacion_periodo'):
            add_run(p_edu, f" - {formacion['ubicacion_periodo']}", size=10, italic=True, color_hex=COLOR_MUTED)
        if formacion.get('nota'):
            p_nota = cell_bottom_right.add_paragraph()
            p_nota.paragraph_format.space_after = Pt(2)
            p_nota.paragraph_format.line_spacing = 1.0
            add_run(p_nota, formacion['nota'], size=9.5, italic=True, color_hex=COLOR_TEAL_OSCURO)

    # --- Sección: Experiencia ---
    p_t_expr = cell_bottom_right.add_paragraph()
    p_t_expr.paragraph_format.keep_with_next = True
    p_t_expr.paragraph_format.space_before = Pt(25)
    add_run(p_t_expr, data['experiencia']['titulo'], size=15, bold=True, color_hex=COLOR_TEAL_OSCURO)
    add_bottom_border(p_t_expr, COLOR_TEAL_OSCURO, size=10)

    for i, exp in enumerate(data['experiencia']['datos']):
        # Nombre de la empresa como título principal del bloque de experiencia
        p_company = cell_bottom_right.add_paragraph()
        p_company.paragraph_format.keep_with_next = True
        p_company.paragraph_format.space_before = Pt(10) if i == 0 else Pt(14)
        p_company.paragraph_format.space_after = Pt(2)
        add_run(p_company, exp['empresa'], size=11, bold=True, color_hex=COLOR_TEAL_OSCURO)

        # Iterar sobre cada rol dentro de la empresa
        for role in exp['roles']:
            p_role = cell_bottom_right.add_paragraph()
            p_role.paragraph_format.keep_with_next = True
            p_role.paragraph_format.space_before = Pt(4)
            add_run(p_role, f"{role['puesto']}", size=10.5, bold=True)
            add_run(p_role, f"  |  {role['periodo']}", size=9.5, italic=True, color_hex=COLOR_MUTED)

            for bullet in role['descripcion']:
                p_b = cell_bottom_right.add_paragraph(style='List Bullet')
                p_b.paragraph_format.space_after = Pt(2)
                p_b.paragraph_format.keep_lines_together = True # Evita que un solo bullet se parta
                p_b.paragraph_format.left_indent = Inches(0.25)
                add_run(p_b, bullet, size=9.5)
            
            if role.get('tecnologias'):
                p_tech_role = cell_bottom_right.add_paragraph()
                p_tech_role.paragraph_format.space_before = Pt(4)
                add_run(p_tech_role, "Tecnologías: ", size=9.5, bold=True, color_hex=COLOR_MUTED)
                add_run(p_tech_role, role['tecnologias'], size=9.5, italic=True, color_hex=COLOR_MUTED)

    # --- Sección: Proyectos Destacados ---
    if 'proyectos_destacados' in data and data['proyectos_destacados']['datos']:
        p_t_proy = cell_bottom_right.add_paragraph()
        p_t_proy.paragraph_format.keep_with_next = True
        p_t_proy.paragraph_format.space_before = Pt(25)
        add_run(p_t_proy, data['proyectos_destacados']['titulo'], size=15, bold=True, color_hex=COLOR_TEAL_OSCURO)
        add_bottom_border(p_t_proy, COLOR_TEAL_OSCURO, size=10)

        for i, proyecto in enumerate(data['proyectos_destacados']['datos']):
            p_proy_nombre = cell_bottom_right.add_paragraph()
            p_proy_nombre.paragraph_format.keep_with_next = True
            p_proy_nombre.paragraph_format.space_before = Pt(10) if i == 0 else Pt(14)
            p_proy_nombre.paragraph_format.space_after = Pt(2)
            add_run(p_proy_nombre, proyecto['nombre'], size=11, bold=True, color_hex=COLOR_TEAL_OSCURO)

            p_proy_desc = cell_bottom_right.add_paragraph()
            p_proy_desc.paragraph_format.keep_lines_together = True
            p_proy_desc.paragraph_format.space_after = Pt(2)
            add_run(p_proy_desc, proyecto['descripcion'], size=9.5)

            p_proy_tech = cell_bottom_right.add_paragraph()
            add_run(p_proy_tech, "Tecnologías: ", size=9.5, bold=True, color_hex=COLOR_MUTED)
            add_run(p_proy_tech, proyecto['tecnologias'], size=9.5, italic=True, color_hex=COLOR_MUTED)

    # --- Sección: Referencias Laborales ---
    if 'referencias_laborales' in data and data['referencias_laborales']['datos']:
        p_t_ref_l = cell_bottom_right.add_paragraph()
        p_t_ref_l.paragraph_format.keep_with_next = True
        p_t_ref_l.paragraph_format.space_before = Pt(25)
        add_run(p_t_ref_l, data['referencias_laborales']['titulo'], size=15, bold=True, color_hex=COLOR_TEAL_OSCURO)
        add_bottom_border(p_t_ref_l, COLOR_TEAL_OSCURO, size=10)

        for i, ref in enumerate(data['referencias_laborales']['datos']):
            p_ref = cell_bottom_right.add_paragraph()
            p_ref.paragraph_format.keep_lines_together = True
            p_ref.paragraph_format.space_before = Pt(8) if i == 0 else Pt(12)
            add_run(p_ref, f"👤 {ref['nombre']}\n", size=10.5, bold=True)
            empresa_str = f" | {ref['empresa']}" if ref.get('empresa') else ""
            add_run(p_ref, f"💼 {ref['puesto']}{empresa_str}\n", size=10, color_hex=COLOR_MUTED)
            add_run(p_ref, f"📱 Tel: {ref['telefono']}", size=10, color_hex=COLOR_MUTED)
            if ref.get('correo'):
                add_run(p_ref, f"  |  ✉️ {ref['correo']}", size=10, color_hex=COLOR_MUTED)

    # --- Sección: Referencias Personales ---
    if 'referencias_personales' in data and data['referencias_personales']['datos']:
        p_t_ref_p = cell_bottom_right.add_paragraph()
        p_t_ref_p.paragraph_format.keep_with_next = True
        p_t_ref_p.paragraph_format.space_before = Pt(25)
        add_run(p_t_ref_p, data['referencias_personales']['titulo'], size=15, bold=True, color_hex=COLOR_TEAL_OSCURO)
        add_bottom_border(p_t_ref_p, COLOR_TEAL_OSCURO, size=10)

        for i, ref in enumerate(data['referencias_personales']['datos']):
            p_ref = cell_bottom_right.add_paragraph()
            p_ref.paragraph_format.keep_lines_together = True
            p_ref.paragraph_format.space_before = Pt(8) if i == 0 else Pt(12)
            add_run(p_ref, f"👤 {ref['nombre']}\n", size=10.5, bold=True)
            add_run(p_ref, f"💼 {ref['puesto']}\n", size=10, color_hex=COLOR_MUTED)
            add_run(p_ref, f"📱 Tel: {ref['telefono']}", size=10, color_hex=COLOR_MUTED)
            if ref.get('correo'):
                add_run(p_ref, f"  |  ✉️ Correo: {ref['correo']}", size=10, color_hex=COLOR_MUTED)

    # Guardado del documento
    nombre_salida = "CV_Eliseo_Lopez.docx"
    doc.save(nombre_salida)
    print(f"¡Éxito! Se ha generado el archivo completo: '{nombre_salida}'")

if __name__ == "__main__":
    crear_cv_completo()