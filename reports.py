try:
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas
    from reportlab.lib import colors
    from reportlab.graphics.shapes import Drawing
    from reportlab.graphics.charts.piecharts import Pie
    from reportlab.graphics.charts.barcharts import VerticalBarChart
    from reportlab.graphics.charts.legends import Legend
    from reportlab.graphics import renderPDF
    LIBRERIA_PDF_DISPONIBLE = True
except ImportError:
    LIBRERIA_PDF_DISPONIBLE = False

import datetime
import app_config as cfg

def generar_reporte_pdf(ruta_pdf, df_base, lead_time_func, tipo_reporte, data_pack_main, data_pack_antes=None, col_inicio=0):
    if not LIBRERIA_PDF_DISPONIBLE: return
    
    c = canvas.Canvas(ruta_pdf, pagesize=letter)
    width, height = letter
    
    # -------------------------------------------------------------------------
    # FUNCIÓN INTERNA PARA DIBUJAR UNA SECCIÓN (Global, País o Categoría)
    # -------------------------------------------------------------------------
    def dibujar_seccion(data_main, data_antes, titulo_seccion, subtitulo_seccion=""):
        status_data = data_main.get('status_data', [])
        heatmap_data = data_main.get('heatmap', {})
        resumen_cambios = data_main.get('resumen_cambios', {})
        total_skus = data_main.get('count', 0)
        
        status_data_antes = data_antes.get('status_data', []) if data_antes else []
        # --- CORRECCIÓN AQUÍ: Definimos la variable faltante ---
        heatmap_antes = data_antes.get('heatmap', {}) if data_antes else {} 
        
        # Categorías
        rows_in_scope = [item['row'] for item in status_data]
        cats_encontradas = set()
        for r in rows_in_scope:
            prod_val = str(df_base.iloc[r-1, 2])
            prefix = prod_val[:2].upper()
            if prefix in cfg.CATEGORIAS_MAP: cats_encontradas.add(cfg.CATEGORIAS_MAP[prefix])
        txt_categorias = " | ".join(sorted(cats_encontradas)) if cats_encontradas else "General / Otros"

        # Conteos
        oos_count = sum(1 for x in status_data if x['status'] == 'OOS')
        ustn_count = sum(1 for x in status_data if x['status'] == 'USTN')
        ostn_h_count = sum(1 for x in status_data if x['status'] == 'OSTN High')
        ostn_m_count = sum(1 for x in status_data if x['status'] == 'OSTN Med')
        
        # Calculo de OK puro (Residual)
        ok_puro = total_skus - oos_count - ustn_count - ostn_h_count - ostn_m_count
        
        # Para el Pie Chart visual: OK + OSTN se consideran "Abastecidos" (Ocultamos azules)
        ok_visual = ok_puro + ostn_h_count + ostn_m_count
        
        oos_antes = sum(1 for x in status_data_antes if x['status'] == 'OOS') if status_data_antes else 0
        ustn_antes = sum(1 for x in status_data_antes if x['status'] == 'USTN') if status_data_antes else 0

        # Listas Separadas (Riesgos vs Excesos)
        lista_riesgos = []
        lista_excesos = []
        
        for item in status_data:
            st = item['status']
            prod_name = str(df_base.iloc[item['row']-1, 2])
            gap_val = item.get('gap', 0)
            avg_a = item.get('avg_dos_act', 0)
            avg_t = item.get('avg_dos_tgt', 0)
            rec_week = item.get('recovery', '-')
            
            entry = {
                'prod': prod_name, 'status': st, 'gap': gap_val, 
                'avg_a': avg_a, 'avg_t': avg_t, 'rec': rec_week
            }
            
            if st in ['OOS', 'USTN']:
                entry['prio'] = 1 if st == 'OOS' else 2
                lista_riesgos.append(entry)
            elif st in ['OSTN High', 'OSTN Med']:
                entry['prio'] = 1 if st == 'OSTN High' else 2
                lista_excesos.append(entry)

        # Ordenar listas
        lista_riesgos.sort(key=lambda x: (x['prio'], -x['gap'])) # Mayor faltante primero
        lista_excesos.sort(key=lambda x: (x['prio'], x['gap']))  # Mayor sobrante (gap negativo) primero
        
        top_riesgos = lista_riesgos[:5]
        top_excesos = lista_excesos[:5]

        # --- DRAWING HELPERS ---
        def dibujar_encabezado_local():
            c.setFillColorRGB(0.0, 0.15, 0.45) 
            c.rect(0, height - 100, width, 100, fill=True, stroke=False)
            c.setFillColorRGB(1, 1, 1) 
            titulo_main = "Reporte de Escenario Ideal" if tipo_reporte == "IDEAL" else "Reporte de Situación Actual"
            c.setFont("Helvetica-Bold", 18); c.drawString(30, height - 35, titulo_main)
            c.setFont("Helvetica-Bold", 14); c.drawString(30, height - 55, f"{titulo_seccion}")
            if subtitulo_seccion: c.setFont("Helvetica-Oblique", 12); c.drawString(300, height - 55, f"{subtitulo_seccion}")
            c.setFont("Helvetica", 11); c.drawString(30, height - 72, "Market Planning - Philip Morris International")
            c.setFont("Helvetica-Oblique", 9); c.setFillColorRGB(0.8, 0.8, 0.8)
            c.drawString(30, height - 88, f"Categorías: {txt_categorias}")
            c.setFillColorRGB(1, 1, 1); c.setFont("Helvetica", 9)
            c.drawRightString(width - 30, height - 35, f"Fecha: {datetime.datetime.now().strftime('%Y-%m-%d')}")

        def dibujar_separador_local(y_pos):
            c.setStrokeColorRGB(0.85, 0.85, 0.85); c.setLineWidth(1)
            c.line(30, y_pos, width - 30, y_pos)
            return y_pos - 30

        def label_formatter(val): return str(int(val)) if val > 0 else ""

        # Función Graficadora Maestra
        def dibujar_grafico_evolucion(heatmap_source, titulo_grafico, pos_y, modo="RIESGO", altura_grafico=125):
            c.setFillColorRGB(0, 0, 0); c.setFont("Helvetica-Bold", 10)
            c.drawString(30, pos_y, titulo_grafico)
            
            start_col = col_inicio + 1
            weeks_to_show = cfg.SEMANAS_A_EVALUAR
            
            data_serie_1 = []; data_serie_2 = []; labels = []
            
            # Configuración según modo
            if modo == "RIESGO":
                colores = [colors.HexColor("#" + cfg.COLOR_PDF_ROJO), colors.HexColor("#" + cfg.COLOR_PDF_AMARILLO)]
                keys = ["OOS", "USTN"]
            else: # EXCESO
                colores = [colors.HexColor("#" + cfg.COLOR_PDF_AZUL_OSCURO), colors.HexColor("#" + cfg.COLOR_PDF_AZUL_CLARO)]
                keys = ["OSTN High", "OSTN Med"]

            for i in range(weeks_to_show):
                c_idx = start_col + i
                if c_idx >= df_base.shape[1]: break
                lbl = str(df_base.iloc[1, c_idx])
                if lbl == 'nan' or lbl == '': lbl = str(df_base.iloc[0, c_idx]).split(" ")[0]
                labels.append(lbl)
                
                items_semana = heatmap_source.get(c_idx, [])
                c1 = 0; c2 = 0
                for item in items_semana:
                    st = item[1]
                    if st == keys[0]: c1 += 1
                    elif st == keys[1]: c2 += 1
                data_serie_1.append(c1); data_serie_2.append(c2)
                
            bc = VerticalBarChart()
            bc.x = 50; bc.y = 50; bc.height = altura_grafico; bc.width = 450
            bc.data = [data_serie_1, data_serie_2]
            bc.categoryAxis.style = 'stacked'
            bc.categoryAxis.labels.boxAnchor = 'ne'; bc.categoryAxis.labels.dx = 8; bc.categoryAxis.labels.dy = -2
            bc.categoryAxis.labels.angle = 30; bc.categoryAxis.categoryNames = labels
            
            max_val = 0
            for i in range(len(data_serie_1)):
                if data_serie_1[i] + data_serie_2[i] > max_val: max_val = data_serie_1[i] + data_serie_2[i]
            
            bc.valueAxis.valueMin = 0; bc.valueAxis.valueMax = max_val + (2 if max_val < 10 else 5)
            bc.valueAxis.valueStep = 1 if max_val < 10 else None
            
            bc.bars[0].fillColor = colores[0]
            bc.bars[1].fillColor = colores[1]
            bc.barLabels.nudge = 5; bc.barLabelFormat = label_formatter; bc.barLabels.fontSize = 7
            
            d = Drawing(500, altura_grafico + 50); d.add(bc)
            renderPDF.draw(d, c, 30, pos_y - (altura_grafico + 40))
            
            return pos_y - (altura_grafico + 60)

        dibujar_encabezado_local()
        y = height - 130
        
        # 1. RESUMEN GLOBAL (SOLO OK vs RIESGOS)
        c.setFillColorRGB(0, 0, 0); c.setFont("Helvetica-Bold", 14)
        c.drawString(30, y, "1. Resumen Ejecutivo (Foco Riesgos)")
        y -= 30
        
        col1_x = 30; col2_x = 230; col3_x = 420
        y_text = y
        c.setFont("Helvetica", 10)
        c.drawString(col1_x, y_text, f"• Total SKUs: {total_skus}")
        y_text -= 18
        c.drawString(col1_x, y_text, f"• Críticos (OOS/USTN): {oos_count + ustn_count}")
        y_text -= 18
        c.drawString(col1_x, y_text, f"• Exceso (OSTN): {ostn_h_count + ostn_m_count}")
        
        d = Drawing(150, 100); pc = Pie()
        pc.x = 25; pc.y = 0; pc.width = 100; pc.height = 100
        # Data: OK(incluye ostn), USTN, OOS
        pc.data = [ok_visual, ustn_count, oos_count]; pc.labels = None 
        
        color_ok = colors.HexColor("#" + cfg.COLOR_PDF_VERDE)
        color_ustn = colors.HexColor("#" + cfg.COLOR_PDF_AMARILLO)
        color_oos = colors.HexColor("#" + cfg.COLOR_PDF_ROJO)
        
        pc.slices[0].fillColor = color_ok; pc.slices[1].fillColor = color_ustn; pc.slices[2].fillColor = color_oos
        d.add(pc)
        renderPDF.draw(d, c, col2_x, y - 60)
        
        def draw_legend_item(canvas_obj, x, y, color, label, count, total):
            pct = (count / total * 100) if total > 0 else 0
            canvas_obj.setFillColor(color)
            canvas_obj.rect(x, y, 10, 10, fill=True, stroke=False)
            canvas_obj.setFillColorRGB(0,0,0)
            canvas_obj.setFont("Helvetica-Bold", 9)
            canvas_obj.drawString(x + 15, y + 2, label)
            canvas_obj.setFont("Helvetica", 9)
            canvas_obj.drawRightString(x + 120, y + 2, f"{count}")
            canvas_obj.drawString(x + 125, y + 2, f"({pct:.0f}%)")

        y_leg = y + 10
        draw_legend_item(c, col3_x, y_leg, color_ok, "Abasto OK", ok_visual, total_skus); y_leg -= 25
        draw_legend_item(c, col3_x, y_leg, color_ustn, "Riesgo (USTN)", ustn_count, total_skus); y_leg -= 25
        draw_legend_item(c, col3_x, y_leg, color_oos, "Crítico (OOS)", oos_count, total_skus); y_leg -= 25
        
        y -= 100
        y = dibujar_separador_local(y)

        # 2. IMPACTO
        if tipo_reporte == "IDEAL" and status_data_antes:
            c.setFillColorRGB(0, 0, 0); c.setFont("Helvetica-Bold", 14)
            c.drawString(30, y, "2. Impacto de la Optimización")
            y -= 25
            c.setFillColorRGB(0.95, 0.95, 0.95); c.rect(30, y-5, 350, 15, fill=True, stroke=False)
            c.setFillColorRGB(0, 0, 0); c.setFont("Helvetica-Bold", 9)
            c.drawString(40, y, "Métrica"); c.drawString(160, y, "Antes"); c.drawString(240, y, "Después"); c.drawString(320, y, "Variación")
            y -= 20
            c.setFont("Helvetica", 10); c.drawString(40, y, "Productos OOS")
            c.drawString(160, y, str(oos_antes)); c.drawString(240, y, str(oos_count))
            dif_oos = oos_count - oos_antes
            txt_oos = f"{dif_oos}" if dif_oos <= 0 else f"+{dif_oos}"
            c.setFillColorRGB(0, 0.6, 0) if dif_oos < 0 else c.setFillColorRGB(0.8, 0, 0)
            c.setFont("Helvetica-Bold", 10); c.drawString(320, y, txt_oos)
            y -= 18
            c.setFillColorRGB(0, 0, 0); c.setFont("Helvetica", 10)
            c.drawString(40, y, "Productos USTN")
            c.drawString(160, y, str(ustn_antes)); c.drawString(240, y, str(ustn_count))
            dif_ustn = ustn_count - ustn_antes
            txt_ustn = f"{dif_ustn}" if dif_ustn <= 0 else f"+{dif_ustn}"
            c.setFillColorRGB(0, 0.6, 0) if dif_ustn < 0 else c.setFillColorRGB(0.8, 0.4, 0)
            c.setFont("Helvetica-Bold", 10); c.drawString(320, y, txt_ustn)
            y -= 20
            y = dibujar_separador_local(y)

        # 3. HEATMAP DE RIESGOS (ROJO/AMARILLO)
        if heatmap_data:
            if y < 150: c.showPage(); dibujar_encabezado_local(); y = height - 130
            c.setFillColorRGB(0, 0, 0); c.setFont("Helvetica-Bold", 14)
            c.drawString(30, y, f"3. Tendencia de Riesgos (OOS / USTN)")
            y -= 35
            
            # Leyenda
            c.setFont("Helvetica", 8)
            c.setFillColorRGB(0.9, 0.4, 0.4); c.rect(400, y+10, 8, 8, fill=True, stroke=False)
            c.setFillColorRGB(0, 0, 0); c.drawString(412, y+10, "OOS")
            c.setFillColorRGB(0.9, 0.8, 0.4); c.rect(440, y+10, 8, 8, fill=True, stroke=False)
            c.setFillColorRGB(0, 0, 0); c.drawString(452, y+10, "USTN")
            
            # Aquí era donde fallaba antes: heatmap_antes ahora sí existe
            if tipo_reporte == "IDEAL" and heatmap_antes:
                y = dibujar_grafico_evolucion(heatmap_antes, "A. Situación Original", y, "RIESGO", 90)
                y -= 20 
                y = dibujar_grafico_evolucion(heatmap_data, "B. Escenario Ideal", y, "RIESGO", 90)
            else:
                y = dibujar_grafico_evolucion(heatmap_data, "Distribución de faltantes por semana", y, "RIESGO", 110)
            y = dibujar_separador_local(y)

        # 4. TABLA TOP 5 RIESGOS
        if y < 150: c.showPage(); dibujar_encabezado_local(); y = height - 130
        c.setFillColorRGB(0, 0, 0); c.setFont("Helvetica-Bold", 14)
        c.drawString(30, y, "4. Top 5 Riesgos (Prioridad Faltantes)")
        y -= 25

        if top_riesgos:
            c.setFillColorRGB(0.95, 0.95, 0.95); c.rect(30, y-5, 530, 15, fill=True, stroke=False)
            c.setFillColorRGB(0, 0, 0); c.setFont("Helvetica-Bold", 9)
            c.drawString(40, y, "Producto"); c.drawString(250, y, "GAP Inv.")
            c.drawString(330, y, "Avg DoS"); c.drawString(420, y, "Recuperación")
            c.drawString(500, y, "Estatus")
            y -= 20
            for prod in top_riesgos:
                c.setFont("Helvetica", 10)
                nombre = prod['prod'][:35]
                c.drawString(40, y, f"• {nombre}")
                gap_str = f"{prod['gap']:,.0f}"; c.setFillColorRGB(0.8, 0, 0); c.drawRightString(290, y, gap_str)
                c.setFillColorRGB(0, 0, 0); dos_str = f"{prod['avg_a']:.1f} vs {prod['avg_t']:.1f}"; c.drawRightString(400, y, dos_str)
                c.drawString(430, y, str(prod['rec']))
                st = prod['status']
                color_st = colors.HexColor("#" + cfg.COLOR_PDF_ROJO) if st == "OOS" else colors.HexColor("#" + cfg.COLOR_PDF_AMARILLO)
                c.setFillColor(color_st); c.setFont("Helvetica-Bold", 10); c.drawString(500, y, st); c.setFillColorRGB(0, 0, 0)
                y -= 18; c.setStrokeColorRGB(0.9, 0.9, 0.9); c.setLineWidth(0.5); c.line(30, y+14, 560, y+14)
        else:
            c.setFont("Helvetica-Oblique", 10); c.setFillColorRGB(0, 0.5, 0); c.drawString(40, y, "Sin riesgos críticos.")
        
        y -= 15; y = dibujar_separador_local(y)

        # 5. HEATMAP DE EXCESOS (AZUL) - NUEVA SECCIÓN GRÁFICA
        if heatmap_data:
            if y < 150: c.showPage(); dibujar_encabezado_local(); y = height - 130
            c.setFillColorRGB(0, 0, 0); c.setFont("Helvetica-Bold", 14)
            c.drawString(30, y, f"5. Tendencia de Excesos (OSTN)")
            y -= 35
            
            # Leyenda Excesos
            c.setFont("Helvetica", 8)
            c.setFillColorRGB(0, 0.2, 0.6); c.rect(400, y+10, 8, 8, fill=True, stroke=False)
            c.setFillColorRGB(0, 0, 0); c.drawString(412, y+10, "High")
            c.setFillColorRGB(0.6, 0.8, 1.0); c.rect(440, y+10, 8, 8, fill=True, stroke=False)
            c.setFillColorRGB(0, 0, 0); c.drawString(452, y+10, "Med")
            
            # Dibujamos el gráfico en modo EXCESO
            y = dibujar_grafico_evolucion(heatmap_data, "Distribución de sobrantes por semana", y, "EXCESO", 110)
            y = dibujar_separador_local(y)

        # 6. PLAN DE ACCIÓN (SOLO SI HAY CAMBIOS)
        if tipo_reporte == "IDEAL":
            if y < 150: c.showPage(); dibujar_encabezado_local(); y = height - 130
            c.setFillColorRGB(0, 0, 0); c.setFont("Helvetica-Bold", 14)
            c.drawString(30, y, "6. Plan de Acción Detallado")
            y -= 25
            
            if not resumen_cambios:
                c.setFont("Helvetica-Oblique", 10); c.drawString(30, y, "El escenario ideal se logra sin movimientos adicionales.")
            
            for prod, cambios in resumen_cambios.items():
                if y < 100: c.showPage(); dibujar_encabezado_local(); y = height - 130
                c.setFillColorRGB(0, 0, 0.6); c.setFont("Helvetica-Bold", 10)
                c.drawString(30, y, f"PRODUCTO: {prod}"); y -= 15
                
                for cambio in cambios:
                    tipo = "ADELANTAR" if cambio['tipo'] == 'recorrido' else "PRODUCIR EXTRA"
                    cant = cambio['cantidad']
                    str_dest = cambio.get('str_destino', '??')
                    
                    c.setFont("Helvetica", 9); c.setFillColorRGB(0,0,0)
                    if tipo == "ADELANTAR":
                        str_orig = cambio.get('str_origen', '??')
                        c.setFillColorRGB(0, 0.5, 0)
                        txt = f">> ADELANTAR: {cant:,.0f} u. (De {str_orig} -> Para {str_dest})"
                    else:
                        c.setFillColorRGB(0.7, 0, 0)
                        txt = f">> PRODUCIR EXTRA: {cant:,.0f} u. (Para {str_dest})"
                    
                    c.drawString(40, y, txt)
                    y -= 12
                
                y -= 5; c.setStrokeColorRGB(0.9, 0.9, 0.9); c.line(30, y, width-30, y); y -= 15

    # -------------------------------------------------------------------------
    # BUCLE PRINCIPAL DE GENERACIÓN
    # -------------------------------------------------------------------------
    
    # 1. PÁGINA GLOBAL
    dibujar_seccion(data_pack_main['global'], 
                    data_pack_antes['global'] if data_pack_antes else None, 
                    "GLOBAL (Todos los Países)")
    
    # 2. PÁGINAS POR PAÍS Y CATEGORÍA
    odms_main = data_pack_main['odm']
    odms_antes = data_pack_antes['odm'] if data_pack_antes else {}
    odm_cat_main = data_pack_main['odm_cat']
    odm_cat_antes = data_pack_antes['odm_cat'] if data_pack_antes else {}
    
    lista_odms = sorted(odms_main.keys())
    
    for odm in lista_odms:
        # A) Página de Resumen de País
        c.showPage()
        nombre_pais = cfg.ODM_MAP.get(odm, odm)
        titulo_pais = f"{nombre_pais} ({odm})"
        dibujar_seccion(odms_main.get(odm), odms_antes.get(odm), titulo_pais, "Resumen Total País")
        
        # B) Páginas por Categoría
        claves_cat = [k for k in odm_cat_main.keys() if k[0] == odm]
        claves_cat.sort(key=lambda x: x[1])
        
        for key in claves_cat:
            c.showPage()
            cat_nombre = key[1]
            dibujar_seccion(odm_cat_main.get(key), odm_cat_antes.get(key), titulo_pais, f"Categoría: {cat_nombre}")

    c.save()