import os
import re
import time
import requests
from bs4 import BeautifulSoup
import pandas as pd
import cv2
import numpy as np

class SeniatOCR:
    def __init__(self):
        """Inicializa el bot de consulta con sesión global."""
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'es-ES,es;q=0.8,en;q=0.5,en-US;q=0.3'
        })
        
        self.url_base = "http://contribuyente.seniat.gob.ve/BuscaRif/BuscaRif.jsp"
        self.url_captcha = "http://contribuyente.seniat.gob.ve/BuscaRif/Captcha.jpg"
        
        print("[*] Inicializando motor de reconocimiento visual...")
        import easyocr
        self.reader = easyocr.Reader(['en', 'es'], gpu=False, verbose=False)
        
        self.carpeta_tmp = 'tmp'
        if not os.path.exists(self.carpeta_tmp):
            os.makedirs(self.carpeta_tmp)

    def calcula_digito_cedula(self, c):
        """Calcula el dígito verificador para cédulas que vienen sin él."""
        l = c[0].upper()
        n_str = c[1:]
        n = n_str.zfill(8)
        val = l + n + '1'
        digitos = list(val)
        
        digitos_mult = [
            0,
            int(digitos[1]) * 3,
            int(digitos[2]) * 2,
            int(digitos[3]) * 7,
            int(digitos[4]) * 6,
            int(digitos[5]) * 5,
            int(digitos[6]) * 4,
            int(digitos[7]) * 3,
            int(digitos[8]) * 2
        ]
        digito_especial = 1 if digitos[0] == 'V' else (2 if digitos[0] == 'E' else 0)
        suma = (sum(digitos_mult)) + (digito_especial * 4)
        residuo = suma % 11
        resta = 11 - residuo
        digito_verificador = 0 if resta >= 10 else resta
        return f"{l}{n}{digito_verificador}"

    def _limpiar_captcha(self, ruta_original, ruta_procesada):
        """Procesa y limpia el CAPTCHA para mejorar la precisión del OCR."""
        if not os.path.exists(ruta_original) or os.path.getsize(ruta_original) == 0:
            return False
        img = cv2.imread(ruta_original, cv2.IMREAD_GRAYSCALE)
        if img is None:
            return False
        img = cv2.resize(img, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_LANCZOS4)
        blur = cv2.GaussianBlur(img, (3, 3), 0)
        _, thresh = cv2.threshold(blur, 140, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        kernel = np.ones((2, 2), np.uint8)
        cierre = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)
        cv2.imwrite(ruta_procesada, cierre)
        return True

    def consultar_rif(self, doc):
        """Ejecuta la consulta y extrae todos los campos fiscales del SENIAT."""
        doc_limpio = str(doc).upper().replace("-", "").replace(".", "").strip()
        
        if re.match(r'^[JGP]', doc_limpio):
            documento = doc_limpio
        else:
            if not re.match(r'^[VE]', doc_limpio):
                doc_limpio = 'V' + doc_limpio
            
            letra = doc_limpio[0]
            numeros = doc_limpio[1:]
            
            if len(numeros) == 9:
                documento = doc_limpio
            elif len(numeros) in [7, 8]:
                documento = self.calcula_digito_cedula(doc_limpio)
            else:
                cedula_base = re.sub(r'[^0-9]', '', numeros)[:8]
                documento = self.calcula_digito_cedula(letra + cedula_base)

        ruta_orig = os.path.join(self.carpeta_tmp, 'captcha_orig.jpg')
        ruta_clean = os.path.join(self.carpeta_tmp, 'captcha_clean.jpg')
        
        try:
            self.session = requests.Session()
            self.session.headers.update({
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
                'Referer': 'http://contribuyente.seniat.gob.ve/BuscaRif/BuscaRif.jsp'
            })
            
            res_captcha = self.session.get(self.url_captcha, timeout=25)
            if res_captcha.status_code != 200:
                return {"status": "fail_ocr", "message": "Fallo descarga CAPTCHA."}
                
            with open(ruta_orig, 'wb') as f:
                f.write(res_captcha.content)
            
            if not self._limpiar_captcha(ruta_orig, ruta_clean):
                return {"status": "fail_ocr", "message": "Fallo al procesar CAPTCHA."}

            resultados = self.reader.readtext(ruta_clean, detail=0, allowlist='0123456789abcdefghijklmnopqrstuvwxyz')
            texto_captcha = "".join(resultados).strip().lower().replace(" ", "")
            
            print(f"    [OCR] Texto detectado: '{texto_captcha}'")

            if not texto_captcha or len(texto_captcha) < 4:
                return {"status": "fail_ocr", "message": "OCR vacío o incompleto."}
            
            payload = {'codigo': texto_captcha, 'p_rif': documento}
            respuesta_final = self.session.post(self.url_base, data=payload, timeout=25)
            html_resultado = respuesta_final.text
            
            if any(err in html_resultado.lower() for err in ["código de seguridad incorrecto", "código incorrecto", "código no coincide", "no coincide con la imagen"]):
                return {"status": "fail_ocr", "message": "CAPTCHA rechazado por el servidor.", "intentado": texto_captcha}
            
            if "No existe el Contribuyente" in html_resultado or "El Contribuyente no existe" in html_resultado:
                return {"status": "not_found", "message": "No existe."}
            
            soup = BeautifulSoup(html_resultado, 'html.parser')
            tables = soup.find_all('table')

            rif_encontrado = documento
            nombre_limpio = ""
            siglas = ""
            actividad_economica = ""
            condicion = "NO SUJETO / OTROS"
            agente_retencion = ""

            # Recorrer las tablas buscando los datos fiscales exactos
            for tbl in tables:
                txt = tbl.get_text(separator=" ", strip=True)
                txt = re.sub(r'\s+', ' ', txt)

                # Buscar RIF y Nombre Principal
                if documento in txt and not nombre_limpio:
                    match_rn = re.search(r'([VEJPG]\d{9,10})\s+(.+)', txt)
                    if match_rn:
                        rif_encontrado = match_rn.group(1)
                        nombre_limpio = match_rn.group(2).replace("Firmas Personales", "").strip()

                # Buscar Actividad Económica
                if "Actividad Económica:" in txt:
                    m_act = re.search(r'Actividad Económica:\s*(.*?)(?=\s+Condición:|\s+La condición|$)', txt, re.IGNORECASE)
                    if m_act:
                        actividad_economica = m_act.group(1).strip()
                    else:
                        partes_act = txt.split("Actividad Económica:")
                        if len(partes_act) > 1:
                            actividad_economica = partes_act[1].split("La condición")[0].split("Condición")[0].strip()

                # Buscar Condición del IVA
                if "Condición:" in txt:
                    m_cond = re.search(r'Condición:\s*(.*?)(?=\s+La condición|$)', txt, re.IGNORECASE)
                    if m_cond:
                        condicion = m_cond.group(1).strip()
                elif "Contribuyente Ordinario" in txt:
                    condicion = "CONTRIBUYENTE ORDINARIO"
                elif "Contribuyente Formal" in txt:
                    condicion = "CONTRIBUYENTE FORMAL"

                # Buscar Texto exclusivo de Retención (separado de lo demás)
                if "La condición de este contribuyente" in txt:
                    m_ret = re.search(r'(La condición de este contribuyente.*)', txt, re.IGNORECASE)
                    if m_ret:
                        agente_retencion = m_ret.group(1).strip()

            # Limpieza final de caracteres extra
            nombre_limpio = re.sub(r'_+', '', nombre_limpio).strip()
            actividad_economica = re.sub(r'_+', '', actividad_economica).strip()
            agente_retencion = re.sub(r'_+', '', agente_retencion).strip()

            # Detección de IVA y Retención
            es_ordinario_o_formal = "ORDINARIO" in condicion.upper() or "FORMAL" in condicion.upper() or "ORDINARIO" in agente_retencion.upper()
            requiere_retencion = "REQUIERE LA RETENCIÓN" in agente_retencion.upper() or "100%" in agente_retencion or "75%" in agente_retencion

            return {
                "status": "success",
                "rif": rif_encontrado,
                "nombre": nombre_limpio,
                "siglas": siglas,
                "actividad_economica": actividad_economica.upper(),
                "condicion": condicion.upper(),
                "retencion": agente_retencion.upper(),
                "contribuyente_iva": es_ordinario_o_formal,
                "agente_retencion": requiere_retencion
            }
            
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def consultar_con_reintentos(self, rif, max_intentos=8):
        """Bucle de reintentos automáticos integrado."""
        intento = 1
        while intento <= max_intentos:
            print(f"    [Intento {intento}/{max_intentos}] Consultando...")
            resultado = self.consultar_rif(rif)
            if resultado["status"] in ["success", "not_found"]:
                return resultado
            if resultado["status"] == "fail_ocr":
                intento += 1
                time.sleep(1)
                continue
            return resultado
        return {"status": "error", "message": f"Superado el límite de {max_intentos} reintentos."}