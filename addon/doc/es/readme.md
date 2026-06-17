# simpleCopy

**Copie, añada y gestione texto de manera eficiente con NVDA**

**autor:** chai chaimee  
**url:** https://github.com/chaichaimee/simpleCopy

---

## Descripción

**simpleCopy** es un complemento ligero para NVDA que simplifica la copia de texto, la extracción de URL y la gestión del historial de voz.

Esta herramienta le ayuda a capturar y organizar información rápidamente sin interrumpir su flujo de trabajo. Ya sea que esté copiando texto, capturando enlaces web o guardando contenido hablado, simpleCopy proporciona atajos de teclado intuitivos que funcionan perfectamente con NVDA.

---

## Atajos de teclado

Todos los comandos utilizan un sistema de pulsaciones múltiples. Pulse la combinación de teclas una, dos o tres veces en rápida sucesión para realizar diferentes acciones.

### CTRL+Mayús+A — Captura de URL y enlaces

- **Una pulsación:** Copia la URL de la página web actual.
- **Dos pulsaciones:** Copia la URL de destino del hipervínculo enfocado.

### CTRL+Mayús+V — Copiar, añadir y gestionar el portapapeles

- **Una pulsación:** Copia el texto seleccionado. Si ya hay texto en el portapapeles, la nueva selección se añade a él.
- **Dos pulsaciones:** Copia el texto desde la posición actual del cursor de revisión. Funciona con cualquier selección realizada con el cursor de revisión de NVDA, incluyendo selecciones de varias líneas y selección de documento completo.
- **Tres pulsaciones:** Borra todo el contenido del portapapeles.

### F9 — Captura y gestión de voz

- **Una pulsación:** Copia la salida de voz más reciente de NVDA.
- **Dos pulsaciones:** Añade la salida de voz más reciente al contenido existente del portapapeles.
- **Tres pulsaciones:** Copia toda la salida de voz acumulada desde la primera pulsación de F9.

### Mayús+F9 — Navegación por el historial de voz

- **Una pulsación:** Navega al elemento anterior del historial de voz.
- **Dos pulsaciones:** Navega al siguiente elemento del historial de voz.
- **Tres pulsaciones:** Abre el archivo de registro completo del historial de voz.

---

## Características

Así es como funciona cada característica en la práctica:

### 1. Copiar URL de página web

Pulse **CTRL+Mayús+A una vez** mientras navega por cualquier sitio web. La URL de la página actual se copia en su portapapeles. NVDA lo confirma leyendo la URL copiada.

### 2. Extraer URL de hipervínculo

Enfoque cualquier enlace y pulse **CTRL+Mayús+A dos veces**. La URL de destino se extrae y se copia sin abrir el enlace.

### 3. Copiar y añadir texto

Seleccione texto y pulse **CTRL+Mayús+V una vez**. Si el portapapeles está vacío, el texto se copia. Si el portapapeles ya contiene texto, la nueva selección se añade con un salto de línea.

### 4. Copiar desde el cursor de revisión

Utilice el cursor de revisión de NVDA para seleccionar texto (usando NVDA+Mayús+Flecha abajo o NVDA+CTRL+Mayús+Flecha abajo para seleccionar varias líneas) y, a continuación, pulse **CTRL+Mayús+V dos veces**. Todo el texto seleccionado desde la posición del cursor de revisión se copia al portapapeles. Funciona con cualquier tamaño de selección, desde una sola palabra hasta un documento completo.

### 5. Borrar portapapeles

Pulse **CTRL+Mayús+V tres veces** para borrar instantáneamente todo el contenido del portapapeles. NVDA lo confirma con el mensaje "Clean".

### 6. Copiar última voz

Cuando NVDA diga algo que desee guardar, pulse **F9 una vez**. La última frase hablada se copia en su portapapeles.

### 7. Añadir voz

Pulse **F9 dos veces** para añadir la última frase hablada al contenido existente del portapapeles.

### 8. Registrar historial de voz

Pulse **F9 tres veces** para copiar toda la salida de voz acumulada durante su sesión actual.

### 9. Navegar por el historial de voz

Utilice **Mayús+F9 una vez** para retroceder en el historial de voz, y **Mayús+F9 dos veces** para avanzar. Esto le permite revisar la salida de voz anterior sin cambiar su enfoque actual.

### 10. Acceder al archivo de registro de voz

Pulse **Mayús+F9 tres veces** para abrir el archivo completo del historial de voz en su editor de texto predeterminado para revisarlo, buscarlo o copiarlo.

### 11. Conciencia de contexto inteligente

Cuando está escribiendo en campos editables, simpleCopy no interfiere. Los comandos solo se activan cuando son útiles, preservando su flujo de trabajo normal.

---

## Apóyeme

Si este complemento le ayuda a trabajar de manera más eficiente, considere hacer una pequeña donación para apoyar el desarrollo futuro.

[![Apóyame](https://img.shields.io/badge/Donate-Support%20Me-blue?style=for-the-badge&logo=stripe)](https://buy.stripe.com/dRm9AU1xQ3Ds22N6VK1VK01)

Su apoyo ayuda a mantener este proyecto vivo y en mejora.

---

© 2026 Chai Chaimee Complemento para NVDA Publicado bajo la Licencia Pública General de GNU