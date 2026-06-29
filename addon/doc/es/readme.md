# BOA: Better Office Accessibility

BOA es un potente conjunto de mejoras de accesibilidad para Microsoft Office, diseñado para mejorar enormemente la experiencia del lector de pantalla para los usuarios de NVDA. Parchea directamente los componentes de interfaz inaccesibles e introduce herramientas de navegación rápida para Excel y PowerPoint.

---

## ⌨️ Referencia de teclas rápidas

| Función | Combinación de teclas | Contexto / Notas |
| :--- | :--- | :--- |
| **Entrar en modo de comando** | `[Prefix]` (Por defecto: `NVDA+E`) | Activa el modo de prefijo de comando (emite un pitido agudo) |
| **Cancelar modo de comando** | `Escape` | Sale del modo de prefijo de comando |
| **MEJORAS DE EXCEL** | | |
| **Analizar diseño de hoja** | `[Prefix]`, luego `L` | Ejecutar dentro de Excel antes de navegar bloques de datos |
| **Saltar al bloque de datos más cercano** | `[Prefix]`, luego `J` | Requiere análisis de diseño previo |
| **Abrir organizador masivo de hojas** | `[Prefix]`, luego `X` | Abre el diálogo accesible de reordenación de hojas |
| **Anunciador de fórmula sin procesar** | `[Prefix]`, luego `F2` | Pulse una vez para escuchar la cadena de fórmula sin procesar |
| **Editor avanzado de fórmulas** | `[Prefix]`, luego `F2` dos veces | Pulse dos veces para abrir el editor de fórmulas multilínea accesible |
| **Rastrear precedentes** | `[Prefix]`, luego `Shift+P` | Rastrear precedentes, la misma función de forma accesible.|
| **Rastrear dependientes** | `[Prefix]`, luego `Shift+D` | Rastrear dependientes de forma accesible; al presionar Enter sobre una celda, se teletransportará a ella.|
| **Formato condicional detallado**| `[Prefix]`, luego `F` | Anuncia los detalles completos de formato de la celda enfocada |
| **Mover hoja activa a la izquierda** | `NVDA+Shift+LeftArrow` | Desplaza la hoja activa una posición hacia arriba |
| **Mover hoja activa a la derecha** | `NVDA+Shift+RightArrow` | Desplaza la hoja activa una posición hacia abajo |
| **Mover hoja al inicio/final** | `NVDA+Shift+Home` / `End` | Envía la hoja a los extremos absolutos |
| **Ocultar / Mostrar fila** | `Ctrl+9` / `Ctrl+Shift+9` | Atajo nativo; BOA anuncia explícitamente el cambio de visibilidad |
| **Ocultar / Mostrar columna** | `Ctrl+0` / `Ctrl+Shift+0` | Atajo nativo; BOA anuncia explícitamente el cambio de visibilidad |
| **Mostrar columna (alternativo)** | `NVDA+Ctrl+Shift+0` | Evita conflictos con la tecla rápida del idioma de entrada de Windows |
| **Asignar celda a ranura de memoria** | `[Prefix]`, luego `Shift+1` a `Shift+9` | Asigna la celda actual a una ranura de monitoreo en segundo plano |
| **Leer ranura de celda monitoreada** | `[Prefix]`, luego `1` a `9` | Recupera y lee el valor de la ranura asignada |
| **Salto directo a ranura** | `Alt` + `1` a `9` | Salta instantáneamente el cursor a una ranura monitoreada |
| **Regresar a la celda anterior** | `[Prefix]`, luego `\` | Teletransporta instantáneamente de vuelta después de consultar una ranura |
| **Diálogo del gestor de ranuras** | `[Prefix]`, luego `Alt+M` | Abre un diálogo para ver y gestionar todos los monitores activos |
| **Alternar monitoreo en segundo plano** | `[Prefix]`, luego `M` | Alterna manualmente el seguimiento de cálculos en segundo plano |
| **Borrar todas las ranuras de memoria** | `[Prefix]`, luego `Backspace` | Elimina todos los monitores de celdas en segundo plano guardados |
| **MEJORAS DE POWERPOINT** | | |
| **Analizador de diseño de diapositiva** | `[Prefix]`, luego `L` | Analiza y anuncia el diseño espacial de la diapositiva actual |
| **Analizador de documento** | `[Prefix]`, luego `D` | Genera un índice completo y un informe de estado |
| **Organizador masivo de diapositivas** | `[Prefix]`, luego `X` | Abre el diálogo accesible para reordenar múltiples diapositivas |
| **MEJORAS DE WORD** | | |
| **Auditor de formato** | `[Prefix]`, luego `F` | Audita el documento actual en busca de inconsistencias de formato |
| **Analizador de documento** | `[Prefix]`, luego `D` | Analiza el diseño y la estructura del documento de Word actual |

---

## 🚀 Características

### Mejoras de Excel

#### 1. Analizador y caché de diseño de hoja
Escanee instantáneamente cualquier hoja de cálculo de Excel para comprender su estructura, elementos ocultos y bloques de datos.
* **Cómo funciona:** BOA escanea rápidamente la hoja y anuncia los bloques de datos activos. También le advierte sobre **pestañas de hojas ocultas**, **filtros** activos, **modos de protección** y **bordes exteriores ocultos** (p. ej., si las columnas cerca del borde derecho de la hoja están ocultas, evitando que se pierda datos fuera de la pantalla).
* **Navegación de datos:** Después del escaneo, puede usar las teclas rápidas de salto entre bloques de datos para desplazar instantáneamente el cursor entre los bloques de datos descubiertos, evitando sin esfuerzo miles de celdas vacías.

#### 2. Organizador masivo de hojas
Reordene y organice instantáneamente múltiples hojas a la vez usando un diálogo completamente accesible.
* **Cómo funciona:** Abre un diálogo donde puede seleccionar una hoja y asignarle una nueva posición. Los movimientos programados se listan en una tabla de datos (presione `Del` para eliminar un error). Haga clic en `OK` y su libro de trabajo se reorganizará instantáneamente.

#### 3. Movimiento rápido de hojas
Mueva la hoja activa a la izquierda, a la derecha, al principio o al final instantáneamente usando sus atajos de teclado.

#### 4. Renombrado accesible de hojas
* Al renombrar una hoja, NVDA tiene dificultades nativas para leer los caracteres que está escribiendo.
* BOA inyecta una clase personalizada `ExcelSheetRenameEdit` que utiliza el motor `SafeRichEdit`, lo que significa que puede leer con precisión por carácter, palabra o línea mientras renombra. Esto funciona como una mejora del comportamiento de renombrado predeterminado existente.

#### 5. Rastreador de filas/columnas ocultas
* Rastrea proactivamente su movimiento a través de la cuadrícula para evitar que se pierdan datos ocultos o filtrados.
* **Celdas fragmentadas cruzadas:** Si salta a través de una sección muy fragmentada u oculta de la cuadrícula (p. ej., moviéndose de la Fila 3 a la Fila 10 porque las Filas 4–9 están ocultas), BOA anuncia explícitamente "Filas 4 a 9 ocultas". Esto garantiza que siempre sepa cuándo se han omitido datos en la estructura.

#### 6. Anunciador de formato condicional
* Lee automáticamente el color, el estilo de fuente y la sombra de fondo de las celdas que han sido modificadas dinámicamente por las reglas de formato condicional de Excel.
* Le proporciona el estado visual real de la celda en lugar de solo el valor subyacente sin procesar. Inicialmente, al enfocar la celda, anuncia "tiene formato condicional y algunos otros detalles menores". Para información completa, use la configuración de tecla rápida detallada que es NVDA E y F.

#### 7. Mejor anuncio de selección
Lee si la celda o rango está seleccionado o deseleccionado.

#### 8 Monitor de celdas:
* **Monitor de celdas:** Use rutas de comando para asignar celdas específicas a ranuras de memoria. Puede volver y leerlas en cualquier momento usando la ranura numérica asignada.
* **Monitoreo continuo:** Las celdas asignadas se monitorean automáticamente en segundo plano. Si Excel activa un recálculo o una edición de celda, BOA anuncia instantáneamente el nuevo valor. Alterne manualmente o borre todo mediante las ranuras de comando.
* **Excel: Mejoras del Monitor de Celdas Pro:**
  - **Diálogo del gestor de ranuras (`NVDA+E`, luego `Alt+M`):** Abre un diálogo que lista todas sus celdas monitoreadas activamente. Presione `Enter` para saltar instantáneamente a una.
  - **Regresar (`NVDA+E`, luego `\`):** Le teletransporta instantáneamente de vuelta a su celda de trabajo anterior después de consultar una ranura.
  - **Salto directo a ranura (`Prefix + Alt` + `Número de ranura`):** Salte directamente a una ranura de celda asignada.

#### 9 Editor avanzado
* **Excel: El Editor Avanzado (Editor de Fórmulas Accesible):** Un cambio absoluto para modificar fórmulas extensas.
  - **Pulsación simple `NVDA+E`, luego `F2`:** Anuncia instantáneamente la cadena de fórmula sin procesar de la celda activa (o anuncia "Sin fórmula").
  - **Doble pulsación `NVDA+E`, luego `F2`:** Abre un editor multilínea completamente accesible para modificar de forma segura fórmulas extensas y anidadas. `Enter` nativo añade saltos de línea para facilitar la lectura, y `Ctrl+Enter` guarda los cambios de vuelta en Excel.
  - *Comprobaciones de seguridad:* Captura de forma segura los errores de sintaxis antes de que corrompan su hoja, y detecta errores posteriores al cálculo (como `#NAME?` o `#DIV/0!`) para advertirle instantáneamente si una fórmula falló.

#### 10 Mejoras de auditoría y evaluación de fórmulas:
* **Excel: Auditoría y evaluación de fórmulas:** Se añadieron atajos personalizados (`NVDA+E`, luego `Shift+P` y `NVDA+E`, luego `Shift+D`) para rastrear de forma fiable precedentes y dependientes. Además, el diálogo nativo "Evaluar fórmula" de Excel ahora es completamente accesible; ¡NVDA lee automáticamente los resultados evaluados a medida que avanza paso a paso por el cálculo!

### Mejoras de PowerPoint

#### 1. Selectores de color accesibles
* Desbloquea el diálogo de color personalizado en PowerPoint.
* Identifica y lee explícitamente los cuadros de edición "Rojo", "Verde" y "Azul" correctamente (anulando `PowerPointRGBEdit`).
* Mapea el campo de entrada Hex previamente invisible para que NVDA pueda leer el valor de color Hex completo de forma limpia.

#### 2. Soporte de cuadrícula de colores estándar
* Navegar por la cuadrícula hexagonal de colores "Estándar" de PowerPoint normalmente se lee como "Gráfico" o silencio.
* BOA rastrea sus teclas de flecha a través del hexágono y obtiene silenciosamente el valor de color oculto, anunciándolo en tiempo real (p. ej., "Color #FF0000").

#### 3 Organizador masivo de diapositivas:
* **PowerPoint: Organizador masivo de diapositivas (experimental) (`NVDA+E`, luego `X`):** Similar a la función de Excel, ahora puede reordenar, mover y organizar instantáneamente múltiples diapositivas de PowerPoint a la vez usando un diálogo completamente accesible.

#### 4 Analizador de diseño de diapositiva
* **PowerPoint: Analizador de diseño de diapositiva (experimental) (`NVDA+E`, luego `L`):** Escanea instantáneamente su diapositiva activa actual para comprender su diseño espacial y restricciones de accesibilidad, garantizando una experiencia de lector de pantalla completamente fluida y receptiva. Es decir, aquí obtendrá detalles sobre la diapositiva actual similar al analizador de diseño de hoja de Excel.


#### 5 Analizador completo de documento [PPT]
* **PowerPoint: Analizador completo de documento (experimental) (`NVDA+E`, luego `D`):** Una herramienta de accesibilidad altamente avanzada, procesada en segundo plano, que mapea toda una presentación sin congelar el motor de voz de NVDA. Proporciona un índice virtual profundamente navegable, detecta desajustes en el orden de lectura (orden visual vs. orden Z), marca diapositivas con "muro de texto" y mapea objetos complejos como SmartArt y tablas de datos.

#### 6 Mejoras de movimiento [ajuste] de formas:
* **PowerPoint: Modo de audio de movimiento de formas (experimental):** Introduce señales de audio espacial 3D en el lienzo de PowerPoint. Proporciona retroalimentación auditiva que indica la dirección y los límites de un objeto a medida que lo mueve, mejorando enormemente la conciencia espacial.

### Mejoras de Word:
#### 1. Document Analyzer inspirado y derivado del complemento Word Access de Paul:
* **Word: Analizador de documento (`NVDA+E`, luego `D`):** Obtenga instantáneamente una visión general estructural de su documento de Word. *(Un agradecimiento especial a Paul: Esta función fue directamente inspirada por su brillante complemento "Word Access". ¡Estamos profundamente agradecidos por su trabajo fundamental en este ámbito!)*

#### 2 Auditor de formato
* **Word: Auditor de formato (`NVDA+E`, luego `F`):** Escanea su documento de Word en busca de inconsistencias de formato para garantizar estándares visuales.

#### 3 Lector de notas al pie:
* **Word: Anunciador automático de notas al pie:** Las notas al pie ahora se anunciarán automáticamente en línea mientras lee, dependiendo de su configuración personalizada de BOA. *(Nota: El soporte para notas finales y comentarios está planificado para una versión futura).*

### Infraestructura y mecanismos técnicos

#### El modo de prefijo de comando
Para prevenir conflictos de atajos de teclado con otros complementos de NVDA, BOA utiliza un **modo de prefijo de comando**:
1. Presione la tecla de activación para entrar en el modo de comando. Escuchará un pitido agudo. El valor predeterminado es NVDA más E.
2. Presione una tecla secundaria para activar una función específica.
3. Si presiona una tecla no válida, escuchará un pitido de error.

#### Personalización y panel de configuración
* Las funciones de BOA son completamente modulares y se pueden habilitar o deshabilitar en cualquier momento. Vaya a `Menú NVDA -> Preferencias -> Opciones -> BOA Office Enhancements` para activar o desactivar funciones individuales.
* **Teclas aceleradoras inteligentes:** Cada configuración cuenta con un atajo acelerador `Alt+Tecla` matemáticamente único dentro del panel. Por ejemplo, presione `Alt+E` para saltar instantáneamente al grupo de Excel, `Alt+P` para PowerPoint y `Alt+W` para Word.
* La configuración se guarda de forma segura en un archivo JSON independiente (`boa_settings.json`), garantizando que su configuración principal de NVDA nunca se corrompa.
* Si Microsoft Office corrige oficialmente un error de accesibilidad en el futuro, puede deshabilitar de forma segura el gancho de anulación específico de BOA sin perder el resto de la funcionalidad del complemento.
* **Personalización de gestos de entrada:** Todas las funciones en todas las aplicaciones de Office han sido explícitamente expuestas al diálogo nativo de Gestos de Entrada de NVDA, otorgándole libertad completa para personalizar cada atajo de teclado.

#### Seguridad y límites de integración
* Las inyecciones de portapapeles verifican estrictamente los identificadores de proceso de la ventana en primer plano para evitar la fuga de datos a otras aplicaciones.
* Algunos atajos de teclado personalizados están completamente expuestos en el diálogo de Gestos de Entrada de NVDA bajo la categoría "Better Office Accessibility".

---

## 📋 Requisitos

* **NVDA:** Versión 2026.1.0 o posterior.
* **Aplicaciones:** Microsoft Excel y Microsoft PowerPoint.

---

## 💾 Instalación

1. Descargue el archivo de la última versión `.nvda-addon`, o localícelo dentro de la tienda de complementos nativa de NVDA.
2. Si instala desde un archivo, abra el archivo o use `Tienda de complementos de NVDA -> Instalar desde archivo externo`.
3. Reinicie NVDA.

---

## 🛠️ Registro de cambios

### Versión 2.0.0
#### Nuevas funciones
* **PowerPoint: Analizador completo de documento (experimental) (`NVDA+E`, luego `D`):** Una herramienta de accesibilidad altamente avanzada, procesada en segundo plano, que mapea toda una presentación sin congelar el motor de voz de NVDA. Proporciona un índice virtual profundamente navegable, detecta desajustes en el orden de lectura (orden visual vs. orden Z), marca diapositivas con "muro de texto" y mapea objetos complejos como SmartArt y tablas de datos.
* **PowerPoint: Analizador de diseño de diapositiva (experimental) (`NVDA+E`, luego `L`):** Escanea instantáneamente su diapositiva activa actual para comprender su diseño espacial y restricciones de accesibilidad, garantizando una experiencia de lector de pantalla completamente fluida y receptiva. Es decir, aquí obtendrá detalles sobre la diapositiva actual similar al analizador de diseño de hoja de Excel.
* **PowerPoint: Organizador masivo de diapositivas (experimental) (`NVDA+E`, luego `X`):** Similar a la función de Excel, ahora puede reordenar, mover y organizar instantáneamente múltiples diapositivas de PowerPoint a la vez usando un diálogo completamente accesible.
* **PowerPoint: Modo de audio de movimiento de formas (experimental):** Introduce señales de audio espacial 3D en el lienzo de PowerPoint. Proporciona retroalimentación auditiva que indica la dirección y los límites de un objeto a medida que lo mueve, mejorando enormemente la conciencia espacial. Como se mencionó, esto es experimental; esperamos retroalimentación para mejorarlo.
* **Word: Auditor de formato (`NVDA+E`, luego `F`):** Escanea su documento de Word en busca de inconsistencias de formato para garantizar estándares visuales.
* **Word: Analizador de documento (`NVDA+E`, luego `D`):** Obtenga instantáneamente una visión general estructural de su documento de Word. *(Un agradecimiento especial a Paul: Esta función fue directamente inspirada por su brillante complemento "Word Access". ¡Estamos profundamente agradecidos por su trabajo fundamental en este ámbito!)*
* **Word: Anunciador automático de notas al pie:** Las notas al pie ahora se anunciarán automáticamente en línea mientras lee, dependiendo de su configuración personalizada de BOA. *(Nota: El soporte para notas finales y comentarios está planificado para una versión futura).*
* **Excel: El Editor Avanzado (Editor de Fórmulas Accesible):** Un cambio absoluto para modificar fórmulas extensas.
  - **Pulsación simple `NVDA+E`, luego `F2`:** Anuncia instantáneamente la cadena de fórmula sin procesar de la celda activa (o anuncia "Sin fórmula").
  - **Doble pulsación `NVDA+E`, luego `F2`:** Abre un editor multilínea completamente accesible para modificar de forma segura fórmulas extensas y anidadas. `Enter` nativo añade saltos de línea para facilitar la lectura, y `Ctrl+Enter` guarda los cambios de vuelta en Excel.
  - *Comprobaciones de seguridad:* Captura de forma segura los errores de sintaxis antes de que corrompan su hoja, y detecta errores posteriores al cálculo (como `#NAME?` o `#DIV/0!`) para advertirle instantáneamente si una fórmula falló.
* **Excel: Auditoría y evaluación de fórmulas:** Se añadieron atajos personalizados (`NVDA+E`, luego `Shift+P` y `NVDA+E`, luego `Shift+D`) para rastrear de forma fiable precedentes y dependientes. Además, el diálogo nativo "Evaluar fórmula" de Excel ahora es completamente accesible; ¡NVDA lee automáticamente los resultados evaluados a medida que avanza paso a paso por el cálculo!
* **Excel: Mejoras del Monitor de Celdas Pro:**
  - **Diálogo del gestor de ranuras (`NVDA+E`, luego `Alt+M`):** Abre un diálogo que lista todas sus celdas monitoreadas activamente. Presione `Enter` para saltar instantáneamente a una.
  - **Regresar (`NVDA+E`, luego `\`):** Le teletransporta instantáneamente de vuelta a su celda de trabajo anterior después de consultar una ranura.
  - **Salto directo a ranura (`Alt` + `Número de ranura`):** Salte directamente y al instante a una ranura de celda asignada.
* **Personalización de gestos de entrada:** Todas las funciones en todas las aplicaciones de Office han sido explícitamente expuestas al diálogo nativo de Gestos de Entrada de NVDA, otorgándole libertad completa para personalizar cada atajo de teclado.

#### UX/UI Enhancements
* **Informes navegables unificados:** Hemos adoptado un sistema de informes HTML unificado en todo el complemento. Funciones como el anunciador de formato condicional de Excel, los analizadores de diseño y los analizadores de documentos ya no solo leen bloques masivos de texto; sus resultados ahora se abren en una ventana HTML nativa y navegable, permitiéndole revisar los datos a su propio ritmo.
* **Excel: Seguimiento mejorado de dependientes/precedentes:** Se mejoró enormemente la salida de voz para los atajos nativos de rastreo de fórmulas de Excel (`Ctrl+[` para precedentes directos y `Ctrl+]` para dependientes directos). NVDA ahora anuncia explícitamente exactamente qué celdas fueron seleccionadas.
* **Excel: Soporte de celdas combinadas:** Las celdas combinadas ahora se detectan correctamente y se anuncian explícitamente por el rastreador de celdas con salto de espacios.

#### Corrección de errores
* **Word: Doble lectura de elementos de lista:** Se implementó un parche temporal para corregir el error donde NVDA lee dos veces los elementos de lista de párrafos en ciertas vistas de Word.
* **Excel: Error de localización del monitor de celdas:** Se resolvieron errores de seguimiento subyacentes causados por las recientes actualizaciones de localización de traducción.

### Novedades en v1.6.1
* **Localización profunda de archivos**: Se corrigieron traducciones de cadenas faltantes en las profundidades de los módulos de mejora de Excel (como el analizador de diseño de hoja y el movimiento rápido de hojas) para garantizar una cobertura de localización del 100%.
* **Soporte de traducción ampliado**: Se añadieron 7 nuevos idiomas al sistema (turco, polaco, coreano, ucraniano, checo, urdu y panyabí).
  *(Nota: Estas traducciones fueron generadas por IA, por lo que pueden contener algunos errores o imprecisiones menores de traducción.)*

### v1.6.0
* **Soporte de traducción integral**: El complemento ahora está completamente localizado con soporte para 17 idiomas globales.
  *(Nota: Estas traducciones fueron generadas por IA, por lo que pueden contener algunos errores o imprecisiones menores de traducción.)*
* **Gobernanza estricta del código**: Se aplicaron encabezados de derechos de autor GPL-2.0 en toda la base de código.""",

### Versión 1.5.0
#### Nuevas funciones
##### Radar de fin de datos
Al navegar por hojas de cálculo extensas, puede ser difícil saber si una celda vacía significa que ha llegado al final de una lista o si simplemente hay un espacio en los datos. El **Radar de fin de datos** actúa como una verificación de perímetro inteligente para evitar que navegue a ciegas por el espacio vacío.
Cada vez que navega hacia una celda vacía, BOA escanea instantáneamente las celdas restantes en su dirección de viaje. Si no quedan datos en absoluto, anunciará proactivamente:
* *"No hay más datos abajo"*
* *"No hay más datos arriba"*
* *"No hay más datos a la derecha"*
* *"No hay más datos a la izquierda"*
**Opciones de configuración:**
Puede configurar esta función a través de `Preferencias de NVDA -> Opciones -> BOA Office Enhancements`. Debido a que las hojas de cálculo pueden contener complejidades ocultas (como fórmulas invisibles o filas contraídas), el radar proporciona tres modos de operación:
1. **Desactivado**: Deshabilita el radar por completo.
2. **Verificación estricta de memoria (CountA) [Predeterminado]**: El enfoque más seguro y rápido. Verifica la memoria sin procesar de la hoja de cálculo. Si detecta *cualquier cosa* debajo de usted (incluidas filas ocultas, texto, números o fórmulas invisibles), permanece completamente en silencio para evitar falsas alarmas. Solo anuncia "No hay más datos" cuando el resto de la hoja está 100% matemáticamente en blanco.
3. **Solo datos visibles (Motor matemático)**: Un motor altamente avanzado diseñado para hojas complejas. Filtra inteligentemente las filas ocultas y las fórmulas invisibles (p. ej., `=""`). Solo permanecerá en silencio si quedan números o texto reales y visibles en su camino.

### Versión 1.4 - 2026-06-12
#### Nuevas funciones
* **Monitor de celdas:** Use rutas de comando para asignar celdas específicas a ranuras de memoria. Puede volver y leerlas en cualquier momento usando la ranura numérica asignada.
* **Monitoreo continuo:** Las celdas asignadas se monitorean automáticamente en segundo plano. Si Excel activa un recálculo o una edición de celda, BOA anuncia instantáneamente el nuevo valor. Alterne manualmente o borre todo mediante las ranuras de comando.

#### Corrección de errores

### Versión 1.3.0 — 2026-06-05
*Versión final.*

#### Nuevas funciones
* **Analizador de diseño de hoja:** Se añadió una potente infraestructura de escaneo de diseño. Detecta instantáneamente la protección de hojas, filtros de columnas activos, pestañas de hojas ocultas y bordes absolutos ocultos mientras almacena en caché los bloques de datos descubiertos.
* **Navegación guiada de bloques de datos:** La navegación posterior al análisis permite saltos inmediatos del cursor entre los principales grupos de datos, evitando celdas vacías de forma fluida.
* **Anunciador de formato condicional:** Detecta y lee automáticamente el color dinámico, el estilo de fuente y la sombra de fondo de las celdas alteradas por las reglas de formato condicional de Excel.
* **Aceleradores explícitos de configuración:** Se renovó completamente la GUI de configuración de BOA para cumplir estrictamente con la arquitectura de NVDA. Cada casilla de verificación de función ahora posee un atajo único global `Alt+Letra`, previniendo el cicleo de teclado y eliminando fallos de navegación por primera letra.

#### Corrección de errores
* **Detección de bordes absolutos:** Se reemplazaron las verificaciones de bordes COM nativas `UsedRange` con verificaciones de límites matemáticos absolutos 1D (`Fila 1048576` y `Columna 16384`) para garantizar la detección de filas/columnas ocultas incluso si se encuentran muy lejos del bloque de datos activo.
* **Interrupciones seguras de propiedades COM lentas:** Se reforzaron los bucles de propiedades COM para evitar congelamientos del hilo de NVDA al evaluar millones de estructuras ocultas contiguas.

### Versión 1.2.0 — 2026-06-03
*Versión final.*

#### Nuevas funciones
* **Caché al iniciar aplicación:** Gran renovación arquitectónica. Los módulos principales ahora se cargan de forma diferida exactamente cuando enfoca las aplicaciones de Office, eliminando el retraso de arranque, resolviendo completamente el fallo del objeto 'desconocido' en los diálogos de renombrado y preservando la estructura de código multifichero.
* **Rastreador de celdas mejorado (matemáticas COM 1D):** Se reescribió la lógica de detección de espacios en celdas ocultas para evaluar solo secciones transversales unidimensionales (`current_col` o `current_row`). Esto reduce la carga de cálculo COM en más de 16 millones de celdas, eliminando instantáneamente los congelamientos de navegación al saltar rangos ocultos.
* **Limpieza de memoria de proceso:** Se implementó el seguimiento del identificador de ventana de Excel (`Hwnd`) para detectar cuándo el usuario cierra y reabre Excel. Esto limpia activamente la memoria de estado global obsoleta y resuelve completamente el falso anuncio de "Hoja oculta" al abrir un nuevo "Libro1".

#### Corrección de errores
* **Doble anuncio de selección:** Se migró del uso no fiable de `winUser.getKeyState` asíncrono y se implementó `api.getLastInputGesture()` para suprimir perfectamente los dobles anuncios al usar las teclas Shift+Flecha.
* **Desactivación del detector de límites:** El detector proactivo de límites ha sido desactivado para proteger la estabilidad de la navegación nativa de NVDA, recurriendo completamente al rastreador con salto de espacios.

### Versión 1.1.0 — 2026-05-30
*Versión final.*

#### Nuevas funciones
* **GUI de configuración:** Se añadió un panel nativo BOA Office Enhancements dentro de `NVDA -> Preferencias -> Opciones` para activar o desactivar funciones fácilmente.
* **Gancho SafeRichEdit:** Previene cierres silenciosos de NVDA al interactuar con controles RichEdit en Office 2024.
* **Teclas rápidas personalizables:** Todas las teclas rápidas de BOA están ahora completamente expuestas en el diálogo de Gestos de Entrada de NVDA bajo la categoría "Better Office Accessibility".
* **Excel: Detección de salto de filas/columnas ocultas:** Anuncia proactivamente cuando navega más allá de filas o columnas ocultas, asegurando que nunca pierda datos filtrados. Se puede alternar en la configuración.

#### Corrección de errores
* **Seguridad de hilos:** Se eliminaron todos los retrasos bloqueantes (`time.sleep`) y se reemplazaron con callbacks asíncronos no bloqueantes de NVDA para garantizar que el lector de pantalla nunca tartamudee durante las operaciones en segundo plano.

### Versión 1.0.0 — 2026-05-24
*Versión pública inicial.*

#### Nuevas funciones
* **Excel: Organizador masivo de hojas:** Reordene instantáneamente múltiples hojas a la vez usando un diálogo completamente accesible.
* **Excel: Quick Sheet Mover:** Mueva la hoja activa a la izquierda, derecha, al inicio o al final mediante comandos de teclado.
* **Excel: Renombrado accesible de hojas:** Intercepta el campo de renombrado nativo inaccesible y lo reemplaza con un diálogo accesible fiable.
* **Excel: Seguimiento inteligente de selección:** Anuncia con precisión las selecciones y deselecciones de rangos de múltiples celdas.
* **PowerPoint: Selectores de color accesibles:** Permite a NVDA leer con precisión los valores RGB y Hex dentro del diálogo de color personalizado.
* **PowerPoint: Standard Color Grid Support:** Intercepta la navegación con teclas de flecha para leer los códigos Hex ocultos de la cuadrícula hexagonal de colores inaccesible.