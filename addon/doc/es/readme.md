# BOA: Better Office Accessibility

BOA es una potente suite de mejoras de accesibilidad para Microsoft Office, diseñada para mejorar enormemente la experiencia del lector de pantalla para usuarios de NVDA. Parchea directamente componentes de interfaz de usuario inaccesibles e introduce herramientas de navegación rápida para Excel y PowerPoint.

---

## ⌨️ Referencia de teclas de acceso rápido

| Característica | Combinación de teclas | Contexto / Notas |
| :--- | :--- | :--- |
| **Entrar al Modo de comando** | `NVDA+E` | Activa el modo de prefijo de comando (activa un pitido agudo) |
| **Analizar diseño de hoja** | `NVDA+E`, luego `L` | Ejecutar dentro de Excel antes de navegar por bloques de datos |
| **Saltar al bloque de datos más cercano** | `NVDA+E`, luego `J` /  | Requiere un análisis de diseño primero |
| **Abrir organizador de hojas en masa** | `NVDA+E`, luego `X` | Abre el cuadro de diálogo de reordenación de hoja accesible |
| **Mover hoja activa a la izquierda** | `NVDA+Shift+LeftArrow` | Desplaza la hoja activa una posición hacia arriba|
| **Mover hoja activa a la derecha** | `NVDA+Shift+RightArrow` | Desplaza la hoja de cálculo activa una posición hacia abajo|
| **Mover hoja al inicio/fin** | `NVDA+Shift+Home` / `End` | Envía la hoja de cálculo a los límites absolutos |
| **Formato condicional detallado**| `NVDA+E`, luego `F` | Anuncia los detalles completos de formato de la celda enfocada |
| **Asignar celda a espacio de memoria** | `NVDA+E`, luego `Shift+1` a `Shift+9` | Asigna la celda actual a un espacio de monitor en segundo plano |
| **Leer espacio de celda monitoreada** | `NVDA+E`, luego `1` a `9` | Recuerda y lee el valor del espacio asignado |
| **Alternar monitoreo en segundo plano** | `NVDA+E`, luego `M` | Alterna manualmente el seguimiento de cálculos en segundo plano |
| **Borrar todos los espacios de memoria** | `NVDA+E`, luego `Backspace` | Purgar todos los monitores de celda en segundo plano guardados |
| **Cancelar modo de comando** | `Escape` | Sale del modo de prefijo de comando |

---

## 🚀 Características

### Mejoras de Excel

#### 1. Analizador de diseño de hoja y almacenamiento en caché
Escanee instantáneamente cualquier hoja de cálculo de Excel para comprender su estructura, elementos ocultos y bloques de datos.
* **Cómo funciona:** BOA escanea rápidamente la hoja y anuncia los bloques de datos activos. También le advierte sobre **pestañas de hoja de cálculo ocultas**, **filtros** activos, **modos protegidos** y **límites exteriores ocultos** (por ejemplo, si las columnas cerca del borde derecho de la hoja están ocultas, evitando que se pierda los datos fuera de la pantalla).
* **Navegación de datos:** Después de escanear, puede usar las teclas de acceso rápido de salto de bloque de datos para teletransportar instantáneamente su cursor entre bloques de datos descubiertos, omitiendo sin esfuerzo miles de celdas vacías.

#### 2. Organizador de hojas en masa
Reordene y organice instantáneamente varias hojas a la vez mediante un cuadro de diálogo totalmente accesible .
* **Cómo funciona:** Abre un cuadro de diálogo donde puede seleccionar una hoja y asignarla a una nueva posición. Los movimientos programados se enumeran en una tabla de datos (presione `Del` para eliminar un error). Haga clic en `OK` y su libro de trabajo se reorganizará instantáneamente.

#### 3. Movimiento rápido de hoja
Mueva la hoja activa a la izquierda, a la derecha, al principio o al final instantáneamente usando los atajos de teclado.

#### 4. Renombrado de hoja accesible
* Al renombrar una hoja, NVDA tiene dificultades de forma nativa para leer los caracteres que está escribiendo.
* BOA inyecta una clase `ExcelSheetRenameEdit` personalizada que utiliza el motor `SafeRichEdit`, lo que significa que puede leer con precisión por carácter, palabra o línea mientras cambia el nombre. Esto sirve como una mejora al comportamiento de renombrado predeterminado existente.

#### 5. Rastreador de fila/columna oculta
* Realiza un seguimiento proactivo de su movimiento a través de la cuadrícula para evitar que se pierda datos ocultos o filtrados.
* **Celdas fragmentadas cruzadas:** Si salta a través de una sección muy fragmentada u oculta de la cuadrícula (por ejemplo, moviéndose de la Fila 3 a la Fila 10 porque las Filas 4–9 están ocultas), BOA anuncia explícitamente "Filas 4 hasta la 9 ocultas". Esto asegura que siempre sepa cuándo se han omitido datos en la estructura.

#### 6. Anunciador de formato condicional
* Lee automáticamente el color, el estilo de fuente y el tono de fondo de las celdas que han sido cambiadas dinámicamente por las reglas de Formato Condicional de Excel.
* Le da el verdadero estado visual de la celda en lugar de solo el valor subyacente sin procesar. Inicialmente, al enfocar la celda, anuncia "tiene formato condicional y algunos otros detalles menores". Para obtener información completa, use la configuración detallada de teclas de acceso rápido, que es NVDA E y F.

#### 7. Mejor anuncio de selección
lee si la celda o rango están seleccionados o deseleccionados.

#### 8 Monitor de celda:
* **Monitor de celda:** Use rutas de comandos para asignar celdas específicas a espacios de memoria. Puede volver atrás y leerlas en cualquier momento usando el espacio numérico asignado.
* **Monitoreo continuo:** Las celdas asignadas son monitoreadas automáticamente en segundo plano. Si Excel desencadena un recálculo o una edición de celda, BOA anuncia instantáneamente el nuevo valor. Alterne manualmente o borre todo a través de los espacios de comando.

### Mejoras de PowerPoint

#### 1. Selectores de color accesibles
* Desbloquea el cuadro de diálogo de Color personalizado en PowerPoint.
* Identifica y lee explícitamente los cuadros de edición de "Rojo", "Verde" y "Azul" de forma correcta (al anular `PowerPointRGBEdit`).
* Asigna el campo de entrada Hex anteriormente invisible para que NVDA pueda leer el valor de color Hex completo limpiamente.

#### 2. Soporte de cuadrícula de color estándar
* Navegar por la cuadrícula hexagonal de color "Estándar" de PowerPoint normalmente se lee como "Gráfico" o silencio.
* BOA rastrea las teclas de flecha a través del hexágono y obtiene en silencio el valor de color oculto, anunciándolo en tiempo real (por ejemplo, "Color #FF0000").

### Infraestructura y Mecanismos Técnicos

#### El modo de prefijo de comando
Para evitar conflictos de pulsaciones de teclas con otros complementos de NVDA, BOA utiliza un **Modo de prefijo de comando**:
1. Presione la tecla de acceso directo de activación para ingresar al Modo de comando. Escuchará un pitido agudo.
2. Presione una tecla secundaria para activar una característica específica.
3. Si presiona una tecla no válida, escuchará un pitido de error.

#### Panel de personalización y configuración
* Las características de BOA son completamente modulares y se pueden habilitar o deshabilitar en cualquier momento. Vaya a `Menú NVDA -> Preferencias -> Opciones -> Mejoras de BOA Office` para alternar características individuales.
* **Teclas aceleradoras inteligentes:** Cada configuración presenta un atajo acelerador único matemáticamente `Alt+Tecla` dentro del panel. Por ejemplo, presione `Alt+E` para saltar instantáneamente al grupo de Excel, `Alt+P` para PowerPoint y `Alt+W` para Word.
* La configuración se guarda de forma segura en un archivo JSON independiente (`boa_settings.json`), asegurando que la configuración principal de NVDA nunca se corrompa.
* Si Microsoft Office repara oficialmente un error de accesibilidad en el futuro, puede deshabilitar de forma segura el gancho de anulación específico de BOA sin perder el resto de la funcionalidad del complemento.

#### Seguridad y Límites de Integración
* Las inyecciones del portapapeles verifican estrictamente los ID de procesos de primer plano de la ventana para evitar la fuga de datos a otras aplicaciones.
* algunas teclas de acceso rápido personalizadas están completamente expuestas en el cuadro de diálogo Gestos de entrada de NVDA bajo la categoría "Mejoras de BOA Office" (Better Office Accessibility).

---

## 📋 Requisitos

* **NVDA:** Versión 2026.1.0 o posterior.
* **Aplicaciones:** Microsoft Excel y Microsoft PowerPoint.

---

## 💾 Instalación

1. Descargue el archivo de lanzamiento `.nvda-addon` más reciente o búsquelo dentro de la Tienda de complementos nativa de NVDA.
2. si instala desde un archivo, abra el archivo o use la `Tienda de complementos de NVDA -> Instalar desde archivo externo`.
3. Reinicie NVDA.

---

## 🛠️ Registro de cambios

### v1.6.0
* **Soporte integral de traducción**: El complemento ahora está completamente localizado con soporte para 17 idiomas globales. 
  *(Nota: Estas traducciones fueron generadas por IA, por lo que pueden estar presentes algunos errores o inexactitudes menores de traducción.)*
* **Gobernanza estricta del código**: Se aplicaron encabezados de derechos de autor GPL-2.0 en toda la base de código.

### Versión 1.5.0 
#### Nuevas características
##### Radar de fin de datos
Al navegar a través de hojas de cálculo grandes, puede ser difícil saber si una celda vacía significa que ha llegado al final de una lista, o si simplemente hay un hueco en los datos. El **Radar de fin de datos** actúa como una verificación de perímetro inteligente para evitar que se mueva a ciegas a través del espacio vacío.
Siempre que navega hacia una celda vacía, BOA escanea instantáneamente las celdas restantes en su dirección de viaje. Si no hay absolutamente ningún dato restante, anunciará de manera proactiva:
* *"No hay más datos abajo"*
* *"No hay más datos arriba"*
* *"No hay más datos a la derecha"*
* *"No hay más datos a la izquierda"*
**Opciones de configuración:**
Puede configurar esta característica a través de `Opciones de NVDA -> Opciones -> Mejoras de BOA Office`. Debido a que las hojas de cálculo pueden contener complejidades ocultas (como fórmulas invisibles o filas contraídas), el radar proporciona tres modos de funcionamiento:
1. **Desactivado**: Deshabilita el radar por completo.
2. **Comprobación estricta de memoria (CountA) [Predeterminado]**: El enfoque más seguro y rápido. Comprueba la memoria sin procesar de la hoja de cálculo. Si detecta *algo* debajo de usted (incluidas filas ocultas, texto, números o fórmulas invisibles), permanece completamente en silencio para evitar falsas alarmas. Solo anuncia "No hay más datos" cuando el resto de la hoja está 100% matemáticamente en blanco.
3. **Solo datos visibles (Motor matemático)**: Un motor altamente avanzado diseñado para hojas complejas. Filtra de manera inteligente filas ocultas y fórmulas invisibles (por ejemplo, `=""`). Solo permanecerá en silencio si quedan números o texto reales y visibles en su camino.

### Versión 1.4 - 2026-06-12
#### Nuevas características
* **Monitor de celda:** Use rutas de comandos para asignar celdas específicas a espacios de memoria. Puede volver y leerlas en cualquier momento usando el espacio numérico asignado.
* **Monitoreo continuo:** Las celdas asignadas se monitorean automáticamente en segundo plano. Si Excel desencadena un recálculo o edición de celda, BOA anuncia instantáneamente el nuevo valor. Alterne manualmente o borre todo a través de espacios de comando.

#### Corrección de errores

### Versión 1.3.0 — 2026-06-05
*Versión final.*

#### Nuevas características
* **Analizador de diseño de hoja:** Se agregó una potente infraestructura de escaneo de diseño. Detecta instantáneamente Protección de hoja de cálculo, Filtros de columna activos, Pestañas de hoja de cálculo ocultas y bordes absolutos ocultos mientras almacena en caché los bloques de datos descubiertos.
* **Navegación guiada por bloques de datos:** La navegación posterior al análisis permite deformaciones inmediatas del cursor entre grupos importantes de datos, omitiendo celdas vacías a la perfección.
* **Anunciador de formato condicional:** Detecta y lee automáticamente el color dinámico, el estilo de fuente y el tono de fondo de las celdas alteradas por las reglas de Formato Condicional de Excel.
* **Aceleradores de configuración explícitos:** Se revisó por completo la GUI de Configuración de BOA para cumplir estrictamente con la arquitectura de NVDA. Cada casilla de verificación de características ahora posee un atajo de `Alt+Letra` único globalmente, lo que evita el ciclo del teclado y elimina las fallas de navegación de la primera letra.

#### Corrección de errores
* **Detección de límite de borde absoluto:** Se reemplazaron las verificaciones de borde de `UsedRange` COM nativas con verificaciones de límites matemáticos 1D absolutos (`Fila 1048576` y `Columna 16384`) para garantizar la detección de filas/columnas ocultas incluso si se encuentran muy fuera del bloque de datos activo.
* **Salvavidas seguros de propiedad COM perezosa:** Se endurecieron los bucles de propiedad COM para evitar bloqueos del hilo de NVDA al evaluar millones de estructuras ocultas contiguas.

### Versión 1.2.0 — 2026-06-03
*Versión final.*

#### Nuevas características
* **Almacenamiento en caché de inicio de aplicación:** Importante revisión de la arquitectura. Los módulos principales ahora se cargan de forma diferida exactamente cuando te enfocas en las aplicaciones de Office, eliminando el retraso de inicio, resolviendo por completo el error de enfoque del objeto 'desconocido' en los cuadros de diálogo de cambio de nombre y preservando la estructura del código de múltiples archivos.
* **Rastreador de celdas mejorado (Matemáticas COM 1D):** Se reescribió la lógica de detección de huecos de celdas ocultas para evaluar solo secciones transversales unidimensionales (`current_col` o `current_row`). Esto reduce la carga de cálculo de COM en más de 16 millones de celdas, eliminando instantáneamente los congelamientos de navegación al saltar rangos ocultos.
* **Borrado de memoria de proceso:** Se implementó el seguimiento del identificador de ventana de Excel (`Hwnd`) para detectar cuándo el usuario cierra y vuelve a abrir Excel. Esto borra activamente la memoria de estado global inactiva y resuelve por completo el falso anuncio de "Hoja oculta" al abrir un nuevo "Book1".

#### Corrección de errores
* **Anuncio de doble selección:** Se migró de la poco confiable `winUser.getKeyState` asincrónica y se implementó `api.getLastInputGesture()` para suprimir perfectamente los anuncios dobles al usar las teclas Shift+Flecha.
* **Desactivación del detector de límites:** El detector de límites proactivo se ha desactivado para proteger la estabilidad de la navegación nativa de NVDA, recayendo por completo en el rastreador de salto de brechas.

### Versión 1.1.0 — 2026-05-30
*Versión final.*

#### Nuevas características
* **GUI de configuración:** Se agregó un panel nativo de Mejoras de BOA Office dentro de `NVDA -> Preferencias -> Opciones` para activar o desactivar fácilmente las características.
* **Gancho SafeRichEdit:** Evita bloqueos silenciosos de NVDA al interactuar con los controles RichEdit en Office 2024.
* **Teclas de acceso rápido personalizables:** Todas las teclas de acceso rápido de BOA ahora están completamente expuestas en el cuadro de diálogo Gestos de entrada de NVDA bajo la categoría "Mejoras de BOA Office" (Better Office Accessibility).
* **Excel: Detección de omisión de fila/columna oculta:** Anuncia proactivamente al navegar más allá de filas o columnas ocultas, asegurando que nunca se pierda los datos filtrados. Se puede alternar en la configuración.

#### Corrección de errores
* **Seguridad de subprocesos:** Se eliminaron todos los retrasos de bloqueo (`time.sleep`) y se reemplazaron por devoluciones de llamada asincrónicas de NVDA sin bloqueo para garantizar que el lector de pantalla nunca tartamudee durante las operaciones en segundo plano.

### Versión 1.0.0 — 2026-05-24
*Versión pública inicial.*

#### Nuevas características
* **Excel: Organizador de hojas en masa:** Reordene instantáneamente varias hojas a la vez usando un cuadro de diálogo totalmente accesible.
* **Excel: Movimiento rápido de hoja:** Mueva la hoja activa a la izquierda, a la derecha, al inicio o al final a través de comandos de teclado.
* **Excel: Renombrado de hoja accesible:** Intercepta el campo de cambio de nombre nativo inaccesible y lo reemplaza con un cuadro de diálogo accesible y confiable.
* **Excel: Seguimiento de selección inteligente:** Anuncia con precisión selecciones y deselecciones de rango de celdas múltiples.
* **PowerPoint: Selectores de color accesibles:** Permite a NVDA leer con precisión valores RGB y Hex dentro del cuadro de diálogo Color personalizado.
* **PowerPoint: Soporte de cuadrícula de color estándar:** Intercepta la navegación de la tecla de flecha para leer códigos hexadecimales ocultos de la cuadrícula hexagonal de color inaccesible.
