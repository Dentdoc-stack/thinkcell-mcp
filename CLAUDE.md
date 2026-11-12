# Instrucciones para el Agente — think-cell MCP server

> Servidor MCP (Python, transporte stdio) que envuelve la automatización JSON
> `.ppttc` de think-cell: construye archivos `.ppttc` a partir de datos planos
> y los convierte a PowerPoint (`.pptx`) ejecutando `ppttc.exe`.

## Aprendizajes del Agente (Mejora Continua)

> **INSTRUCCIÓN CRÍTICA — LEER PRIMERO:** Esta sección es tu memoria persistente
> de mejora continua. **Con cada ciclo de ejecución** (al completar una tarea,
> resolver un error, descubrir un patrón, o ajustar un flujo) **y con cada
> actualización de cualquier Markdown** (CLAUDE.md, README, etc.) **debes
> agregar aquí un aprendizaje nuevo** si surgió algo no trivial. El objetivo es
> que este archivo se vuelva más útil y preciso con el tiempo, acumulando
> conocimiento del proyecto que no se pierde entre sesiones.
>
> **Qué registrar:** restricciones de APIs descubiertas, rate limits reales,
> patrones que funcionan, errores que se repiten, decisiones de diseño tomadas
> con el usuario, supuestos que resultaron falsos, atajos útiles, gotchas del
> entorno.
>
> **Qué NO registrar:** detalles efímeros de una sola tarea, información ya
> documentada en el README, cosas triviales derivables del código.
>
> **Formato de cada aprendizaje:**
> ```
> - **YYYY-MM-DD — [Tema corto]:** Descripción en 1-3 líneas. **Por qué
>   importa:** consecuencia práctica o cómo aplicarlo en el futuro.
> ```
>
> **Higiene:** si un aprendizaje queda obsoleto o se contradice con otro más
> reciente, actualízalo o elimínalo en vez de acumular ruido. Mantén la lista
> ordenada por fecha (más recientes arriba). Si superas ~25 entradas, consolida
> las más antiguas.

### Registro de aprendizajes

- **2026-11-12 — `UpdateBatch` actualiza el deck abierto en sitio:**
  `excel.COMAddIns("thinkcell.addin").Object.CreateUpdate()` →
  `AddRangeData(pres.Slides(n), nombre, rango, False)` → `Send()` rellena el
  gráfico nombrado dentro de la presentación abierta (probado sin ventana,
  think-cell 14). Necesita un Excel con el add-in; `update_open_charts` lanza
  uno propio y oculto con `DispatchEx` y lo cierra. Los nombres `c1`,
  `LeftChart`… sobreviven a ppttc e `InsertFromFile`. **Por qué importa:** es
  la única vía para refrescar un gráfico sin insertar un slide nuevo; los
  gráficos sin nombre no se pueden nombrar por API en el deck vivo.

- **2026-11-12 — "Use Datasheet Fill" = `<m_bExcelOnTop val="1"/>`:** el flag
  vive en el `CSequenceChartDataTable` del gráfico. En la plantilla oficial
  solo `RightChart` lo tiene; por eso `fill`/`semantic` no se ven en
  `left_chart` ni en las 34 plantillas de la biblioteca. Ponerlo a 1 en código
  NO basta (probado: el color queda en el datasheet pero no se dibuja); hay
  que marcarlo en PowerPoint. **Por qué importa:** `fill_template` y
  `create_auto_deck` solo avisan; para verificar colores hay que exportar el
  slide a PNG por COM, no contar `FF0000` en el XML.

- **2026-11-12 — `LoadStyle` funciona sin ventana:**
  `ppt.COMAddIns("thinkcell.addin").Object.LoadStyle(pres.SlideMaster, xml)`
  sobre la plantilla abierta con `WithWindow=False`, luego `Save()`; ppttc
  sigue rellenando después. Los campos de texto, tablas y Harvey balls con
  nombre NO están en el OLE `think-cellXML` (solo gráficos), así que
  `library.py` no puede nombrarlos. **Por qué importa:** el branding de
  esquemas de relleno va por archivo de estilo, no por el tema.

- **2026-11-12 — `os.replace` falla a ratos en Windows:** el antivirus abre
  el archivo destino y da `PermissionError`; `oauth.replace_file` reintenta
  (hasta 1 s). **Por qué importa:** sin el reintento el sign-in fallaba de
  forma intermitente al guardar `oauth_state.json`.

- **2026-11-11 — Los tools `def` bloquean el event loop (SDK 1.29):** FastMCP
  llama las funciones síncronas directamente en el loop, así que un
  `ppttc.exe` (hasta 180 s) congelaba todo el conector `--http`, incluido el
  sign-in. `_slow_tool` en `server.py` registra un wrapper `async` que corre
  la función en un hilo (`anyio.to_thread`) bajo un lock único y devuelve la
  función normal para llamadas directas (los tests la siguen llamando
  síncrona). `pythoncom` (structured storage y PowerPoint) funciona en esos
  hilos. **Por qué importa:** cualquier tool nuevo que lance ppttc o COM debe
  usar `@_slow_tool` en vez de `@mcp.tool()`.

- **2026-11-11 — Plantillas de usuario fuera de la carpeta de instalación:**
  Claude Desktop reemplaza la carpeta de la extensión en cada actualización.
  Con `THINKCELL_OUTPUT_DIR` definido, `library.USER_DIR` es
  `<output>/templates/user`; `BUNDLED_USER_DIR` (las que vienen en el paquete
  del equipo) se sigue leyendo y una guardada con el mismo nombre la
  reemplaza. La plantilla auto-deck con branding sigue en la carpeta de
  instalación (se pierde el branding al actualizar; se rehace con
  `set_deck_branding`). **Por qué importa:** el paquete del equipo incluye
  plantillas a propósito; no quitarlas del zip.

- **2026-11-08 — Plantillas de gráficos sin nombrar a mano (`library.py`):**
  los nombres de think-cell viven en el OLE `ppt/embeddings/oleObject*.bin`,
  stream `think-cellXML`, como `<m_strName>` dentro de cada
  `CSequenceChartSE|CPieChartSE|CScatterChartSE`; se reescriben con
  `pythoncom` (structured storage) y `ppttc.exe` los acepta. Las categorías se
  leen del datasheet (`m_dtable` → primer `<ocol>` → vectores); en archivos de
  think-cell las fuentes van por `idref`, pero tras pasar por ppttc quedan
  INLINE (hay que leer ambas). Los `.potx` se abren con python-pptx cambiando
  el content type en `[Content_Types].xml`. **Por qué importa:** cualquier
  slide (biblioteca instalada o deck del usuario) se vuelve plantilla sin
  abrir PowerPoint.

- **2026-11-08 — Fila "100%=" del datasheet:** si el gráfico la tiene (Mekko,
  algunos 100%), la segunda fila del JSON es la de totales; si falta, la
  primera serie desaparece, y si sobra en otro gráfico aparece una serie
  falsa. Se detecta cuando la fuente `m_varsrcAbsoluteExtent` del primer
  vector contiene `<m_precInitial` (celda formateada). Comparar total vs suma
  de columnas da falsos positivos (línea, combinación). Gantt (`CGanttSE`)
  queda fuera: ppttc rompe el eje de fechas. **Por qué importa:**
  `fill_template` añade el total de la plantilla solo; probarlo visualmente.

- **2026-11-08 — ngrok detrás de Norton Web Shield:** el agente ngrok falla
  con `x509: certificate signed by unknown authority`: Norton re-firma la
  conexión a `connect.ngrok-agent.com` con su "Untrusted Root" (ngrok usa CA
  propia), así que `agent.connect_cas: host` no sirve. Solo funciona una
  exclusión de `ngrok.exe` en el Web Shield (la pone el usuario). Además, el
  ngrok de la Microsoft Store guarda `ngrok.yml` virtualizado en
  `%LOCALAPPDATA%\Packages\ngrok.ngrok_*\LocalCache\Local\ngrok`, y el panel
  de terminal de Claude no permite pegar en `getpass`. **Por qué importa:**
  diagnosticar con el log antes de tocar configs; ofrecer
  `ngrok config add-authtoken` como alternativa al prompt oculto.

- **2026-11-07 — think-cell Assist como cliente MCP:** Assist conecta desde el
  navegador (origen `https://app.prod.ai.think-cell.com`) y solo admite
  cabeceras fijas (API key), no OAuth. `_browser_access` responde el preflight
  CORS antes de la autenticación (el navegador no manda credenciales en
  OPTIONS), devuelve 403 a otros `Origin` (la spec MCP lo exige) y 405 a
  GET/DELETE en `/mcp` (stateless). La API key entra como Bearer en
  `ConnectorAuth.load_access_token`. **Por qué importa:** sin preflight el
  indicador de Assist sale rojo; el formulario `/connect` postea desde el
  propio origen del servidor, que debe estar permitido. Con ngrok gratis hay
  que mandar `ngrok-skip-browser-warning`.

- **2026-11-07 — MCP 2026-07-28 (dual-era) sin migrar a mcp 2.x:** la spec
  nueva es stateless (sin `initialize` ni `Mcp-Session-Id`, versión en
  `params._meta`, `server/discover` obligatorio). El SDK 1.29 solo habla hasta
  2025-11-25, así que `modern_http.modern_era` intercepta en `/mcp` los POST
  con `_meta` moderno y reutiliza los handlers del SDK
  (`mcp._mcp_server.request_handlers[ListToolsRequest/CallToolRequest]`, que
  validan esquemas); lo demás pasa intacto al SDK (se re-inyecta el body).
  Orden: 401 → -32020 (cabeceras) → -32022 (versión) → 202 notificación →
  método (404 si no existe). stdio sigue solo-legacy. **Por qué importa:**
  mcp 2.x elimina `mcp.server.fastmcp`; este wrapper evita esa reescritura.
  Al migrar algún día, borrar `modern_http.py`.

- **2026-11-07 — Norton rompe respuestas HTTP locales con Content-Length:**
  Norton 360 (Web/Mail Shield) intercepta HTTP y HTTPS incluso en 127.0.0.1 y
  convierte toda respuesta con `Content-Length` a POST en una chunked sin fin;
  el túnel lo reporta como 520 ("unexpected EOF") en el 202 de
  `notifications/initialized`. Las respuestas que uvicorn ya envía chunked
  pasan intactas. Fix: `_chunked_only` quita `Content-Length` en modo `--http`.
  **Por qué importa:** sin él el conector se conecta pero falla en la segunda
  llamada; h11/uvicorn no tienen la culpa (h11 en proceso es correcto).

- **2026-11-07 — Conector remoto (`--http`) con OAuth:** los conectores de
  Claude (PowerPoint/claude.ai) se llaman desde la nube de Anthropic: hace
  falta HTTPS público (ngrok, dominio estático gratis) y no hay servicio
  hospedado posible (think-cell vive en el PC del usuario). La entrada es
  OAuth 2.1 del SDK (`oauth.py`, `ConnectorAuth`): Claude se registra solo,
  `/authorize` manda a `/connect` y el usuario escribe el *connect code*.
  Tokens guardados solo como hash. FastMCP con host 127.0.0.1 rechaza el Host
  del túnel: se desactiva la protección DNS-rebinding. **Por qué importa:**
  sin OAuth no aparece el botón Connect; en modo remoto toda ruta pasa por
  `_local_paths` (`output/`, `templates/`, carpetas de `connector.py setup`).

- **2026-11-07 — Extensión `.mcpb` de Claude Desktop:** `manifest.json` con
  `server.type "uv"`: Claude Desktop instala Python y las dependencias desde
  `pyproject.toml`, así que este debe listar TODAS (faltaban `openpyxl` y
  `pywin32`). `build_mcpb.py` empaqueta (zip) respetando `.mcpbignore`; no
  hace falta `npx @anthropic-ai/mcpb`. **Por qué importa:** si un módulo o
  dependencia falta en `pyproject.toml`, la extensión no arranca aunque
  `python server.py` funcione en local.

- **2026-11-07 — Descarga de ngrok bloqueada:** la URL antigua (equinox) da
  403 y la nueva (`bin.ngrok.com`) la corta el antivirus a mitad de descarga.
  `connector.py setup` intenta descargar y, si falla, explica cómo
  instalarlo a mano (Microsoft Store o ngrok.com). **Por qué importa:** no
  dar por hecho que la descarga automática funciona en PCs con Norton/Avast.

- **2026-11-07 — think-cell fuera de la ruta por defecto:** puede instalarse
  en cualquier carpeta. `find_thinkcell_dir()` lee
  `InstallLocation` de las claves `Uninstall` del registro (HKLM, WOW6432Node,
  HKCU) antes de caer al valor por defecto. **Por qué importa:**
  `THINKCELL_DIR` ya no es necesario en instalaciones normales.

- **2026-05-16 — `create_auto_deck` acepta ejes categóricos:** `_validate_chart`
  ya no exige fechas ISO. Si TODAS las categorías de un gráfico son
  `YYYY-MM-DD` → eje de fechas (celdas `date`); si no → eje categórico (celdas
  `string`). Se infiere por gráfico. **Por qué importa:** decks con ejes
  "Q1/Q2", regiones, segmentos, etc., no solo series temporales.
- **2026-05-15 — CI con tests «CI-safe»:** `tests.py` es el gate, pero
  `test_autodeck` necesitaba la plantilla de think-cell (ausente en CI). Fix:
  `build_auto_deck` valida la entrada ANTES de `ensure_auto_template()`, y
  `test_autodeck` salta solo los checks de construcción de deck si la
  plantilla no está disponible (los de validación corren siempre). Workflow
  en `.github/workflows/tests.yml` (Python 3.10 y 3.12). **Por qué importa:**
  la métrica del loop corre sola en cada push/PR.
- **2026-05-15 — `build_presentation` generaba un solo slide:** la librería
  `thinkcell` añade cada `add_chart`/`add_textfield` a `charts[-1]` (la última
  entrada de template). El código llamaba `add_template` una sola vez → todos
  los gráficos en una entrada → un solo slide. Fix: `write_ppttc_slides` llama
  `add_template` una vez por slide; cada slide es su propia entrada `.ppttc`.
  Los nombres pasan a ser únicos *por slide*, no globales. **Por qué importa:**
  decks reales de N slides; `write_ppttc_document` quedó como wrapper de 1
  slide sobre `write_ppttc_slides`.
- **2026-05-15 — Re-branding de la plantilla sin romper la automatización:** el
  logo de think-cell es una imagen `<p:pic>` *visible* en `slideMaster1.xml`;
  sus datos van en `<p:pic>` *ocultos* (`hidden="1"`, tamaño 0). Los colores de
  los gráficos salen de los 6 acentos de `theme1.xml`. Ambos son partes
  estándar de PowerPoint, separadas de los nombres think-cell (que viven en un
  blob LiteDB en los tags). **Por qué importa:** `branding.py` puede quitar el
  logo (borrar `<p:pic>` no ocultos del master) y recolorear (reescribir los
  acentos) sin tocar la automatización. El `fill` por serie sigue mandando.
- **2026-05-15 — Plantilla oficial sin-setup `thinkcell_auto.pptx`:** think-cell
  instala su propia plantilla de automatización con elementos ya nombrados
  (`SlideTitle`, `LeftChartTitle`, `RightChartTitle`, `LeftChart`,
  `RightChart`). Copiada al proyecto en `templates/thinkcell_auto.pptx` y usada
  por el tool `create_auto_deck`. **Por qué importa:** elimina el paso manual
  de abrir PowerPoint y nombrar gráficos; el usuario solo aporta datos.
- **2026-05-15 — Una entrada `.ppttc` = un slide:** cada objeto de nivel
  superior `{"template": ..., "data": [...]}` del array `.ppttc` genera UN
  slide. `build_presentation` mete todos los gráficos en una sola entrada (1
  slide); `create_auto_deck` emite una entrada por slide para decks reales de N
  slides. **Por qué importa:** para N slides hay que emitir N entradas, no
  apilar todo en `data`.
- **2026-05-15 — `validator.py` rechazaba celdas `percentage` válidas:**
  `CELL_TYPE_KEYS` omitía `"percentage"`, pero el esquema oficial
  (`ppttc-schema.json`) y `sample.ppttc` sí lo usan. Corregido: se añadió
  `percentage` a los tipos de celda y a la validación numérica. **Por qué
  importa:** `validate_ppttc` ya no marca como inválidos archivos correctos.
- **2026-05-15 — ppttc.exe sí existe:** think-cell 13.0.35.880 instala
  `C:\Program Files (x86)\think-cell\ppttc.exe` (más `ppttchdl.exe`,
  `tcdiag.exe`, etc.). La automatización `.ppttc` → `.pptx` por línea de
  comandos es real, no hace falta abrir el `.ppttc` a mano. **Por qué importa:**
  `convert_to_pptx` es funcional siempre que el gráfico del template esté
  nombrado.
- **2026-05-15 — Los nombres deben coincidir:** `ppttc.exe` corre con éxito
  (exit 0, genera el `.pptx`) aunque el `chart_name` del `.ppttc` no coincida
  con ningún elemento nombrado en el template; simplemente no rellena nada.
  **Por qué importa:** un "éxito" silencioso no garantiza datos. Verifica que
  el template tenga cada gráfico nombrado (mini-barra de think-cell → campo
  "Name").

---

## Qué es este proyecto

Servidor MCP que expone **15 herramientas** (stdio, `--http` para el conector remoto); las principales:

| Herramienta | Propósito |
| --- | --- |
| `create_chart` | Construir un `.ppttc` para un solo gráfico. |
| `build_presentation` | Combinar varios gráficos/slides en un `.ppttc`. |
| `create_auto_deck` | Construir un deck multi-slide SIN preparar template. |
| `set_deck_branding` | Recolorear la plantilla auto-deck y quitar el logo. |
| `convert_to_pptx` | Ejecutar `ppttc.exe` para producir el `.pptx`. |
| `validate_ppttc` | Validar estructuralmente un documento `.ppttc`. |
| `list_chart_types` | Describir cada tipo de gráfico soportado. |
| `diagnose_thinkcell` | Diagnosticar por qué la automatización no funciona. |
| `list_chart_templates` / `fill_chart_template` | Biblioteca de gráficos de think-cell (34 plantillas) y slides propios guardados. |
| `template_from_slide` | Guardar un slide propio como plantilla (nombra los gráficos en código). |

8 tipos de gráfico: `waterfall`, `bar`, `stacked_bar`, `line`, `scatter`,
`mekko`, `area`, `combo`.

## Arquitectura del código

El proyecto separa responsabilidades en módulos deterministas y testeables —
la lógica de negocio no es probabilística y debe ser consistente:

- **`server.py`** — punto de entrada MCP (FastMCP, stdio). Solo orquesta: valida
  argumentos, llama a los módulos y devuelve dicts estructurados.
- **`autodeck.py`** — construcción del `.ppttc` para `create_auto_deck`. Arma
  el JSON directamente (no usa la librería `thinkcell`, que no expone celdas
  `date`/`percentage`/`fill`) y usa la plantilla oficial incluida
  `templates/thinkcell_auto.pptx`, resuelta de forma absoluta.
- **`branding.py`** — re-tematiza la plantilla auto-deck (`set_deck_branding`):
  reescribe los 6 colores de acento del tema y quita el logo de think-cell del
  slide master. Solo toca partes estándar de PowerPoint; no afecta los
  elementos nombrados de think-cell.
- **`charts/`** — un módulo por tipo de gráfico. Cada uno subclasea
  `ChartBuilder` (`charts/base.py`): valida la entrada para ese tipo y produce
  el par `(categories, series)` que espera la librería `thinkcell`.
- **`converter.py`** — wrapper de `ppttc.exe` con manejo de errores
  estructurado; nunca lanza excepción para los fallos documentados.
- **`validator.py`** — validación estructural del esquema `.ppttc` (JSON).
- **`diagnostics.py`** — chequeos **solo-lectura** de registro/archivos de
  Windows para diagnosticar el entorno think-cell. No lanza procesos.

**Cómo funciona la automatización think-cell:** think-cell **no** crea gráficos
de la nada y no expone una API general para "controlar" su add-in. El único
flujo soportado es: un template `.pptx` que ya contiene gráficos think-cell
**nombrados** → un `.ppttc` (JSON) que mapea cada nombre a una tabla de datos →
`ppttc.exe` empuja los datos en los elementos que coinciden por nombre. Si los
nombres no coinciden, `ppttc.exe` corre pero no actualiza nada.

## Principios de operación

1. **Revisa primero si existe la herramienta.** Antes de escribir código nuevo,
   revisa `charts/`, `server.py` y los módulos existentes.
2. **Auto-corrección cuando algo falla.** Lee el error y el stack trace, corrige,
   y vuelve a probar. No ejecutes `convert_to_pptx` repetidamente sin entender
   el fallo: usa `diagnose_thinkcell` primero.
3. **Prueba antes de dar por terminado.** Ejecuta `python tests.py` tras
   cualquier cambio en `charts/`, `validator.py` o `server.py`.
4. **Actualiza la documentación a medida que aprendes.** README y esta sección
   de aprendizajes son documentos vivos. No los descartes; mejóralos.

## Loop de mejora continua (patrón autoresearch)

Inspirado en [`karpathy/autoresearch`](https://github.com/karpathy/autoresearch):
el proyecto no avanza con mejoras sueltas, sino con un **loop explícito** —
cada vuelta acotada, pequeña y verificada por una métrica única.

**La métrica.** `python tests.py` debe dar **100%**. Es el único *gate*: si no
está verde, el repo no avanza. Equivale al `val_bpb` de autoresearch — una
señal objetiva, reproducible, de una sola orden.

**El loop — cada vuelta:**

1. Elige UN ítem del Backlog (mayor valor / menor riesgo primero).
2. Impleméntalo con un cambio pequeño y acotado.
3. Añade o actualiza tests en `tests.py` que cubran el cambio.
4. Corre `python tests.py`. Si no es 100%: corrige o revierte. **Nunca dejes
   el repo en rojo.**
5. Registra una entrada en «Registro de aprendizajes».
6. Haz un commit pequeño y enfocado; marca el ítem del Backlog como hecho.
7. Repite.

**Invariantes — el loop NUNCA debe romperlos** (equivalen al `prepare.py` que
autoresearch prohíbe tocar):

- Los 15 tools MCP siguen registrados e importables (`test_server_smoke`).
- No se redistribuyen archivos de think-cell: la plantilla se copia del
  install local.
- Toda herramienta devuelve dicts estructurados; nunca lanza excepción para
  fallos documentados.
- Sin rutas ni datos personales en el repo.

### Backlog de mejoras

> Cola hacia adelante. Saca de aquí en el paso 1; añade ideas nuevas arriba.

- [x] ~~CI: GitHub Actions que corra `python tests.py` en cada push/PR.~~
      Hecho 2026-05-15: `.github/workflows/tests.yml`; tests hechos CI-safe.
- [x] ~~`build_presentation` mete todos los gráficos en una sola entrada
      `.ppttc` → un solo slide.~~ Hecho 2026-05-15: `write_ppttc_slides` emite
      una entrada por slide.
- [x] ~~`create_auto_deck`: permitir ejes categóricos, no solo fechas ISO.~~
      Hecho 2026-05-16: `_validate_chart` infiere eje de fecha vs categórico.
- [x] ~~`CHANGELOG.md` + versionado semántico.~~ Hecho 2026-05-16:
      `CHANGELOG.md` (formato Keep a Changelog), v0.1.0.
- [x] ~~CI: actualizar `actions/checkout` y `actions/setup-python` a v6~~
      (cierra el aviso de Node 20). Hecho 2026-05-16.
- [ ] README: GIF o capturas de un deck generado. El *quickstart* en texto ya
      está; el GIF/capturas requieren un humano con un deck renderizado.

## Organización de archivos

```
Thinkcell/
  server.py          15 herramientas MCP (FastMCP, stdio / --http)
  library.py         biblioteca de plantillas think-cell (cortar, nombrar, rellenar)
  autodeck.py        constructor de .ppttc sin-setup (create_auto_deck)
  branding.py        re-tematiza la plantilla auto-deck (set_deck_branding)
  converter.py       wrapper de ppttc.exe
  validator.py       validación estructural de .ppttc
  diagnostics.py     diagnóstico del entorno think-cell
  tests.py           suite de pruebas (python tests.py)
  charts/            base.py + 8 módulos (uno por tipo) + __init__.py
  templates/         thinkcell_auto.pptx (plantilla oficial pre-nombrada)
  output/            .ppttc / .pptx generados (regenerables, no versionar)
  template.pptx      template de PowerPoint con gráficos think-cell nombrados
  requirements.txt   dependencias (mcp[cli], thinkcell)
  README.md          documentación completa
```

**Principio clave:** los archivos en `output/` siempre pueden borrarse y deben
ser reproducibles ejecutando las herramientas de nuevo, nunca editados a mano.

## Entorno

- **Python:** 3.10 o superior.
- **ppttc.exe:** dentro del directorio de think-cell (por defecto
  `C:\Program Files (x86)\think-cell`). Si think-cell está instalado en otra
  ruta, define la variable de entorno `THINKCELL_DIR`; `converter.py`,
  `diagnostics.py` y `autodeck.py` la respetan.
- **Requiere Windows** con think-cell y PowerPoint instalados para
  `convert_to_pptx` y `diagnose_thinkcell`. La construcción y validación de
  `.ppttc` funciona en cualquier sistema operativo.
- **Plantilla de `create_auto_deck`:** no se versiona. `autodeck.py` la copia
  desde la instalación local de think-cell (`<THINKCELL_DIR>/ppttc/template.pptx`)
  la primera vez que se usa; no se redistribuye el archivo de think-cell.
