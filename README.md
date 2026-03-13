# Outlook Email Analyzer

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)](https://python.org)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)](https://github.com)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)
[![Microsoft Graph](https://img.shields.io/badge/API-Microsoft%20Graph%20v1.0-0078D4?logo=microsoft)](https://learn.microsoft.com/en-us/graph/overview)
[![Auth](https://img.shields.io/badge/Auth-OAuth%202.0%20Device%20Code-orange)](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code)
[![Security](https://img.shields.io/badge/Token%20Cache-DPAPI%20Encrypted-critical)](https://learn.microsoft.com/en-us/dotnet/standard/security/how-to-use-data-protection)

Herramienta de linea de comandos que analiza correos de cuentas Microsoft personales
(Hotmail, Outlook.com, Live.com) y muestra un ranking de remitentes ordenado por cantidad,
de mayor a menor. Identifica rapidamente newsletters, notificaciones automaticas y
remitentes frecuentes para facilitar la limpieza del buzon.

---

## Tabla de contenidos

- [Problema que resuelve](#problema-que-resuelve)
- [Como funciona](#como-funciona)
- [Arquitectura](#arquitectura)
- [Requisitos](#requisitos)
- [Instalacion](#instalacion)
- [Configuracion inicial (1 sola vez)](#configuracion-inicial-1-sola-vez)
- [Uso](#uso)
- [Opciones de linea de comandos](#opciones-de-linea-de-comandos)
- [Ejemplo de salida](#ejemplo-de-salida)
- [Seguridad](#seguridad)
- [Estructura del proyecto](#estructura-del-proyecto)
- [Solucion de problemas](#solucion-de-problemas)
- [Licencia](#licencia)

---

## Problema que resuelve

Cuando un buzon acumula miles de correos (newsletters de portales de empleo, notificaciones
automaticas, suscripciones olvidadas), Outlook no ofrece una vista nativa que agrupe los
correos por remitente y los ordene por cantidad. Esta herramienta genera exactamente eso:
un ranking tabular que permite identificar de un vistazo quien envia mas correo.

---

## Como funciona

1. Te autenticas con tu cuenta Microsoft via OAuth2 (Device Code Flow).
   Solo necesitas hacerlo una vez; el token se cachea de forma segura.
2. La app consulta Microsoft Graph API pidiendo unicamente los campos `from` e `isRead`
   de cada correo (no descarga el cuerpo ni adjuntos).
3. Agrupa por direccion de remitente, cuenta y ordena de mayor a menor.
4. Muestra el ranking en la terminal con barra de distribucion visual.
5. **Modo newsletters (`--newsletters`):** Arquitectura de dos pases:
   - **Pase 1:** Lee los headers `List-Unsubscribe` (RFC 2369) de cada correo para detectar remitentes newsletter.
   - **Pase 2:** Para cada remitente unico, descarga el body HTML de 1 solo correo y extrae el link real de desuscripcion (busca URLs con palabras clave como "unsubscribe", "optout", "desuscri", etc.). Esto genera links que funcionan directamente en el navegador.

---

## Arquitectura

```
Tu cuenta Microsoft (Hotmail / Outlook.com / Live.com)
        |
        v
Microsoft Graph REST API  (https://graph.microsoft.com/v1.0)
        |                   GET /me/mailFolders/{folder}/messages
        |                   $select=from,isRead  (solo headers, no body)
        |                   Paginacion automatica (999 msgs/request)
        v
+------------------+     +----------------+     +---------------+
|   ms_graph.py    | --> |  analyzer.py   | --> |  display.py   |
|                  |     |                |     |               |
| - OAuth2 MSAL   |     | - Counter por  |     | - Tabla       |
| - Device Code   |     |   email        |     |   formateada  |
| - Token cache   |     | - Ranking desc |     | - Barra ASCII |
|   (DPAPI)       |     | - Nombre mas   |     | - Resumen     |
| - Graph API     |     |   frecuente    |     |   estadistico |
| - Retry/timeout |     | - Newsletter   |     | - Tabla news- |
| - Newsletters:  |     |   detection    |     |   letters con |
|   2do pase HTML |     +----------------+     |   link unsub  |
|   body parsing  |                            +---------------+
+------------------+
        |
        v
  ~/.outlook-email-analyzer/
    config.json        (client_id - cifrado o permisos restrictivos)
    token_cache.json   (tokens OAuth - cifrado DPAPI en Windows)
```

**Flujo de datos:** Microsoft Graph API --> ms_graph.py (MailData) --> analyzer.py (SenderStats) --> display.py (stdout) o main.py (CSV)

---

## Requisitos

- Python 3.10 o superior
- Una cuenta Microsoft personal (Hotmail, Outlook.com, Live.com)
- Conexion a internet
- Una app registrada en Azure AD (gratuito, se configura en 2 minutos)

---

## Instalacion

```bash
git clone https://github.com/tu-usuario/outlook-email-analyzer.git
cd outlook-email-analyzer

pip install -r requirements.txt
```

Dependencias:

| Paquete | Proposito |
|---------|-----------|
| `msal` | Autenticacion OAuth2 con Microsoft (MSAL) |
| `msal-extensions` | Cifrado DPAPI del token cache en Windows |
| `requests` | Peticiones HTTP a Microsoft Graph API |

---

## Configuracion inicial (1 sola vez)

La primera ejecucion te guia paso a paso. Necesitas registrar una app en Azure:

1. Abre https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade
2. Inicia sesion con tu cuenta Microsoft
3. Click "Nuevo registro":
   - Nombre: `EmailAnalyzer`
   - Tipos de cuenta: "Cuentas en cualquier directorio organizacional y cuentas personales de Microsoft"
   - URI de redireccion: Plataforma "Cliente publico/nativo", URI: `https://login.microsoftonline.com/common/oauth2/nativeclient`
4. Click "Registrar" y copia el "Id. de aplicacion (cliente)"
5. Ve a "Autenticacion" > "Configuracion avanzada" > activa "Permitir flujos de clientes publicos"
6. Ve a "Permisos de API" > "Agregar un permiso" > Microsoft Graph > Delegados > `Mail.Read`
7. Ejecuta `python main.py` y pega el Client ID cuando se te pida

A partir de ahi, el Client ID se guarda localmente y no se vuelve a pedir.

---

## Uso

```bash
# Analizar toda la bandeja de entrada
python main.py

# Solo correos no leidos
python main.py --unread-only

# Top 20 remitentes
python main.py --top 20

# Combinar opciones
python main.py --unread-only --top 20

# Ver carpetas disponibles
python main.py --list-folders

# Analizar otra carpeta
python main.py --folder sentitems

# Detectar newsletters y mostrar links de desuscripcion
python main.py --newsletters

# Newsletters: solo top 10
python main.py --newsletters --top 10

# Newsletters: exportar a CSV con columna de link
python main.py --newsletters --export newsletters.csv

# Exportar a CSV
python main.py --export ranking.csv

# Especificar email sin que lo pregunte
python main.py --user tu-correo@hotmail.com

# Borrar configuracion OAuth y reconfigurar
python main.py --reset-auth

# Ver ayuda completa
python main.py --help
```

---

## Opciones de linea de comandos

| Opcion | Descripcion | Valor por defecto |
|--------|-------------|-------------------|
| `--user EMAIL` | Direccion de correo (si no se da, se pide interactivamente) | Interactivo |
| `--unread-only` | Analizar solo correos no leidos | Todos |
| `--newsletters` | Detectar newsletters y mostrar link de desuscripcion por remitente | Desactivado |
| `--top N` | Mostrar solo los top N remitentes | Sin limite |
| `--folder NOMBRE` | Carpeta a analizar (ver `--list-folders`) | `inbox` |
| `--list-folders` | Listar carpetas disponibles con conteo y salir | -- |
| `--export ARCHIVO.csv` | Exportar resultado a CSV | -- |
| `--reset-auth` | Borrar tokens y configuracion OAuth guardados | -- |
| `--help` | Mostrar ayuda y salir | -- |

Carpetas conocidas de Microsoft Graph: `inbox`, `sentitems`, `drafts`, `deleteditems`, `junkemail`.

---

## Ejemplo de salida

```
===========================================================================
  OUTLOOK EMAIL ANALYZER -- Ranking de remitentes
===========================================================================

  Carpeta: Bandeja de entrada | 3,700 correos (3,200 no leidos)
  Modo: todos -- analizando ~3,700 correos...

  Procesando correo 3,700/3,700... 100%

---------------------------------------------------------------------------
    #   Remitente                       Email                                Cantidad
---------------------------------------------------------------------------
    1   Computrabajo                    alertas@computrabajo.com                  487
    2   LinkedIn Jobs                   jobs@linkedin.com                         312
    3   Magneto Empleos                 noreply@magneto365.com                    289
    4   Microsoft                       no-reply@microsoft.com                     98
    5   GitHub                          notifications@github.com                   76
  ...
---------------------------------------------------------------------------

RESUMEN:
  Correos analizados (totales):       3,700
  Remitentes unicos:                     43
  Top remitente:                 Computrabajo (487 correos, 13.2%)
```

### Con `--newsletters`

```
===========================================================================
  OUTLOOK EMAIL ANALYZER -- Ranking de remitentes
===========================================================================

  Carpeta: Bandeja de entrada | 3,700 correos (3,200 no leidos)
  Modo: todos + deteccion de newsletters -- analizando ~3,700 correos...

  Procesando correo 3,700/3,700... 100%

  Buscando links de desuscripcion en 43 remitentes...

  Escaneando remitentes... 28 links encontrados

───────────────────────────────────────────────────────────────────────────────
    #  Remitente                       Email                          Cant.  Unsub  Link de desuscripcion
───────────────────────────────────────────────────────────────────────────────
    1  Computrabajo                    alertas@computrabajo.com         487  SI     https://computrabajo.com/unsub?t=...
    2  LinkedIn Jobs                   jobs@linkedin.com                312  SI     https://linkedin.com/comm/unsubscribe
    3  Magneto Empleos                 noreply@magneto365.com           289  SI     https://magneto365.com/unsubscribe
    4  Microsoft                       no-reply@microsoft.com            98  --
    5  GitHub                          notifications@github.com          76  --
  ...
───────────────────────────────────────────────────────────────────────────────

RESUMEN:
  Correos analizados (totales):       3,700
  Remitentes unicos:                     43
  Top remitente:                 Computrabajo (487 correos, 13.2%)

  NEWSLETTERS:
    Remitentes con desuscripcion:     28 (65.1%)
    Correos de newsletters:        2,891 (78.1% del total)
    Remitentes sin desuscripcion:     15
```

---

## Seguridad

| Aspecto | Implementacion |
|---------|---------------|
| Almacenamiento de tokens | Cifrado con DPAPI en Windows via `msal-extensions`. En otros SO: permisos 0600. |
| Tokens en memoria | Encapsulados en `_SecureToken` que muestra `Token(***)` en tracebacks y logs. |
| Peticiones HTTP | Timeout de 10s (conexion) + 30s (lectura). Retry automatico para HTTP 429 con tope de 60s. |
| Exportacion CSV | Sanitizacion contra CSV Injection (formulas Excel: `=`, `+`, `-`, `@`). |
| Parametro --folder | Validacion con regex `^[a-zA-Z0-9_-]+$` para prevenir path traversal. |
| Errores de API | Solo se muestra `error.message` del JSON, nunca el body crudo. |
| Eliminacion de config | `--reset-auth` pide confirmacion antes de borrar. |
| .gitignore | Excluye `token_cache.json`, `config.json`, `*.csv`, `.env*`. |
| Permisos de la app | Solo solicita `Mail.Read` (lectura). No puede modificar, enviar ni borrar correos. |

**Nota:** El Client ID de Azure es un identificador publico (no es un secreto). Cada usuario
registra su propia app en Azure. Los tokens se almacenan exclusivamente en la maquina local
del usuario y nunca se transmiten fuera de las peticiones a Microsoft Graph.

---

## Estructura del proyecto

```
outlook-email-analyzer-python/
  main.py              Punto de entrada, CLI (argparse), orquestacion newsletters, exportacion CSV
  ms_graph.py          OAuth2 (MSAL + Device Code) + Graph API + token cache + newsletters (body HTML)
  analyzer.py          Logica de agrupacion (Counter), ranking por remitente, deteccion newsletters
  display.py           Formateo tabular: tabla normal, tabla newsletters con links, barras ASCII
  requirements.txt     Dependencias: msal, msal-extensions, requests
  .gitignore           Exclusiones para repositorio publico seguro
  README.md            Esta documentacion
```

| Archivo | Lineas | Responsabilidad |
|---------|--------|----------------|
| `ms_graph.py` | ~520 | Autenticacion, cache, API REST, validacion, deteccion de newsletters (2do pase HTML) |
| `main.py` | ~220 | CLI, flujo principal, orquestacion newsletters, exportacion CSV con sanitizacion |
| `analyzer.py` | ~80 | Logica pura de conteo, ranking y deteccion de newsletters por remitente |
| `display.py` | ~190 | Renderizado en consola: tabla normal, tabla newsletters, barras, resumen, progreso |

---

## Solucion de problemas

| Error | Causa | Solucion |
|-------|-------|----------|
| `Client ID invalido` | UUID pegado incorrectamente | Verifica que sea un UUID de 36 caracteres con guiones |
| `Error al iniciar autenticacion` | App no configurada como cliente publico | En Azure > tu app > Autenticacion > activar "Permitir flujos de clientes publicos" |
| `Token expirado o invalido` | Cache corrupto o token revocado | Ejecuta `python main.py --reset-auth` y vuelve a autenticar |
| `Carpeta no encontrada` | Nombre incorrecto de carpeta | Usa `python main.py --list-folders` para ver nombres validos |
| `API throttling` | Demasiadas peticiones a Microsoft | La app espera automaticamente y reintenta (hasta 3 veces) |
| `Falta msal` | Dependencia no instalada | `pip install -r requirements.txt` |
| `Timeout` | Conexion lenta o Microsoft no responde | Verifica tu conexion a internet e intenta de nuevo |

---

## Licencia

MIT
