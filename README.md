# VINCI Expedition — Sistema de votación

Aplicación de votación para la jornada VINCI Expedition, construida sobre Google
Apps Script. Cada persona vota **un país por categoría**; los resultados se
ponderan 80/20 entre una lista de votantes designados y el resto.

Corre como web app de Apps Script con una hoja de cálculo de Google como base de
datos. No hay build ni dependencias: los archivos se copian tal cual al proyecto.

## Archivos

| Archivo | Qué es |
|---|---|
| `Code.gs` | Todo el backend: votación, resultados y administración |
| `Index.html` | La aplicación completa — markup, estilos y lógica de cliente |
| `appsscript.json` | Manifiesto: scopes y configuración del despliegue |
| `LoadTest.gs` | Pruebas de carga. **No es parte del sistema — borrar antes del evento** |

## Cómo funciona

### Categorías

Cada país compite en una sola categoría, definida en `CATEGORIAS` (`Code.gs`):

| Categoría | Países |
|---|---|
| Mejora Continua | Argentina, CIB, Perú |
| IA / Modelos | Colombia, Holding, Venezuela |
| Proyecto SDA | España, Uruguay |

Un país que no figura en ninguna categoría **no se vota**: hoy Chile, México y
Turquía. Igual pueden votar como votantes. Para incorporar uno, agregalo a la
lista de su categoría **en `Code.gs` y en `Index.html`** — la definición está
duplicada a propósito para que el cliente arme la pantalla sin una llamada más
al servidor. El backend valida siempre, así que la copia del cliente no puede
habilitar un voto inválido.

Cada tarjeta muestra la bandera, el nombre del país y el título de su
iniciativa. Los títulos viven en `PAISES_INFO` (`Index.html`), campo
`iniciativa`, y **hoy son un texto provisorio**.

Nadie puede votar por su propio país. Si eso deja una categoría sin opciones,
no se le exige elegir ahí.

### Hojas de cálculo

La planilla se resuelve por la propiedad `SS_ID` en las Propiedades del script.

| Hoja | Contenido |
|---|---|
| `Votos` | `Timestamp · Email · PaisVotante · EsPonderado · una columna por categoría` |
| `Lista Ponderada` | `Email · Nombre · Pais · Cargo` — sólo se usa la columna Email |

Los tres votos de una persona van en **una sola fila**. Es lo que mantiene el
registro de un voto en una única escritura, que es de lo que depende el ritmo
cuando entra mucha gente a la vez.

### Ponderación

`PESO_PONDERADO` (0.80) y `PESO_REGULAR` (0.20) se reparten **dentro de cada
categoría**: el denominador es la gente que votó en esa categoría, no el total
de votantes. Por eso los scores no son comparables entre categorías.

El peso es del *bloque*, no de cada persona: cuantos más votantes ponderados
haya, menos pesa cada uno.

## Puesta en marcha

1. Crear el proyecto en [script.google.com](https://script.google.com) y copiar
   `Code.gs`, `Index.html` y `appsscript.json` (activar "Mostrar appsscript.json"
   en Configuración del proyecto).
2. Ejecutar `inicializarSistema()` una vez. Loguea la URL de la planilla.
3. Cargar la hoja `Lista Ponderada` y **borrar la fila de ejemplo**.
4. Ejecutar `refrescarCaches()`.
5. **Implementar → Nueva implementación → Aplicación web**, ejecutar como el
   propietario, acceso: el dominio.
6. Abrir el panel admin → pestaña **Acceso / QR** para el código QR.

## Funciones del editor

| Función | Para qué |
|---|---|
| `inicializarSistema()` | Crea las hojas. Una sola vez |
| `refrescarCaches()` | Vacía todos los cachés |
| `simularVotos(n)` | Prueba de carga (`LoadTest.gs`) |
| `limpiarVotosDePrueba()` | Borra las filas que dejó la prueba |

`refrescarCaches()` hay que correrlo **después de editar a mano** `Lista
Ponderada`: si no, los cambios tardan hasta 25 minutos en verse. Es el paso que
más fácil se olvida, y su síntoma —gente ponderada contando como regular— no es
evidente.

## Rendimiento

El sistema está afinado para un pico de gente votando a la vez:

- El lock de escritura envuelve **sólo** el `appendRow`. La verificación de
  duplicado y la lista de ponderados se resuelven fuera, contra caché.
- Como esa verificación no es autoritativa, el conteo se queda con **la primera
  fila de cada email** y descarta duplicados. Nadie vota dos veces aunque se
  cuelen dos filas.
- Un voto exitoso **suma su email al caché** en vez de invalidarlo: invalidar
  obligaría al siguiente votante a releer la hoja entera.
- La lista de ponderados está cacheada; sin eso se leía en cada validación y
  en cada voto.
- El email detectado se incrusta en la página desde `doGet`: una visita cuesta
  **una** ejecución de Apps Script en vez de varias.

Medir antes de asumir: `simularVotos(100)` reporta latencia p50/p95/max y votos
por segundo. Como el lock serializa, el throughput del sistema es 1 dividido la
latencia, así que ese número es el techo real.

## Configuración

En `CONFIG` (`Code.gs`):

| Clave | Qué controla |
|---|---|
| `ADMIN_PASSWORD` | Acceso al panel de administración |
| `PESO_PONDERADO` / `PESO_REGULAR` | Reparto 80/20 |
| `CACHE_*` | TTL de cada caché, en segundos |

> La contraseña del panel está en el código. Si este repositorio deja de ser
> privado, moverla a las Propiedades del script antes.

## Pruebas de carga

`LoadTest.gs` es autocontenido: borrar el archivo elimina todo el aparato de
testing, incluido su endpoint `doPost`.

```
simularVotos(100)        // mide latencia y votos por segundo
limpiarVotosDePrueba()   // borra sólo las filas con prefijo loadtest-
verResumenPrueba()       // reimprime el último resultado
```

Para concurrencia real hay `simularVotosConcurrentes(n)`, que dispara en
paralelo contra el web app publicado. Requiere configurar `LOADTEST.TOKEN` y
`LOADTEST.WEBAPP_URL` y volver a desplegar. **Dejar el token vacío para el
evento**: con el token vacío el endpoint rechaza todo.
