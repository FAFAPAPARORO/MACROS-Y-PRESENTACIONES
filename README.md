# Macros y presentaciones

Generadores de presentaciones financieras con el formato gráfico ACONTIS
(azul marino / turquesa, tarjetas de indicadores, tablas con franjas y
separadores con fotografía).

## Requisitos

```bash
npm install
```

## Presentaciones

| Comando | Salida |
|---|---|
| `npm run juridica` | `presentaciones/juridica/Analisis_Financiero_Juridica_Junio_2026.pptx` |

Todas las cifras, textos, tablas y gráficos de cada presentación están en el
bloque `DATOS` al inicio de su generador. Para actualizar un mes, se cambian
esos valores y se vuelve a ejecutar el comando.

El `.pptx` resultante es 100 % editable: los gráficos son nativos de
PowerPoint (clic derecho → *Editar datos*), las tablas son tablas reales y
cada texto es un cuadro de texto independiente.

Los recursos gráficos compartidos (fondo y logos ACONTIS) están en `assets/acontis/`.
