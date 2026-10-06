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
| `npm run juridica` | `presentaciones/juridica/Analisis_Financiero_Juridica_Agosto_2026.pptx` (julio vs agosto 2026) |

Cada generador tiene al inicio un bloque `DATOS` con las cifras en pesos
(estado de resultados mes a mes, balance de los dos meses comparados y detalle
de gastos) y un bloque `TEXTOS` con la redacción del análisis. Tablas, tarjetas,
gráficos, márgenes, indicadores y el puente de utilidad se calculan solos a partir
de `DATOS`. Para actualizar un mes se agregan las cifras nuevas, se ajustan los
textos y se vuelve a ejecutar el comando.

El `.pptx` resultante es 100 % editable: los gráficos son nativos de
PowerPoint (clic derecho → *Editar datos*), las tablas son tablas reales y
cada texto es un cuadro de texto independiente.

Los recursos gráficos compartidos (fondo y logos ACONTIS) están en `assets/acontis/`.
