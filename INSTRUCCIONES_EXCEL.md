# INSTRUCCIONES PARA INICIAR EL PROGRAMA EN EXCEL

## ANÁLISIS ECONÓMICO BOMBA DE CALOR TRANSCRÍTICA CO2 - PROYECTO BAXTER

### 🚀 PASOS PARA INICIALIZAR EL SISTEMA

#### **1. PREPARACIÓN INICIAL**
- Abrir Microsoft Excel (versión 2016 o superior recomendada)
- Habilitar macros cuando Excel lo solicite
- Crear un nuevo libro de trabajo (.xlsm)

#### **2. IMPORTAR LOS MÓDULOS VBA**

**Acceder al Editor VBA:**
1. Presionar `Alt + F11` para abrir el Editor de Visual Basic
2. Ir a `Archivo > Importar archivo` o presionar `Ctrl + M`

**Importar los 5 módulos en orden:**
1. **Modulo1** - Configuración inicial y Dashboard
2. **Modulo2** - Análisis CAPEX
3. **Modulo3** - Análisis OPEX  
4. **Modulo4** - Flujo de caja y métricas financieras
5. **Modulo5** - Análisis de sensibilidad

#### **3. EJECUTAR LA CONFIGURACIÓN INICIAL**

**En el Editor VBA:**
1. Presionar `F5` o ir a `Ejecutar > Ejecutar Sub/UserForm`
2. Ejecutar la macro: `ConfigurarLibroCompleto()`

**O desde Excel:**
1. Presionar `Alt + F8`
2. Seleccionar `ConfigurarLibroCompleto`
3. Hacer clic en `Ejecutar`

#### **4. VERIFICACIÓN DE CONFIGURACIÓN EXITOSA**

✅ **El sistema debe crear automáticamente:**
- Hoja **Dashboard** - Panel ejecutivo principal
- Hoja **CAPEX** - Análisis de inversión inicial
- Hoja **OPEX** - Comparativo operacional vapor vs bomba calor
- Hoja **Flujo_Caja** - Análisis financiero 15 años
- Hoja **Sensibilidad** - Análisis de variables críticas
- Hoja **Datos_Base** - Parámetros editables

✅ **Mensaje de confirmación:** "¡Sistema Excel BAXTER configurado exitosamente!"

### 📊 NAVEGACIÓN DEL SISTEMA

#### **DASHBOARD PRINCIPAL**
- **Ubicación:** Primera hoja que se activa automáticamente
- **Función:** Resumen ejecutivo con métricas clave
- **Navegación:** Enlaces directos a todas las hojas

#### **DATOS EDITABLES (Amarillo claro)**
- **Datos_Base:** Precios energía, tasas financieras, parámetros técnicos
- **CAPEX:** Cantidades y costos unitarios de equipos
- **OPEX:** Costos operacionales sistema vapor vs bomba calor

#### **DATOS CALCULADOS (Azul claro)**
- Se actualizan automáticamente al cambiar parámetros base
- VPN, TIR, Payback, Índice de Rentabilidad

### ⚙️ PERSONALIZACIÓN DE PARÁMETROS

#### **Parámetros Energéticos (Hoja Datos_Base):**
- Precio Gas Natural: $0.041/kWh (editable)
- Precio Electricidad: $0.113/kWh (editable)
- Escalación Anual: 5% (editable)

#### **Parámetros Financieros:**
- Tasa Descuento (WACC): 8% (editable)
- Valor Residual: 15% del CAPEX (editable)
- Horizonte: 15 años (fijo)

#### **Parámetros Técnicos:**
- COP Sistema: 3.0 (editable)
- Horas Operación: 8,760/año (editable)

### 🔧 RESOLUCIÓN DE PROBLEMAS

#### **Si aparecen errores #¿VALOR?:**
1. Verificar que todos los módulos estén importados
2. Ejecutar nuevamente `ConfigurarLibroCompleto()`
3. Revisar que las hojas no hayan sido renombradas

#### **Si las macros no funcionan:**
1. Verificar configuración de seguridad de Excel
2. Ir a `Archivo > Opciones > Centro de confianza`
3. Habilitar macros para este archivo

#### **Si faltan nombres de rango:**
1. Ir a `Fórmulas > Administrador de nombres`
2. Verificar existencia de: CAPEX_Total, OPEX_AhorroAnual, VPN_Total, TIR_Proyecto

### 📈 INTERPRETACIÓN DE RESULTADOS

#### **Métricas Clave Dashboard:**
- **VPN > 0:** Proyecto viable financieramente
- **TIR > 8%:** Supera costo de capital (WACC)
- **Payback < 10 años:** Período de recuperación aceptable
- **Índice Rentabilidad > 1:** Genera valor para la empresa

#### **Indicadores Visuales:**
- ✅ **Verde:** Métricas favorables
- ❌ **Rojo:** Métricas desfavorables
- 🟡 **Amarillo:** Celdas editables por usuario

### 📞 SOPORTE TÉCNICO

**Archivo de prueba:** Ejecutar `PruebaCompilacion.vba` para verificar sintaxis
**Validación OPEX:** Ejecutar `ValidarDatosOPEX()` para comprobar coherencia de datos

---
**Versión:** W2502E3PRO  
**Fecha:** Julio 2025  
**Proyecto:** BAXTER S.A. - Planta Cali, Colombia