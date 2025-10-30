const express = require("express");
const XLSX = require("xlsx");
const cors = require("cors");
const path = require("path");

const app = express();

// ✅ Render asigna su propio puerto dinámico
const PORT = process.env.PORT || 3000;

app.use(cors());

// 📂 Ruta al archivo Excel
const FILE_PATH = path.join(__dirname, "documentos", "ReporteDetalleVentasPorHora2025-10-22.xlsx");

// 🧠 Función auxiliar para normalizar texto
const normalize = (text = "") =>
  text.toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();

// 🧾 Función para leer y limpiar datos del Excel
function leerVentas() {
  const workbook = XLSX.readFile(FILE_PATH);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: "", range: 5 });

  const ventas = jsonData
    .filter(r => r["Platillo/Artículo"])
    .map(r => ({
      Platillo: r["Platillo/Artículo"],
      Grupo: r["Grupo"] || "",
      Fecha: r["Fecha"] || "",
      Dia: r["Día"] || "",
      Hora: r["Hora"] || "",
      Cantidad: Number(r["Cantidad"] || 0),
      Subtotal: Number(r["Subtotal"] || 0),
      IVA: Number(r["IVA"] || 0),
      Total: Number(r["Total"] || 0),
      Porcentaje: r["%"] || ""
    }));

  return ventas;
}

// 👋 Ruta raíz obligatoria para Render (salud del servicio)
app.get("/", (req, res) => {
  res.send("🚀 API de Ventas corriendo correctamente en Render");
});

// 📊 Endpoint principal con filtros y resumen opcional
app.get("/api/ventas", (req, res) => {
  try {
    const ventas = leerVentas();

    // Filtros dinámicos
    const { platillo, dia, fecha, resumen } = req.query;
    let resultado = ventas;

    if (platillo) {
      resultado = resultado.filter(v =>
        normalize(v.Platillo).includes(normalize(platillo))
      );
    }

    if (dia) {
      resultado = resultado.filter(v =>
        normalize(v.Dia).includes(normalize(dia))
      );
    }

    if (fecha) {
      resultado = resultado.filter(v => v.Fecha === fecha);
    }

    // Si se pide resumen (?resumen=true)
    if (resumen === "true") {
      const totalUnidades = resultado.reduce((acc, cur) => acc + (cur.Cantidad || 0), 0);
      const totalVentas = resultado.reduce((acc, cur) => acc + (cur.Total || 0), 0);

      // Platillo más vendido
      const agrupado = {};
      resultado.forEach(v => {
        agrupado[v.Platillo] = (agrupado[v.Platillo] || 0) + v.Cantidad;
      });

      const top = Object.entries(agrupado).sort((a, b) => b[1] - a[1])[0];

      const resumenData = {
        total_unidades: totalUnidades,
        total_ventas: totalVentas.toFixed(2),
        platillo_mas_vendido: top ? top[0] : "N/D",
        unidades_vendidas: top ? top[1] : 0
      };

      return res.json(resumenData);
    }

    res.json(resultado);
  } catch (error) {
    console.error("Error leyendo el archivo:", error);
    res.status(500).json({ error: "Error leyendo el archivo Excel" });
  }
});

// 📈 Endpoint separado solo para resumen global
app.get("/api/resumen", (req, res) => {
  try {
    const ventas = leerVentas();

    const totalUnidades = ventas.reduce((acc, cur) => acc + (cur.Cantidad || 0), 0);
    const totalVentas = ventas.reduce((acc, cur) => acc + (cur.Total || 0), 0);

    const agrupado = {};
    ventas.forEach(v => {
      agrupado[v.Platillo] = (agrupado[v.Platillo] || 0) + v.Cantidad;
    });

    const top = Object.entries(agrupado).sort((a, b) => b[1] - a[1])[0];

    const resumen = {
      total_unidades: totalUnidades,
      total_ventas: totalVentas.toFixed(2),
      platillo_mas_vendido: top ? top[0] : "N/D",
      unidades_vendidas: top ? top[1] : 0
    };

    res.json(resumen);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Error generando resumen" });
  }
});

// 🚀 Inicia el servidor
app.listen(PORT, () => {
  console.log(`✅ API corriendo en el puerto ${PORT}`);
});
