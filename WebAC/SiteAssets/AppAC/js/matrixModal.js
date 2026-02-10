(function () {
  const matrix = {
    1: [5, 10, 20, 35, 55],
    2: [15, 25, 40, 60, 80],
    3: [30, 45, 65, 85, 100],
    4: [50, 70, 90, 105, 115],
    5: [75, 95, 110, 120, 125]
  };

  function clampInt(value, min, max) {
    const v = Number(value);
    if (!Number.isFinite(v)) return min;
    return Math.min(max, Math.max(min, Math.round(v)));
  }

  function buildMatrixSvg(severity, ff) {
    const sev = clampInt(severity, 1, 5);
    const freq = clampInt(ff, 1, 5);
    const colIndex = freq - 1;
    const rowIndex = 5 - sev;
    const cellSize = 60;
    const gridX = 190;
    const gridY = 110;
    const cellX = gridX + colIndex * cellSize;
    const cellY = gridY + rowIndex * cellSize;
    const cx = cellX + cellSize / 2;
    const cy = cellY + cellSize / 2;
    const nc = (matrix[sev] || matrix[1])[colIndex] ?? "-";

    const rows = [
      { sev: 5, colors: ["#ef4444", "#ef4444", "#ef4444", "#ef4444", "#ef4444"], values: matrix[5] },
      { sev: 4, colors: ["#FFFF00", "#FFFF00", "#ef4444", "#ef4444", "#ef4444"], values: matrix[4] },
      { sev: 3, colors: ["#22c55e", "#FFFF00", "#FFFF00", "#ef4444", "#ef4444"], values: matrix[3] },
      { sev: 2, colors: ["#22c55e", "#22c55e", "#FFFF00", "#FFFF00", "#ef4444"], values: matrix[2] },
      { sev: 1, colors: ["#22c55e", "#22c55e", "#22c55e", "#FFFF00", "#FFFF00"], values: matrix[1] }
    ];

    const cells = rows.map((row, r) => {
      return row.values.map((value, c) => {
        const x = gridX + c * cellSize;
        const y = gridY + r * cellSize;
        const color = row.colors[c];
        return `
    <rect x="${x}" y="${y}" width="${cellSize}" height="${cellSize}" fill="${color}" />
    <text x="${x + cellSize / 2}" y="${y + cellSize / 2}" text-anchor="middle" dominant-baseline="middle">${value}</text>`;
      }).join("\n");
    }).join("\n");

    const colLabels = [1, 2, 3, 4, 5].map((v, i) => {
      const x = gridX + cellSize / 2 + i * cellSize;
      return `<text x="${x}" y="90" text-anchor="middle">${v}</text>`;
    }).join("\n");

    const rowLabels = [5, 4, 3, 2, 1].map((v, i) => {
      const y = gridY + cellSize / 2 + i * cellSize;
      return `<text x="130" y="${y}" text-anchor="middle">${v}</text>`;
    }).join("\n");

    return `
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 460" role="img" aria-label="Matriz de criticidad">
  <rect x="20" y="20" width="600" height="410" rx="14" fill="#ffffff" stroke="#e5e7eb" />
  <text x="320" y="52" text-anchor="middle" font-size="14" fill="#0f172a" font-weight="700">FACTOR DE FRECUENCIA DE FALLA (FF)</text>
  <text x="34" y="235" text-anchor="middle" font-size="12" fill="#0f172a" font-weight="700" transform="rotate(-90 34 235)">NIVEL DE SEVERIDAD</text>

  <g font-size="12" fill="#0f172a" font-weight="700">
    ${colLabels}
  </g>

  <g font-size="12" fill="#0f172a" font-weight="700">
    ${rowLabels}
  </g>

  <g font-size="12" fill="#0f172a" font-weight="700">
    ${cells}
  </g>

  <line x1="${cx}" y1="100" x2="${cx}" y2="${gridY + cellSize * 5}" stroke="#777373" stroke-width="1.5" stroke-dasharray="4 6" opacity="0.9" />
  <line x1="170" y1="${cy}" x2="${gridX + cellSize * 5 + 10}" y2="${cy}" stroke="#777373" stroke-width="1.5" stroke-dasharray="4 6" opacity="0.9" />
  <circle cx="${cx}" cy="${cy}" r="18" fill="none" stroke="#000000" stroke-width="1.5" />

  <text x="320" y="435" text-anchor="middle" font-size="13" fill="#0f172a" font-weight="700">FF ${freq} · Severidad ${sev} → NC ${nc}</text>
</svg>`;
  }

  window.openCriticidadMatrixDynamic = function ({ severity, ff }) {
    const html = buildMatrixSvg(severity, ff);
    if (window.Swal) {
      Swal.fire({
        title: "Matriz de criticidad",
        html,
        showConfirmButton: true,
        confirmButtonText: "Cerrar",
        width: 720,
        padding: "1rem"
      });
      return;
    }
  };
})();
