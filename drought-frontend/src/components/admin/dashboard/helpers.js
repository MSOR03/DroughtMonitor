export function formatBytes(bytes) {
  if (!bytes) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${(bytes / Math.pow(k, i)).toFixed(1)} ${sizes[i]}`;
}

export function formatDate(dateStr) {
  if (!dateStr) return '—';
  return new Date(dateStr).toLocaleDateString('es-CO', {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });
}

export function classifyFile(filename) {
  const lower = (filename || '').toLowerCase();

  // Fuente por nombre (era5_land antes que era5 generico).
  const sourceLabel = (lower.includes('era5_land') || lower.includes('era5land'))
    ? 'ERA5-Land'
    : lower.includes('imerg') ? 'IMERG'
    : lower.includes('chirps') ? 'CHIRPS'
    : lower.includes('era5') ? 'ERA5'
    : null;

  // Prediccion PRIMERO: un archivo puede llamarse prediction_imerg_*.parquet y
  // no debe caer en la categoria hydromet solo por contener "imerg".
  const isPrediction = lower.includes('pred') || lower.includes('forecast')
    || lower.includes('horizonte') || lower.includes('horizon');
  if (isPrediction) {
    return {
      category: 'prediction',
      source: sourceLabel ? `Predicción ${sourceLabel}` : 'Predicción',
      color: 'purple',
    };
  }

  if (sourceLabel) {
    const color = sourceLabel === 'IMERG' ? 'sky' : sourceLabel === 'CHIRPS' ? 'cyan' : 'blue';
    return { category: 'hydromet', source: sourceLabel, color };
  }

  if (lower.includes('hidro') || lower.includes('hydro') || lower.includes('caudal') || lower.includes('nivel')) {
    return { category: 'hydrological', source: 'Hidrológico', color: 'teal' };
  }
  if (lower.includes('1m') || lower.includes('1mes')) return { category: 'prediction', source: 'Pred. 1 mes', color: 'purple' };
  if (lower.includes('3m') || lower.includes('3mes')) return { category: 'prediction', source: 'Pred. 3 meses', color: 'purple' };
  if (lower.includes('6m') || lower.includes('6mes')) return { category: 'prediction', source: 'Pred. 6 meses', color: 'purple' };
  if (lower.includes('12m') || lower.includes('12mes')) return { category: 'prediction', source: 'Pred. 12 meses', color: 'purple' };
  return { category: 'other', source: 'Otro', color: 'gray' };
}

/**
 * Infiere el dataset_key correcto a partir del nombre de archivo.
 * Devuelve una clave SOLO si existe en el catálogo y la fuente es inequívoca.
 * Ej.: "prediction_imerg_2026-07.parquet" -> "prediction_imerg".
 */
export function inferDatasetKeyFromFilename(filename, catalog = []) {
  const lower = (filename || '').toLowerCase();
  const availableKeys = new Set(catalog.map((d) => d.dataset_key));

  const source = (lower.includes('era5_land') || lower.includes('era5land')) ? 'ERA5_LAND'
    : lower.includes('imerg') ? 'IMERG'
    : lower.includes('chirps') ? 'CHIRPS'
    : lower.includes('era5') ? 'ERA5'
    : null;

  const isPrediction = lower.includes('pred') || lower.includes('forecast')
    || lower.includes('horizonte') || lower.includes('horizon');
  const isHydrological = lower.includes('hidro') || lower.includes('hydro')
    || lower.includes('caudal') || lower.includes('nivel');

  let candidate = null;
  if (isPrediction) {
    // Solo sugerimos si la fuente es clara (no adivinar entre CHIRPS/IMERG/ERA5-Land).
    candidate = source === 'IMERG' ? 'prediction_imerg'
      : source === 'ERA5_LAND' ? 'prediction_era5_land'
      : source === 'CHIRPS' ? 'prediction_chirps'
      : null;
  } else if (isHydrological) {
    candidate = 'hydro_main';
  } else if (source) {
    candidate = source === 'ERA5' ? 'historical_era5'
      : source === 'ERA5_LAND' ? 'historical_era5_land'
      : source === 'IMERG' ? 'historical_imerg'
      : source === 'CHIRPS' ? 'historical_chirps'
      : null;
  }

  return candidate && availableKeys.has(candidate) ? candidate : null;
}
